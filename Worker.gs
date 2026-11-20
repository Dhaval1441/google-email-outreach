/**
 * One timer entry point. Each slot makes at most one Gmail call.
 * Order: recover a crashed send, confirm one reply, read one inbox page, then send one email.
 */

function processQueue() {
  runWorker_({ ignoreWindow: false, slotCount: 0, ignorePause: false });
}

function runOneSlotNow() {
  var result = runWorker_({ ignoreWindow: true, slotCount: 1, ignorePause: false });
  SpreadsheetApp.getUi().alert(result.summary || 'Finished one slot.');
}

function runWorker_(options) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    return { summary: 'Another slot is already running.' };
  }
  var summary = 'Nothing due.';
  try {
    var settings = getSettings();
    if (!options.ignorePause && isPaused(settings)) {
      return { summary: 'Outreach is paused.' };
    }
    resetSenderQuotas_(settings, loadSenders_());
    var slots = options.slotCount || settings.slotSize || 1;
    var started = Date.now();
    for (var i = 0; i < slots; i++) {
      if (Date.now() - started > 4 * 60 * 1000) {
        summary = 'Stopped early to stay inside the execution limit.';
        break;
      }
      var result = runOneSlot_();
      summary = result.summary;
      if (result.stop) break;
      if (i < slots - 1 && result.usedGmail && settings.gapSeconds > 0) {
        Utilities.sleep(settings.gapSeconds * 1000);
      }
    }
  } catch (err) {
    summary = err.message || String(err);
    try {
      appendLog_('', 'Error', '', summary);
    } catch (ignore) {}
  } finally {
    lock.releaseLock();
  }
  return { summary: summary };
}

function runOneSlot_() {
  if (globalCooldownActive_(new Date())) {
    return slotResult_(true, false, 'Waiting after a Gmail rate limit.');
  }
  var ctx = loadContext_();
  maintainLeads_(ctx);

  var recovery = findRecoveryLead_(ctx);
  if (recovery) return recoverLead_(ctx, recovery);

  var queued = findQueuedReply_(ctx);
  if (queued) return confirmReply_(ctx, queued);

  var unfinished = findUnfinishedSender_(ctx);
  if (unfinished) return syncOnePage_(ctx, unfinished, false);

  var prep = prepTarget_(ctx);
  if (prep && prep.urgent) return syncOnePage_(ctx, prep.sender, true);

  var candidate = selectCandidate_(ctx);
  if (candidate) {
    if (candidate.kind === 'followup' && !followUpSyncIsFresh_(candidate.sender, ctx.now)) {
      return syncOnePage_(ctx, candidate.sender, true);
    }
    var system = systemFor_(ctx, candidate.lead.leadId);
    if (candidate.kind === 'followup' && !system.lastMessageId && system.gmailApiMessageId) {
      return fetchRfcMessageId_(ctx, candidate);
    }
    return sendCandidate_(ctx, candidate);
  }

  if (prep) return syncOnePage_(ctx, prep.sender, true);

  if (ctx.skippedForHours) {
    return slotResult_(true, false, 'Nothing to send. Reply checks run in the hours before the next local morning.');
  }
  if (ctx.skippedForCountry) {
    return slotResult_(true, false, 'Add a known country on the lead before it can send.');
  }
  return slotResult_(true, false, 'Nothing due.');
}

function slotResult_(stop, usedGmail, summary) {
  return { stop: stop, usedGmail: usedGmail, summary: summary };
}

function maintainLeads_(ctx) {
  var used = {};
  ctx.leads.forEach(function (lead) {
    if (lead.leadId) used[lead.leadId] = true;
  });
  var nextNumber = 1;
  var assigned = false;
  ctx.leads.forEach(function (lead) {
    if (lead.leadId) return;
    var id = '';
    do {
      id = 'L' + padNumber_(nextNumber);
      nextNumber++;
    } while (used[id]);
    lead.leadId = id;
    used[id] = true;
    assigned = true;
  });
  if (assigned) {
    writeLeadColumn_('Lead ID', ctx.leads, function (lead) {
      return lead.leadId;
    });
    appendLog_('', 'Lead ID', '', 'Assigned missing lead ids');
  }

  var seen = {};
  var duplicates = false;
  ctx.leads.forEach(function (lead) {
    var email = String(lead.email || '').trim().toLowerCase();
    if (!email) return;
    if (!seen[email]) {
      if (lead.status !== 'Duplicate') seen[email] = true;
      return;
    }
    var canMark = !lead.status || lead.status === 'Pending' || (lead.status === 'Error' && !lead.threadId);
    if (!canMark || lead.threadId) return;
    lead.status = 'Duplicate';
    duplicates = true;
  });
  if (duplicates) {
    writeLeadColumn_('Status', ctx.leads, function (lead) {
      return lead.status;
    });
    appendLog_('', 'Duplicate', '', 'Marked later rows that repeat an email address');
  }
}

function findRecoveryLead_(ctx) {
  var found = null;
  ctx.leads.forEach(function (lead) {
    var system = systemFor_(ctx, lead.leadId);
    if (!system.needsRecovery) return;
    var email = lead.senderEmail || system.senderEmail;
    if (email && inCooldown_(email, ctx.now)) return;
    if (!found || lead.rowNumber < found.rowNumber) found = lead;
  });
  return found;
}

function findQueuedReply_(ctx) {
  var found = null;
  ctx.inboxRows.forEach(function (row) {
    if (row.state !== 'queued') return;
    if (inCooldown_(row.senderEmail, ctx.now)) return;
    var sender = ctx.senderByEmail[String(row.senderEmail || '').toLowerCase()];
    if (sender && String(sender.authStatus).toLowerCase() === 'needs auth') return;
    if (!found || row.rowNumber < found.rowNumber) found = row;
  });
  return found;
}

function findUnfinishedSender_(ctx) {
  for (var i = 0; i < ctx.senders.length; i++) {
    var sender = ctx.senders[i];
    if (!sender.inboxPageToken || !sender.active) continue;
    if (inCooldown_(sender.email, ctx.now)) continue;
    if (String(sender.authStatus).toLowerCase() === 'needs auth') continue;
    return sender;
  }
  return null;
}

function leadSendWindow_(ctx, lead) {
  var zone = timezoneForCountry_(lead.country);
  if (!zone) return { ok: false, reason: 'country' };
  if (!withinSendWindow(ctx.settings, ctx.now, zone)) return { ok: false, reason: 'hours', timeZone: zone };
  return { ok: true, timeZone: zone };
}

function rememberCountryIssue_(ctx, lead) {
  if (!lead.leadId) return;
  var system = systemFor_(ctx, lead.leadId);
  var message = 'Country is blank or not recognized. Use a European country, Dubai or UAE, or USA.';
  if (system.lastError === message) return;
  system.lastError = message;
  saveSystem_(system);
  appendLog_(lead.leadId, 'Country', '', message);
}

function selectCandidate_(ctx) {
  var followUps = [];
  var pendings = [];
  ctx.skippedForHours = 0;
  ctx.skippedForCountry = 0;
  ctx.leads.forEach(function (lead) {
    if (FINISHED_STATUSES[lead.status]) return;
    var system = systemFor_(ctx, lead.leadId);
    if (system.needsRecovery) return;
    if (lead.status === 'Error' && lead.followUpDate && lead.followUpDate.getTime() > ctx.now.getTime()) return;
    var status = normalizedStatus_(lead, system);
    if (status === 'Pending') {
      if (lead.followUpDate && lead.followUpDate.getTime() > ctx.now.getTime()) return;
      pendings.push(lead);
      return;
    }
    if (FOLLOW_UP_STATUSES[status]) {
      if (!lead.threadId || !lead.senderEmail) return;
      if (!lead.followUpDate || lead.followUpDate.getTime() <= ctx.now.getTime()) followUps.push(lead);
    }
  });
  followUps.sort(function (a, b) {
    var aTime = a.followUpDate ? a.followUpDate.getTime() : 0;
    var bTime = b.followUpDate ? b.followUpDate.getTime() : 0;
    return aTime - bTime || a.rowNumber - b.rowNumber;
  });
  pendings.sort(function (a, b) {
    return a.rowNumber - b.rowNumber;
  });
  var ordered = followUps.concat(pendings);
  for (var i = 0; i < ordered.length; i++) {
    var lead = ordered[i];
    var system = systemFor_(ctx, lead.leadId);
    var stage = outgoingStage_(normalizedStatus_(lead, system));
    if (!stage) continue;
    if (!isEmailAddress_(lead.email)) {
      lead.status = 'Bounced';
      lead.followUpDate = null;
      saveLead_(lead);
      appendLog_(lead.leadId, 'Bounced', lead.senderEmail, 'Invalid email address');
      continue;
    }
    var local = leadSendWindow_(ctx, lead);
    if (!local.ok) {
      if (local.reason === 'country') {
        ctx.skippedForCountry++;
        if (ctx.skippedForCountry <= 10) rememberCountryIssue_(ctx, lead);
      } else {
        ctx.skippedForHours++;
      }
      continue;
    }
    if (stage === 'initial') {
      var sender = chooseSender_(ctx, false);
      if (!sender) continue;
      return { lead: lead, sender: sender, stage: stage, kind: 'initial' };
    }
    var followSender = ctx.senderByEmail[lead.senderEmail.toLowerCase()];
    if (!followSender || !senderCanSend_(followSender, ctx, true)) continue;
    return { lead: lead, sender: followSender, stage: stage, kind: 'followup' };
  }
  return null;
}

function outgoingStage_(status) {
  if (status === 'Pending') return 'initial';
  if (status === 'Sent') return 'fu1';
  if (status === 'Follow-up 1') return 'fu2';
  if (status === 'Follow-up 2') return 'fu3';
  return '';
}

function chooseSender_(ctx, forFollowUp) {
  var best = null;
  ctx.senders.forEach(function (sender) {
    if (!senderCanSend_(sender, ctx, forFollowUp)) return;
    if (!best || sender.sentToday < best.sentToday) {
      best = sender;
      return;
    }
    if (sender.sentToday === best.sentToday) {
      var senderTime = sender.lastUsedAt ? sender.lastUsedAt.getTime() : 0;
      var bestTime = best.lastUsedAt ? best.lastUsedAt.getTime() : 0;
      if (senderTime < bestTime) best = sender;
    }
  });
  return best;
}

function senderCanSend_(sender, ctx, forFollowUp) {
  if (!sender || !sender.active) return false;
  if (String(sender.authStatus).toLowerCase() !== 'authorized') return false;
  if (inCooldown_(sender.email, ctx.now)) return false;
  var limit = Number(sender.dailyLimit) || Number(ctx.settings.dailyLimitDefault) || 30;
  var sent = Number(sender.sentToday) || 0;
  if (forFollowUp) return sent < limit;
  var reserve = Number(ctx.settings.followUpReserve) || 0;
  return sent < Math.max(0, limit - reserve);
}

function senderLimit_(sender, settings) {
  return Number(sender.dailyLimit) || Number(settings.dailyLimitDefault) || 30;
}

function followUpSyncIsFresh_(sender, now) {
  if (!sender || sender.inboxPageToken || !sender.syncFinishedAt) return false;
  return now.getTime() - sender.syncFinishedAt.getTime() < 3 * 60 * 60 * 1000;
}

function pad2_(value) {
  var text = String(value);
  return text.length < 2 ? '0' + text : text;
}

function parseCivil_(text) {
  var match = /^(\d{4})-(\d{2})-(\d{2}) (\d{2}):(\d{2})$/.exec(text);
  if (!match) return 0;
  return Date.UTC(Number(match[1]), Number(match[2]) - 1, Number(match[3]), Number(match[4]), Number(match[5]), 0);
}

function civilTime_(timeZone, ymd, minutesOfDay) {
  var parts = ymd.split('-');
  var hours = Math.floor(minutesOfDay / 60);
  var mins = minutesOfDay % 60;
  var guess = new Date(Date.UTC(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]), hours, mins, 0));
  var wanted = ymd + ' ' + pad2_(hours) + ':' + pad2_(mins);
  var corrected = guess;
  for (var pass = 0; pass < 2; pass++) {
    var local = Utilities.formatDate(corrected, timeZone, 'yyyy-MM-dd HH:mm');
    var delta = parseCivil_(local) - parseCivil_(wanted);
    corrected = new Date(corrected.getTime() - delta);
  }
  return corrected;
}

function localWeekday_(timeZone, when) {
  return Number(Utilities.formatDate(when, timeZone, 'u'));
}

function nextWindowOpen_(settings, timeZone, now) {
  if (withinSendWindow(settings, now, timeZone)) return new Date(now.getTime());
  for (var addDays = 0; addDays < 8; addDays++) {
    var probe = new Date(now.getTime() + addDays * 24 * 60 * 60 * 1000);
    var ymd = Utilities.formatDate(probe, timeZone, 'yyyy-MM-dd');
    var open = civilTime_(timeZone, ymd, settings.windowStartMin);
    if (open.getTime() <= now.getTime()) continue;
    var weekday = localWeekday_(timeZone, open);
    if (!settings.sendOnWeekends && (weekday === 6 || weekday === 7)) continue;
    return open;
  }
  return new Date(now.getTime() + 24 * 60 * 60 * 1000);
}

function closedPeriodStart_(settings, timeZone, now) {
  for (var back = 0; back < 6; back++) {
    var probe = new Date(now.getTime() - back * 24 * 60 * 60 * 1000);
    var ymd = Utilities.formatDate(probe, timeZone, 'yyyy-MM-dd');
    var closeAt = civilTime_(timeZone, ymd, settings.windowEndMin);
    if (closeAt.getTime() > now.getTime()) continue;
    var weekday = localWeekday_(timeZone, closeAt);
    if (!settings.sendOnWeekends && (weekday === 6 || weekday === 7)) continue;
    return closeAt;
  }
  return new Date(now.getTime() - 24 * 60 * 60 * 1000);
}

function prepTarget_(ctx) {
  var best = null;
  var lookahead = 15 * 60 * 60 * 1000;
  var finalStretch = 3 * 60 * 60 * 1000;
  ctx.senders.forEach(function (sender) {
    if (!sender.active || sender.inboxPageToken) return;
    if (inCooldown_(sender.email, ctx.now)) return;
    if (String(sender.authStatus).toLowerCase() === 'needs auth') return;
    var key = sender.email.toLowerCase();
    var soonest = 0;
    var soonestZone = '';
    ctx.leads.forEach(function (lead) {
      if (String(lead.senderEmail || '').toLowerCase() !== key) return;
      if (FINISHED_STATUSES[lead.status] || !lead.threadId || !lead.followUpDate) return;
      var system = systemFor_(ctx, lead.leadId);
      if (system.needsRecovery) return;
      var status = normalizedStatus_(lead, system);
      if (!FOLLOW_UP_STATUSES[status]) return;
      var zone = timezoneForCountry_(lead.country);
      if (!zone || withinSendWindow(ctx.settings, ctx.now, zone)) return;
      var open = nextWindowOpen_(ctx.settings, zone, ctx.now);
      if (lead.followUpDate.getTime() > open.getTime()) return;
      if (open.getTime() - ctx.now.getTime() > lookahead) return;
      if (!soonest || open.getTime() < soonest) {
        soonest = open.getTime();
        soonestZone = zone;
      }
    });
    if (!soonestZone) return;
    var closedStart = closedPeriodStart_(ctx.settings, soonestZone, ctx.now);
    var syncedSinceClose = sender.syncFinishedAt && sender.syncFinishedAt.getTime() >= closedStart.getTime();
    var finalHours = soonest - ctx.now.getTime() <= finalStretch;
    if (syncedSinceClose && !finalHours) return;
    if (finalHours && followUpSyncIsFresh_(sender, ctx.now)) return;
    var urgent = finalHours;
    if (!best || (urgent && !best.urgent) || (urgent === best.urgent && soonest < best.opensAt)) {
      best = { sender: sender, opensAt: soonest, urgent: urgent };
    }
  });
  return best;
}

function oldestSentForSender_(ctx, email) {
  var oldest = null;
  var key = String(email || '').toLowerCase();
  ctx.leads.forEach(function (lead) {
    if (String(lead.senderEmail || '').toLowerCase() !== key) return;
    if (!lead.sentDate || FINISHED_STATUSES[lead.status]) return;
    if (!oldest || lead.sentDate.getTime() < oldest.getTime()) oldest = lead.sentDate;
  });
  return oldest;
}

function recoverLead_(ctx, lead) {
  var system = systemFor_(ctx, lead.leadId);
  var stage = system.sendStage || 'initial';
  if (alreadyApplied_(lead, stage)) {
    clearRecovery_(system);
    saveSystem_(system);
    return slotResult_(false, false, 'Send for ' + lead.leadId + ' was already recorded.');
  }
  var email = lead.senderEmail || system.senderEmail;
  var sender = email ? ctx.senderByEmail[email.toLowerCase()] : null;
  if (!sender) {
    clearRecovery_(system);
    system.lastError = 'Could not recover the send because the sender account is missing.';
    saveSystem_(system);
    return markLeadError_(ctx, lead, system, system.lastError, 12);
  }
  var token = outreachToken_(lead.leadId, stage);
  var result = gmailListMessages_(sender.email, sentSearchQuery_(lead.email, token), 5, '');
  var failure = handleGmailFailure_(ctx, sender, result, null);
  if (failure) return failure;
  var messages = (result.data && result.data.messages) || [];
  if (messages.length) {
    applySent_(ctx, lead, sender, system, stage, messages[0].threadId, messages[0].id, system.lastSubject);
    appendLog_(lead.leadId, 'Recovered', sender.email, 'Found the sent ' + stage + ' mail and saved the thread.');
    return slotResult_(false, true, 'Recovered the sent email for ' + lead.leadId + '.');
  }
  system.recoveryTries = (Number(system.recoveryTries) || 0) + 1;
  if (system.recoveryTries >= 2) {
    clearRecovery_(system);
    saveSystem_(system);
    appendLog_(lead.leadId, 'Recovered', sender.email, 'No sent copy found. A later slot can send it.');
    return slotResult_(false, true, 'No sent copy found for ' + lead.leadId + '. A later slot can send it.');
  }
  saveSystem_(system);
  appendLog_(lead.leadId, 'Recovered', sender.email, 'Sent mail is not indexed yet.');
  return slotResult_(false, true, 'Sent mail for ' + lead.leadId + ' is not indexed yet.');
}

function confirmReply_(ctx, row) {
  var sender = ctx.senderByEmail[String(row.senderEmail || '').toLowerCase()];
  if (!sender) {
    markInbox_(ctx, function (item) {
      return item.messageId === row.messageId;
    }, 'done');
    return slotResult_(false, false, 'Skipped a reply check with no sender row.');
  }
  var result = gmailGetMetadata_(sender.email, row.messageId);
  var failure = handleGmailFailure_(ctx, sender, result, null);
  if (failure) {
    if (result.code === 404) {
      markInbox_(ctx, function (item) {
        return item.messageId === row.messageId;
      }, 'done');
    }
    return failure;
  }
  var headers = messageHeaders_(result.data);
  var from = headerValue_(headers, 'From');
  if (!from) {
    markInbox_(ctx, function (item) {
      return item.messageId === row.messageId;
    }, 'done');
    return slotResult_(false, true, 'Inbox message had no From header.');
  }
  var lead = findLeadByThread_(ctx, sender.email, row.threadId);
  if (!lead || FINISHED_STATUSES[lead.status]) {
    markInbox_(ctx, function (item) {
      return item.senderEmail.toLowerCase() === sender.email.toLowerCase() && item.threadId === row.threadId;
    }, 'done');
    return slotResult_(false, true, 'Inbox message did not match an open lead.');
  }
  if (isMailerDaemon_(from)) {
    lead.status = 'Bounced';
    lead.followUpDate = null;
    saveLead_(lead);
    clearRecovery_(systemFor_(ctx, lead.leadId));
    saveSystem_(systemFor_(ctx, lead.leadId));
    markInbox_(ctx, function (item) {
      return item.threadId === row.threadId && item.senderEmail.toLowerCase() === sender.email.toLowerCase();
    }, 'done');
    appendLog_(lead.leadId, 'Bounced', sender.email, from);
    return slotResult_(false, true, 'Marked ' + lead.leadId + ' bounced.');
  }
  if (isAutoReply_(headers) || extractEmailAddress_(from) === sender.email.toLowerCase()) {
    markInbox_(ctx, function (item) {
      return item.messageId === row.messageId;
    }, 'done');
    appendLog_(lead.leadId, 'Auto reply', sender.email, 'Ignored an automatic message.');
    return slotResult_(false, true, 'Ignored an automatic message on ' + lead.leadId + '.');
  }
  var when = Number(result.data.internalDate);
  lead.status = 'Replied';
  lead.replyDate = when ? new Date(when) : ctx.now;
  lead.followUpDate = null;
  saveLead_(lead);
  var system = systemFor_(ctx, lead.leadId);
  clearRecovery_(system);
  system.lastError = '';
  saveSystem_(system);
  markInbox_(ctx, function (item) {
    return item.threadId === row.threadId && item.senderEmail.toLowerCase() === sender.email.toLowerCase();
  }, 'done');
  var preview = (headerValue_(headers, 'Subject') + ' ' + (result.data.snippet || '')).toLowerCase();
  var action = /\b(unsubscribe|stop|remove)\b/.test(preview) ? 'Unsubscribe reply' : 'Reply';
  appendLog_(lead.leadId, action, sender.email, from);
  return slotResult_(false, true, 'Recorded a reply on ' + lead.leadId + '.');
}

function syncOnePage_(ctx, sender, startNew) {
  var query = sender.syncQuery;
  var pageToken = sender.inboxPageToken;
  var keptRows = ctx.inboxRows;
  if (startNew) {
    keptRows = ctx.inboxRows.filter(function (row) {
      return String(row.senderEmail || '').toLowerCase() !== sender.email.toLowerCase();
    });
    query = inboxQueryForDays_(newerThanDays_(oldestSentForSender_(ctx, sender.email), ctx.now));
    pageToken = '';
  }
  if (!query) {
    query = inboxQueryForDays_(newerThanDays_(oldestSentForSender_(ctx, sender.email), ctx.now));
  }
  var result = gmailListMessages_(sender.email, query, ctx.settings.inboxPageSize, pageToken);
  var failure = handleGmailFailure_(ctx, sender, result, null);
  if (failure) return failure;

  var messages = (result.data && result.data.messages) || [];
  var known = {};
  keptRows.forEach(function (row) {
    if (row.messageId) known[row.messageId] = true;
  });
  var matched = 0;
  messages.forEach(function (message) {
    if (!message.id || known[message.id]) return;
    var lead = findOpenLeadByThread_(ctx, sender.email, message.threadId);
    if (!lead) return;
    keptRows.push({
      senderEmail: sender.email,
      messageId: message.id,
      threadId: message.threadId || '',
      listedAt: ctx.now,
      state: 'queued'
    });
    known[message.id] = true;
    matched++;
  });
  ctx.inboxRows = keptRows;
  saveInboxRows_(ctx.inboxRows);

  if (result.data && result.data.nextPageToken) {
    sender.inboxPageToken = result.data.nextPageToken;
    sender.syncQuery = query;
    sender.syncFinishedAt = null;
  } else {
    sender.inboxPageToken = '';
    sender.syncQuery = '';
    sender.syncFinishedAt = ctx.now;
  }
  sender.authStatus = 'Authorized';
  saveSender_(sender);
  appendLog_('', 'Inbox sync', sender.email, 'Read one page. Matched ' + matched + ' open thread' + (matched === 1 ? '' : 's') + '.');
  var done = sender.syncFinishedAt ? ' Inbox sync is complete.' : ' More pages remain.';
  return slotResult_(false, true, 'Synced one inbox page for ' + sender.email + '.' + done);
}

function fetchRfcMessageId_(ctx, candidate) {
  var lead = candidate.lead;
  var sender = candidate.sender;
  var system = systemFor_(ctx, lead.leadId);
  var result = gmailGetMetadata_(sender.email, system.gmailApiMessageId);
  if (!result.ok && result.code === 404) {
    system.gmailApiMessageId = '';
    saveSystem_(system);
    return slotResult_(false, true, 'The stored Gmail message was gone. The follow-up will use the thread id.');
  }
  var failure = handleGmailFailure_(ctx, sender, result, null);
  if (failure) return failure;
  system.lastMessageId = rfcMessageIdFrom_(result.data);
  if (!system.lastMessageId) system.gmailApiMessageId = '';
  saveSystem_(system);
  appendLog_(lead.leadId, 'Message id', sender.email, system.lastMessageId || 'No Message-ID header');
  return slotResult_(false, true, 'Saved the message id for ' + lead.leadId + '. The follow-up waits for the next slot.');
}

function sendCandidate_(ctx, candidate) {
  var lead = candidate.lead;
  var sender = candidate.sender;
  var stage = candidate.stage;
  var local = leadSendWindow_(ctx, lead);
  if (!local.ok) {
    return slotResult_(true, false, 'Skipped ' + (lead.leadId || lead.email) + '. Local time is outside weekday 9:00am to 8:00pm.');
  }
  var system = systemFor_(ctx, lead.leadId);
  if (!system.setId) system.setId = pickSetId_(ctx.templates);
  var template = findTemplate_(ctx.templates, system.setId, stage);
  if (!system.setId || !template) {
    return markLeadError_(ctx, lead, system, 'Missing an active template for set ' + (system.setId || '(none)') + ' stage ' + stage + '.', 12);
  }
  var subject = personalize_(template.subject, lead);
  var token = outreachToken_(lead.leadId, stage);
  system.needsRecovery = true;
  system.sendStage = stage;
  system.claimedAt = ctx.now;
  system.recoveryTries = 0;
  system.senderEmail = sender.email;
  system.lastSubject = subject;
  system.lastError = '';
  saveSystem_(system);
  if (!lead.senderEmail) {
    lead.senderEmail = sender.email;
    saveLead_(lead);
  }
  SpreadsheetApp.flush();

  var replyId = stage === 'initial' ? '' : formatRfcMessageId_(system.lastMessageId);
  var raw = buildRawMessage_({
    from: formatAddress_(sender.displayName || sender.email.split('@')[0], sender.email),
    to: formatAddress_(lead.name, lead.email),
    subject: subject,
    date: ctx.now,
    inReplyTo: replyId,
    references: replyId,
    token: token,
    html: htmlBody_(personalize_(template.body, lead), token)
  });
  var result = gmailSend_(sender.email, raw, stage === 'initial' ? '' : lead.threadId);
  if (!result.ok) {
    clearRecovery_(system);
    saveSystem_(system);
    return handleGmailFailure_(ctx, sender, result, lead);
  }
  if (!result.data || !result.data.threadId) {
    appendLog_(lead.leadId, 'Sent ' + stage, sender.email, 'Gmail did not return a thread id. The next slot will look up the sent mail.');
    return slotResult_(false, true, 'Sent mail for ' + lead.leadId + ' needs a thread id. The next slot will look it up.');
  }
  applySent_(ctx, lead, sender, system, stage, result.data.threadId, result.data.id, subject);
  appendLog_(lead.leadId, 'Sent ' + stage, sender.email, subject);
  return slotResult_(false, true, 'Sent ' + stage + ' to ' + lead.email + '.');
}

function applySent_(ctx, lead, sender, system, stage, threadId, apiMessageId, subject) {
  var gaps = ctx.settings.followUpGaps || [3, 4, 5];
  if (threadId) lead.threadId = threadId;
  if (stage === 'initial') {
    lead.status = 'Sent';
    lead.senderEmail = sender.email;
    lead.sentDate = ctx.now;
    lead.replyDate = null;
    lead.followUpDate = addGap_(ctx.now, gaps[0] || 3, ctx.settings.jitterHoursMax);
  } else if (stage === 'fu1') {
    lead.status = 'Follow-up 1';
    lead.followUpDate = addGap_(ctx.now, gaps[1] || 4, ctx.settings.jitterHoursMax);
  } else if (stage === 'fu2') {
    lead.status = 'Follow-up 2';
    lead.followUpDate = addGap_(ctx.now, gaps[2] || 5, ctx.settings.jitterHoursMax);
  } else if (stage === 'fu3') {
    lead.status = 'Closed';
    lead.followUpDate = ctx.now;
  }
  saveLead_(lead);
  clearRecovery_(system);
  system.lastError = '';
  system.resumeStatus = '';
  system.lastSubject = subject || system.lastSubject;
  system.gmailApiMessageId = apiMessageId || system.gmailApiMessageId;
  system.senderEmail = sender.email;
  if (stage === 'initial') system.setId = system.setId || '';
  saveSystem_(system);
  sender.sentToday = (Number(sender.sentToday) || 0) + 1;
  sender.lastUsedAt = ctx.now;
  sender.quotaDate = todayString_(ctx.settings.timezone, ctx.now);
  sender.authStatus = 'Authorized';
  saveSender_(sender);
}

function alreadyApplied_(lead, stage) {
  if (stage === 'initial') return lead.status === 'Sent' && !!lead.threadId;
  if (stage === 'fu1') return lead.status === 'Follow-up 1' || lead.status === 'Follow-up 2' || lead.status === 'Closed';
  if (stage === 'fu2') return lead.status === 'Follow-up 2' || lead.status === 'Closed';
  if (stage === 'fu3') return lead.status === 'Closed';
  return false;
}

function clearRecovery_(system) {
  system.needsRecovery = false;
  system.sendStage = '';
  system.claimedAt = null;
  system.recoveryTries = 0;
}

function markLeadError_(ctx, lead, system, message, backoffHours) {
  var resume = normalizedStatus_(lead, system);
  if (!resume || resume === 'Error') resume = 'Pending';
  system.resumeStatus = resume;
  system.lastError = message;
  clearRecovery_(system);
  lead.status = 'Error';
  lead.followUpDate = new Date(ctx.now.getTime() + backoffHours * 60 * 60 * 1000);
  saveSystem_(system);
  saveLead_(lead);
  appendLog_(lead.leadId, 'Error', lead.senderEmail || system.senderEmail, message);
  return slotResult_(false, false, message);
}

function handleGmailFailure_(ctx, sender, result, lead) {
  if (result.ok) return null;
  if (result.type === 'rate_limit') {
    setCooldown_(sender.email, 30 * 60 * 1000);
    setGlobalCooldown_(15 * 60 * 1000);
    appendLog_(lead ? lead.leadId : '', 'Rate limit', sender.email, result.message);
    return slotResult_(true, true, 'Gmail rate limit. This tick stopped.');
  }
  if (result.type === 'quota') {
    if (lead) {
      sender.sentToday = senderLimit_(sender, ctx.settings);
      sender.quotaDate = todayString_(ctx.settings.timezone, ctx.now);
      saveSender_(sender);
    } else {
      setCooldown_(sender.email, 60 * 60 * 1000);
    }
    appendLog_(lead ? lead.leadId : '', 'Quota', sender.email, result.message);
    return slotResult_(true, true, sender.email + ' hit a Gmail quota. Remaining work waits.');
  }
  if (result.type === 'auth') {
    sender.authStatus = 'Needs auth';
    saveSender_(sender);
    setCooldown_(sender.email, 6 * 60 * 60 * 1000);
    appendLog_(lead ? lead.leadId : '', 'Auth', sender.email, result.message);
    return slotResult_(true, true, sender.email + ' needs authorization again.');
  }
  if (result.type === 'bounce' && lead) {
    lead.status = 'Bounced';
    lead.followUpDate = null;
    saveLead_(lead);
    var system = systemFor_(ctx, lead.leadId);
    system.lastError = result.message;
    clearRecovery_(system);
    saveSystem_(system);
    appendLog_(lead.leadId, 'Bounced', sender.email, result.message);
    return slotResult_(true, true, 'Bounced ' + lead.email + '.');
  }
  if (lead) {
    var record = systemFor_(ctx, lead.leadId);
    record.lastError = result.message;
    clearRecovery_(record);
    var resume = normalizedStatus_(lead, record);
    record.resumeStatus = resume && resume !== 'Error' ? resume : 'Pending';
    lead.status = 'Error';
    lead.followUpDate = new Date(ctx.now.getTime() + 60 * 60 * 1000);
    saveSystem_(record);
    saveLead_(lead);
    appendLog_(lead.leadId, 'Error', sender.email, result.message);
    return slotResult_(true, true, result.message);
  }
  appendLog_('', 'Gmail error', sender ? sender.email : '', result.message);
  return slotResult_(true, true, result.message);
}

function findLeadByThread_(ctx, senderEmail, threadId) {
  var email = String(senderEmail || '').toLowerCase();
  for (var i = 0; i < ctx.leads.length; i++) {
    var lead = ctx.leads[i];
    if (lead.threadId === threadId && String(lead.senderEmail || '').toLowerCase() === email) return lead;
  }
  return null;
}

function findOpenLeadByThread_(ctx, senderEmail, threadId) {
  var lead = findLeadByThread_(ctx, senderEmail, threadId);
  if (!lead || FINISHED_STATUSES[lead.status]) return null;
  return lead;
}

function markInbox_(ctx, predicate, state) {
  var changed = false;
  ctx.inboxRows.forEach(function (row) {
    if (predicate(row) && row.state !== state) {
      row.state = state;
      changed = true;
    }
  });
  if (changed) saveInboxRows_(ctx.inboxRows);
}

function pickSetId_(templates) {
  var sets = [];
  templates.forEach(function (template) {
    if (!template.active || template.stage !== 'initial') return;
    if (sets.indexOf(template.setId) === -1) sets.push(template.setId);
  });
  sets.sort();
  if (!sets.length) return '';
  var props = PropertiesService.getScriptProperties();
  var last = props.getProperty('LAST_TEMPLATE_SET') || '';
  var index = sets.indexOf(last);
  var next = sets[(index + 1) % sets.length];
  props.setProperty('LAST_TEMPLATE_SET', next);
  return next;
}

function findTemplate_(templates, setId, stage) {
  var wanted = String(setId);
  for (var i = 0; i < templates.length; i++) {
    var template = templates[i];
    if (template.active && template.setId === wanted && template.stage === stage) return template;
  }
  return null;
}

function padNumber_(value) {
  var text = String(value);
  while (text.length < 4) text = '0' + text;
  return text;
}

function globalCooldownActive_(now) {
  var raw = PropertiesService.getScriptProperties().getProperty('GLOBAL_COOLDOWN');
  return !!(raw && Number(raw) > now.getTime());
}

function setGlobalCooldown_(ms) {
  PropertiesService.getScriptProperties().setProperty('GLOBAL_COOLDOWN', String(Date.now() + ms));
}

function inCooldown_(email, now) {
  if (!email) return false;
  var raw = PropertiesService.getScriptProperties().getProperty('COOLDOWN_' + String(email).toLowerCase());
  return !!(raw && Number(raw) > now.getTime());
}

function setCooldown_(email, ms) {
  PropertiesService.getScriptProperties().setProperty('COOLDOWN_' + String(email).toLowerCase(), String(Date.now() + ms));
}
