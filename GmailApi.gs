/**
 * Gmail API calls. Each function makes one HTTP request.
 * Callers must not invoke more than one of these in a single slot.
 */

var GMAIL_ROOT = 'https://gmail.googleapis.com/gmail/v1/users/me/';

function gmailFetch_(email, method, path, queryPairs, payload) {
  var token = senderAccessToken_(email);
  if (!token) {
    return { ok: false, type: 'auth', message: 'Needs auth for ' + email };
  }
  var url = GMAIL_ROOT + path;
  var query = queryString_(queryPairs || []);
  if (query) url += '?' + query;
  var options = {
    method: method,
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  };
  if (payload !== undefined && payload !== null) {
    options.contentType = 'application/json';
    options.payload = JSON.stringify(payload);
  }
  var response;
  try {
    response = UrlFetchApp.fetch(url, options);
  } catch (err) {
    return { ok: false, type: 'other', message: err.message };
  }
  var code = response.getResponseCode();
  var text = response.getContentText() || '';
  var parsed = {};
  if (text) {
    try {
      parsed = JSON.parse(text);
    } catch (err) {
      parsed = { error: { message: text.slice(0, 300) } };
    }
  }
  if (code >= 200 && code < 300) return { ok: true, data: parsed };
  return { ok: false, type: classifyGmailFailure_(code, parsed), message: gmailErrorMessage_(parsed, code), code: code };
}

function queryString_(pairs) {
  var parts = [];
  pairs.forEach(function (pair) {
    if (pair[1] === undefined || pair[1] === null || pair[1] === '') return;
    parts.push(encodeURIComponent(pair[0]) + '=' + encodeURIComponent(pair[1]));
  });
  return parts.join('&');
}

function gmailErrorMessage_(parsed, code) {
  if (parsed && parsed.error) {
    if (parsed.error.message) return String(parsed.error.message).slice(0, 300);
    if (parsed.error.errors && parsed.error.errors.length) {
      return String(parsed.error.errors[0].message || parsed.error.errors[0].reason || '').slice(0, 300);
    }
  }
  return 'Gmail HTTP ' + code;
}

function classifyGmailFailure_(code, parsed) {
  var reason = '';
  var message = '';
  if (parsed && parsed.error) {
    message = String(parsed.error.message || '');
    if (parsed.error.errors && parsed.error.errors.length) {
      reason = String(parsed.error.errors[0].reason || '');
      if (!message) message = String(parsed.error.errors[0].message || '');
    }
    if (!reason && parsed.error.status) reason = String(parsed.error.status);
  }
  var blob = (reason + ' ' + message).toLowerCase();
  if (code === 429 || blob.indexOf('ratelimit') !== -1 || blob.indexOf('rate limit') !== -1 || blob.indexOf('userratelimit') !== -1) {
    return 'rate_limit';
  }
  if (blob.indexOf('daily') !== -1 || blob.indexOf('quotaexceeded') !== -1 || blob.indexOf('sending quota') !== -1) {
    return 'quota';
  }
  if (
    code === 401 ||
    blob.indexOf('invalid credentials') !== -1 ||
    blob.indexOf('invalid_grant') !== -1 ||
    blob.indexOf('insufficient') !== -1 ||
    blob.indexOf('unauthorized') !== -1
  ) {
    return 'auth';
  }
  if (
    blob.indexOf('invalid to') !== -1 ||
    blob.indexOf('invalid recipient') !== -1 ||
    blob.indexOf('recipient address') !== -1 ||
    blob.indexOf('invalid email') !== -1
  ) {
    return 'bounce';
  }
  return 'other';
}

function gmailListMessages_(email, query, maxResults, pageToken) {
  var pairs = [
    ['q', query],
    ['maxResults', String(maxResults || 15)],
    ['fields', 'messages(id,threadId),nextPageToken']
  ];
  if (pageToken) pairs.push(['pageToken', pageToken]);
  return gmailFetch_(email, 'get', 'messages', pairs, null);
}

function gmailGetMetadata_(email, messageId) {
  var pairs = [
    ['format', 'metadata'],
    ['metadataHeaders', 'From'],
    ['metadataHeaders', 'Subject'],
    ['metadataHeaders', 'Auto-Submitted'],
    ['metadataHeaders', 'Precedence'],
    ['metadataHeaders', 'Message-ID'],
    ['metadataHeaders', 'Message-Id'],
    ['fields', 'id,threadId,internalDate,snippet,payload/headers']
  ];
  return gmailFetch_(email, 'get', 'messages/' + encodeURIComponent(messageId), pairs, null);
}

function gmailSend_(email, raw, threadId) {
  var body = { raw: raw };
  if (threadId) body.threadId = threadId;
  return gmailFetch_(email, 'post', 'messages/send', [['fields', 'id,threadId']], body);
}

function outreachToken_(leadId, stage) {
  return 'outreach:' + String(leadId || '').replace(/["\r\n]/g, '') + ':' + stage;
}

function buildRawMessage_(options) {
  var lines = [];
  lines.push('From: ' + options.from);
  lines.push('To: ' + options.to);
  lines.push('Subject: ' + encodeMimeHeader_(options.subject));
  lines.push('Date: ' + options.date.toUTCString());
  lines.push('MIME-Version: 1.0');
  if (options.inReplyTo) {
    lines.push('In-Reply-To: ' + options.inReplyTo);
    lines.push('References: ' + (options.references || options.inReplyTo));
  }
  lines.push('X-Outreach-Token: ' + options.token);
  lines.push('Content-Type: text/html; charset=UTF-8');
  lines.push('Content-Transfer-Encoding: base64');
  lines.push('');
  lines.push(wrapBase64_(Utilities.base64Encode(Utilities.newBlob(options.html).getBytes())));
  var raw = lines.join('\r\n');
  return Utilities.base64EncodeWebSafe(Utilities.newBlob(raw).getBytes()).replace(/=+$/, '');
}

function htmlBody_(text, token) {
  var source = String(text || '');
  var looksLikeHtml = /<\/?[a-z][\s\S]*>/i.test(source);
  var hidden = '<div style="display:none;max-height:0;overflow:hidden">' + escapeHtml_(token) + '</div>';
  if (looksLikeHtml) return source + hidden;
  var escaped = escapeHtml_(source).replace(/\r\n/g, '\n').replace(/\n/g, '<br>\n');
  return '<div>' + escaped + '</div>' + hidden;
}

function encodeMimeHeader_(value) {
  var clean = String(value || '').replace(/[\r\n]+/g, ' ').trim();
  if (/^[\x00-\x7F]*$/.test(clean)) return clean;
  return '=?UTF-8?B?' + Utilities.base64Encode(Utilities.newBlob(clean).getBytes()) + '?=';
}

function formatAddress_(name, email) {
  var display = String(name || '').replace(/[\r\n]+/g, ' ').trim();
  if (!display) return email;
  if (/^[\x20-\x7E]*$/.test(display)) {
    return '"' + display.replace(/\\/g, '\\\\').replace(/"/g, '\\"') + '" <' + email + '>';
  }
  return encodeMimeHeader_(display) + ' <' + email + '>';
}

function wrapBase64_(value) {
  return String(value).replace(/.{1,76}/g, function (chunk) {
    return chunk + '\r\n';
  }).trim();
}

function headerValue_(headers, name) {
  var wanted = String(name || '').toLowerCase();
  var list = headers || [];
  for (var i = 0; i < list.length; i++) {
    if (String(list[i].name || '').toLowerCase() === wanted) return list[i].value || '';
  }
  return '';
}

function messageHeaders_(data) {
  if (!data || !data.payload) return [];
  return data.payload.headers || [];
}

function extractEmailAddress_(fromHeader) {
  var text = String(fromHeader || '').trim();
  var match = /<([^>]+)>/.exec(text);
  if (match) return match[1].trim().toLowerCase();
  return text.toLowerCase();
}

function isAutoReply_(headers) {
  var auto = headerValue_(headers, 'Auto-Submitted').trim().toLowerCase();
  if (auto && auto !== 'no') return true;
  var precedence = headerValue_(headers, 'Precedence').trim().toLowerCase();
  return precedence === 'auto_reply';
}

function isMailerDaemon_(fromHeader) {
  var email = extractEmailAddress_(fromHeader);
  return email.indexOf('mailer-daemon') !== -1 || email.indexOf('postmaster') !== -1;
}

function formatRfcMessageId_(value) {
  var id = String(value || '').trim();
  if (!id) return '';
  if (id.charAt(0) !== '<') id = '<' + id + '>';
  return id;
}

function rfcMessageIdFrom_(data) {
  var headers = messageHeaders_(data);
  return formatRfcMessageId_(headerValue_(headers, 'Message-ID') || headerValue_(headers, 'Message-Id'));
}

function sentSearchQuery_(email, token) {
  return 'in:sent to:"' + String(email || '').replace(/"/g, '') + '" "' + token + '" newer_than:14d';
}

function inboxQueryForDays_(days) {
  var windowDays = days;
  if (!windowDays || windowDays < 1) windowDays = 1;
  if (windowDays > 45) windowDays = 45;
  return 'in:inbox newer_than:' + windowDays + 'd';
}

function newerThanDays_(oldestSent, now) {
  if (!oldestSent) return 45;
  var ms = now.getTime() - oldestSent.getTime();
  var days = Math.ceil(ms / (24 * 60 * 60 * 1000)) + 1;
  if (days < 1) return 1;
  if (days > 45) return 45;
  return days;
}
