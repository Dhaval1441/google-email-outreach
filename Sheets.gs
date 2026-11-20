/**
 * Spreadsheet access for the outreach queue.
 * Leads stays on the 11 user columns. System and InboxIndex are hidden helpers.
 */

var LEADS_SHEET = 'Leads';
var SENDERS_SHEET = 'Senders';
var TEMPLATES_SHEET = 'Templates';
var SETTINGS_SHEET = 'Settings';
var LOG_SHEET = 'Log';
var SYSTEM_SHEET = 'System';
var INBOX_SHEET = 'InboxIndex';

var LEAD_HEADERS = [
  'Lead ID',
  'Name',
  'Email ID',
  'Business Name',
  'Website URL',
  'Country',
  'Sender Email',
  'Sent Date',
  'Status',
  'Reply Date',
  'Follow-up Date',
  'Gmail Thread ID'
];

var SENDER_HEADERS = [
  'Email',
  'DisplayName',
  'Active',
  'DailyLimit',
  'SentToday',
  'QuotaDate',
  'LastUsedAt',
  'AuthStatus',
  'InboxPageToken',
  'SyncFinishedAt',
  'SyncQuery'
];

var TEMPLATE_HEADERS = ['SetId', 'Stage', 'Subject', 'Body', 'Active'];

var SYSTEM_HEADERS = [
  'Lead ID',
  'SetId',
  'LastMessageId',
  'LastSubject',
  'LastError',
  'ClaimedAt',
  'ResumeStatus',
  'NeedsRecovery',
  'SendStage',
  'GmailApiMessageId',
  'RecoveryTries',
  'SenderEmail'
];

var INBOX_HEADERS = ['Sender Email', 'Message ID', 'Thread ID', 'Listed At', 'State'];
var LOG_HEADERS = ['Time', 'Lead ID', 'Action', 'Sender', 'Detail'];
var SETTINGS_HEADERS = ['Key', 'Value'];

var FINISHED_STATUSES = {
  Replied: true,
  Closed: true,
  Bounced: true,
  Duplicate: true
};

var FOLLOW_UP_STATUSES = {
  Sent: true,
  'Follow-up 1': true,
  'Follow-up 2': true
};

function spreadsheet_() {
  var active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) {
    PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', active.getId());
    return active;
  }
  var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) {
    throw new Error('Open the bound spreadsheet and run Outreach → Setup sheets first.');
  }
  return SpreadsheetApp.openById(id);
}

function sheetByName_(name) {
  var sheet = spreadsheet_().getSheetByName(name);
  if (!sheet) {
    throw new Error('Missing the ' + name + ' tab. Run Outreach → Setup sheets.');
  }
  return sheet;
}

function defaultSettingsList_() {
  return [
    ['Timezone', 'Asia/Kolkata'],
    ['SendWindowStart', '09:00'],
    ['SendWindowEnd', '20:00'],
    ['SendOnWeekends', false],
    ['SlotSize', 1],
    ['GapSeconds', 20],
    ['TimerMinutes', 10],
    ['DailyLimitDefault', 30],
    ['FollowUpReserve', 10],
    ['FollowUpGapDays', '3,4,5'],
    ['JitterHoursMax', 6],
    ['InboxPageSize', 15],
    ['SyncReuseHours', 12],
    ['Pause', false]
  ];
}

function headerIndex_(headers) {
  var index = {};
  for (var i = 0; i < headers.length; i++) {
    index[String(headers[i]).trim()] = i;
  }
  return index;
}

function readGrid_(sheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  if (lastRow < 1) {
    return { headers: [], index: {}, values: [] };
  }
  var values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var headers = values[0].map(function (cell) {
    return String(cell).trim();
  });
  return { headers: headers, index: headerIndex_(headers), values: values };
}

function cell_(row, index, name) {
  var col = index[name];
  if (col === undefined || col >= row.length) return '';
  return row[col];
}

function parseSheetDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  if (value === '' || value === null || value === undefined) return null;
  var parsed = new Date(value);
  if (isNaN(parsed.getTime())) return null;
  return parsed;
}

function asBool_(value) {
  if (value === true) return true;
  if (value === false || value === '' || value === null || value === undefined) return false;
  var text = String(value).trim().toLowerCase();
  return text === 'true' || text === 'yes' || text === '1';
}

function asDayString_(value, timeZone) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, timeZone, 'yyyy-MM-dd');
  }
  var text = String(value || '').trim();
  var match = /^(\d{4}-\d{2}-\d{2})/.exec(text);
  return match ? match[1] : '';
}

function todayString_(timeZone, now) {
  return Utilities.formatDate(now || new Date(), timeZone || 'Asia/Kolkata', 'yyyy-MM-dd');
}

function minutesFromValue_(value, timeZone, fallback) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    var formatted = Utilities.formatDate(value, timeZone, 'H:m');
    var parts = formatted.split(':');
    return Number(parts[0]) * 60 + Number(parts[1]);
  }
  var match = /^(\d{1,2}):(\d{2})/.exec(String(value || '').trim());
  if (!match) return fallback;
  return Number(match[1]) * 60 + Number(match[2]);
}

function blankToNullDate_(value) {
  return value ? value : '';
}

function isBlankLead_(lead) {
  return !lead.leadId && !lead.name && !lead.email && !lead.businessName && !lead.websiteUrl;
}

function leadFromRow_(rowNumber, index, row) {
  return {
    rowNumber: rowNumber,
    leadId: String(cell_(row, index, 'Lead ID') || '').trim(),
    name: String(cell_(row, index, 'Name') || '').trim(),
    email: String(cell_(row, index, 'Email ID') || '').trim(),
    businessName: String(cell_(row, index, 'Business Name') || '').trim(),
    websiteUrl: String(cell_(row, index, 'Website URL') || '').trim(),
    country: String(cell_(row, index, 'Country') || '').trim(),
    senderEmail: String(cell_(row, index, 'Sender Email') || '').trim(),
    sentDate: parseSheetDate_(cell_(row, index, 'Sent Date')),
    status: String(cell_(row, index, 'Status') || '').trim(),
    replyDate: parseSheetDate_(cell_(row, index, 'Reply Date')),
    followUpDate: parseSheetDate_(cell_(row, index, 'Follow-up Date')),
    threadId: String(cell_(row, index, 'Gmail Thread ID') || '').trim()
  };
}

function senderFromRow_(rowNumber, index, row) {
  return {
    rowNumber: rowNumber,
    email: String(cell_(row, index, 'Email') || '').trim(),
    displayName: String(cell_(row, index, 'DisplayName') || '').trim(),
    active: asBool_(cell_(row, index, 'Active')),
    dailyLimit: Number(cell_(row, index, 'DailyLimit')) || 0,
    sentToday: Number(cell_(row, index, 'SentToday')) || 0,
    quotaDate: cell_(row, index, 'QuotaDate'),
    lastUsedAt: parseSheetDate_(cell_(row, index, 'LastUsedAt')),
    authStatus: String(cell_(row, index, 'AuthStatus') || '').trim(),
    inboxPageToken: String(cell_(row, index, 'InboxPageToken') || '').trim(),
    syncFinishedAt: parseSheetDate_(cell_(row, index, 'SyncFinishedAt')),
    syncQuery: String(cell_(row, index, 'SyncQuery') || '').trim()
  };
}

function systemFromRow_(rowNumber, index, row) {
  return {
    rowNumber: rowNumber,
    leadId: String(cell_(row, index, 'Lead ID') || '').trim(),
    setId: String(cell_(row, index, 'SetId') || '').trim(),
    lastMessageId: String(cell_(row, index, 'LastMessageId') || '').trim(),
    lastSubject: String(cell_(row, index, 'LastSubject') || '').trim(),
    lastError: String(cell_(row, index, 'LastError') || '').trim(),
    claimedAt: parseSheetDate_(cell_(row, index, 'ClaimedAt')),
    resumeStatus: String(cell_(row, index, 'ResumeStatus') || '').trim(),
    needsRecovery: asBool_(cell_(row, index, 'NeedsRecovery')),
    sendStage: String(cell_(row, index, 'SendStage') || '').trim(),
    gmailApiMessageId: String(cell_(row, index, 'GmailApiMessageId') || '').trim(),
    recoveryTries: Number(cell_(row, index, 'RecoveryTries')) || 0,
    senderEmail: String(cell_(row, index, 'SenderEmail') || '').trim()
  };
}

function templateFromRow_(rowNumber, index, row) {
  return {
    rowNumber: rowNumber,
    setId: String(cell_(row, index, 'SetId') || '').trim(),
    stage: String(cell_(row, index, 'Stage') || '').trim(),
    subject: String(cell_(row, index, 'Subject') || ''),
    body: String(cell_(row, index, 'Body') || ''),
    active: asBool_(cell_(row, index, 'Active'))
  };
}

function inboxFromRow_(rowNumber, index, row) {
  return {
    rowNumber: rowNumber,
    senderEmail: String(cell_(row, index, 'Sender Email') || '').trim(),
    messageId: String(cell_(row, index, 'Message ID') || '').trim(),
    threadId: String(cell_(row, index, 'Thread ID') || '').trim(),
    listedAt: parseSheetDate_(cell_(row, index, 'Listed At')),
    state: String(cell_(row, index, 'State') || '').trim()
  };
}

function loadLeads_() {
  var grid = readGrid_(sheetByName_(LEADS_SHEET));
  var leads = [];
  for (var i = 1; i < grid.values.length; i++) {
    var lead = leadFromRow_(i + 1, grid.index, grid.values[i]);
    if (!isBlankLead_(lead)) leads.push(lead);
  }
  return leads;
}

function loadSenders_() {
  var grid = readGrid_(sheetByName_(SENDERS_SHEET));
  var senders = [];
  for (var i = 1; i < grid.values.length; i++) {
    var sender = senderFromRow_(i + 1, grid.index, grid.values[i]);
    if (sender.email) senders.push(sender);
  }
  return senders;
}

function loadTemplates_() {
  var grid = readGrid_(sheetByName_(TEMPLATES_SHEET));
  var templates = [];
  for (var i = 1; i < grid.values.length; i++) {
    var template = templateFromRow_(i + 1, grid.index, grid.values[i]);
    if (template.setId && template.stage) templates.push(template);
  }
  return templates;
}

function loadSystemMap_() {
  var grid = readGrid_(sheetByName_(SYSTEM_SHEET));
  var map = {};
  for (var i = 1; i < grid.values.length; i++) {
    var record = systemFromRow_(i + 1, grid.index, grid.values[i]);
    if (record.leadId) map[record.leadId] = record;
  }
  return map;
}

function loadInboxRows_() {
  var grid = readGrid_(sheetByName_(INBOX_SHEET));
  var rows = [];
  for (var i = 1; i < grid.values.length; i++) {
    var record = inboxFromRow_(i + 1, grid.index, grid.values[i]);
    if (record.messageId || record.threadId) rows.push(record);
  }
  return rows;
}

function getSettings() {
  var defaults = {};
  defaultSettingsList_().forEach(function (pair) {
    defaults[pair[0]] = pair[1];
  });
  var stored = {};
  var sheet = spreadsheet_().getSheetByName(SETTINGS_SHEET);
  if (sheet) {
    var grid = readGrid_(sheet);
    for (var i = 1; i < grid.values.length; i++) {
      var key = String(cell_(grid.values[i], grid.index, 'Key') || '').trim();
      if (!key) continue;
      stored[key] = cell_(grid.values[i], grid.index, 'Value');
    }
  }
  var timezone = String(stored.Timezone || defaults.Timezone).trim() || 'Asia/Kolkata';
  var gaps = String(stored.FollowUpGapDays || defaults.FollowUpGapDays)
    .split(',')
    .map(function (part) {
      return Number(String(part).trim());
    })
    .filter(function (n) {
      return n > 0;
    });
  if (gaps.length < 3) gaps = [3, 4, 5];
  var slotSize = Number(stored.SlotSize);
  if (!slotSize) slotSize = Number(defaults.SlotSize);
  slotSize = Math.max(1, Math.min(20, slotSize));
  var pageSize = Number(stored.InboxPageSize);
  if (!pageSize) pageSize = Number(defaults.InboxPageSize);
  pageSize = Math.max(1, Math.min(20, pageSize));
  return {
    timezone: timezone,
    windowStartMin: minutesFromValue_(stored.SendWindowStart || defaults.SendWindowStart, timezone, 9 * 60),
    windowEndMin: minutesFromValue_(stored.SendWindowEnd || defaults.SendWindowEnd, timezone, 20 * 60),
    sendOnWeekends: asBool_(stored.SendOnWeekends !== undefined ? stored.SendOnWeekends : defaults.SendOnWeekends),
    slotSize: slotSize,
    gapSeconds: Math.max(0, Number(stored.GapSeconds) || Number(defaults.GapSeconds)),
    timerMinutes: Number(stored.TimerMinutes) || Number(defaults.TimerMinutes),
    dailyLimitDefault: Number(stored.DailyLimitDefault) || Number(defaults.DailyLimitDefault),
    followUpReserve: Number(stored.FollowUpReserve === '' || stored.FollowUpReserve === undefined
      ? defaults.FollowUpReserve
      : stored.FollowUpReserve) || 0,
    followUpGaps: gaps,
    jitterHoursMax: Number(stored.JitterHoursMax === '' || stored.JitterHoursMax === undefined
      ? defaults.JitterHoursMax
      : stored.JitterHoursMax) || 0,
    inboxPageSize: pageSize,
    syncReuseHours: Number(stored.SyncReuseHours) || Number(defaults.SyncReuseHours),
    pause: asBool_(stored.Pause !== undefined ? stored.Pause : defaults.Pause)
  };
}

function isPaused(settings) {
  return !!(settings && settings.pause);
}

function withinSendWindow(settings, now, timeZone) {
  now = now || new Date();
  var zone = timeZone || settings.timezone || 'Asia/Kolkata';
  if (!settings.sendOnWeekends) {
    var weekday = Number(Utilities.formatDate(now, zone, 'u'));
    if (weekday === 6 || weekday === 7) return false;
  }
  var minutes = minutesFromValue_(now, zone, 0);
  return minutes >= settings.windowStartMin && minutes < settings.windowEndMin;
}

function countryKey_(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/\./g, '')
    .replace(/\s+/g, ' ');
}

function timezoneForCountry_(country) {
  var key = countryKey_(country);
  if (!key) return '';
  var zones = {
    'usa': 'America/New_York',
    'us': 'America/New_York',
    'united states': 'America/New_York',
    'united states of america': 'America/New_York',
    'america': 'America/New_York',
    'usa eastern': 'America/New_York',
    'us eastern': 'America/New_York',
    'eastern usa': 'America/New_York',
    'usa central': 'America/Chicago',
    'us central': 'America/Chicago',
    'central usa': 'America/Chicago',
    'usa mountain': 'America/Denver',
    'us mountain': 'America/Denver',
    'mountain usa': 'America/Denver',
    'usa pacific': 'America/Los_Angeles',
    'us pacific': 'America/Los_Angeles',
    'pacific usa': 'America/Los_Angeles',
    'west coast': 'America/Los_Angeles',
    'dubai': 'Asia/Dubai',
    'uae': 'Asia/Dubai',
    'united arab emirates': 'Asia/Dubai',
    'abu dhabi': 'Asia/Dubai',
    'europe': 'Europe/Berlin',
    'european': 'Europe/Berlin',
    'uk': 'Europe/London',
    'united kingdom': 'Europe/London',
    'great britain': 'Europe/London',
    'britain': 'Europe/London',
    'england': 'Europe/London',
    'scotland': 'Europe/London',
    'wales': 'Europe/London',
    'ireland': 'Europe/Dublin',
    'portugal': 'Europe/Lisbon',
    'spain': 'Europe/Madrid',
    'france': 'Europe/Paris',
    'germany': 'Europe/Berlin',
    'deutschland': 'Europe/Berlin',
    'italy': 'Europe/Rome',
    'netherlands': 'Europe/Amsterdam',
    'holland': 'Europe/Amsterdam',
    'belgium': 'Europe/Brussels',
    'luxembourg': 'Europe/Luxembourg',
    'switzerland': 'Europe/Zurich',
    'austria': 'Europe/Vienna',
    'poland': 'Europe/Warsaw',
    'sweden': 'Europe/Stockholm',
    'norway': 'Europe/Oslo',
    'denmark': 'Europe/Copenhagen',
    'finland': 'Europe/Helsinki',
    'greece': 'Europe/Athens',
    'czech republic': 'Europe/Prague',
    'czechia': 'Europe/Prague',
    'hungary': 'Europe/Budapest',
    'romania': 'Europe/Bucharest',
    'bulgaria': 'Europe/Sofia',
    'croatia': 'Europe/Zagreb',
    'slovakia': 'Europe/Bratislava',
    'slovenia': 'Europe/Ljubljana',
    'estonia': 'Europe/Tallinn',
    'latvia': 'Europe/Riga',
    'lithuania': 'Europe/Vilnius',
    'iceland': 'Atlantic/Reykjavik',
    'malta': 'Europe/Malta',
    'cyprus': 'Asia/Nicosia',
    'ukraine': 'Europe/Kiev',
    'serbia': 'Europe/Belgrade',
    'albania': 'Europe/Tirane',
    'bosnia': 'Europe/Sarajevo',
    'bosnia and herzegovina': 'Europe/Sarajevo',
    'montenegro': 'Europe/Podgorica',
    'north macedonia': 'Europe/Skopje',
    'macedonia': 'Europe/Skopje',
    'turkey': 'Europe/Istanbul'
  };
  return zones[key] || '';
}

function loadContext_() {
  var settings = getSettings();
  var senders = loadSenders_();
  var senderByEmail = {};
  senders.forEach(function (sender) {
    senderByEmail[sender.email.toLowerCase()] = sender;
  });
  return {
    settings: settings,
    now: new Date(),
    leads: loadLeads_(),
    systemByLeadId: loadSystemMap_(),
    senders: senders,
    senderByEmail: senderByEmail,
    templates: loadTemplates_(),
    inboxRows: loadInboxRows_()
  };
}

function columnIndex_(sheet, headerName) {
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i]).trim() === headerName) return i + 1;
  }
  return 0;
}

function writeLeadColumn_(headerName, leads, valueForLead) {
  if (!leads.length) return;
  var sheet = sheetByName_(LEADS_SHEET);
  var col = columnIndex_(sheet, headerName);
  if (!col) return;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  var current = sheet.getRange(2, col, lastRow - 1, 1).getValues();
  var changed = false;
  leads.forEach(function (lead) {
    var offset = lead.rowNumber - 2;
    if (offset < 0 || offset >= current.length) return;
    var next = valueForLead(lead);
    if (next === undefined) return;
    if (String(current[offset][0]) !== String(next)) {
      current[offset][0] = next;
      changed = true;
    }
  });
  if (changed) sheet.getRange(2, col, lastRow - 1, 1).setValues(current);
}

function saveLead_(lead) {
  var sheet = sheetByName_(LEADS_SHEET);
  var headers = readGrid_(sheet).headers;
  var index = headerIndex_(headers);
  var current = sheet.getRange(lead.rowNumber, 1, 1, headers.length).getValues()[0];
  var values = {
    'Lead ID': lead.leadId || '',
    'Name': lead.name || '',
    'Email ID': lead.email || '',
    'Business Name': lead.businessName || '',
    'Website URL': lead.websiteUrl || '',
    'Country': lead.country || '',
    'Sender Email': lead.senderEmail || '',
    'Sent Date': blankToNullDate_(lead.sentDate),
    'Status': lead.status || '',
    'Reply Date': blankToNullDate_(lead.replyDate),
    'Follow-up Date': blankToNullDate_(lead.followUpDate),
    'Gmail Thread ID': lead.threadId || ''
  };
  Object.keys(values).forEach(function (header) {
    if (index[header] === undefined) return;
    current[index[header]] = values[header];
  });
  sheet.getRange(lead.rowNumber, 1, 1, headers.length).setValues([current]);
}

function systemValues_(record) {
  return [
    record.leadId || '',
    record.setId || '',
    record.lastMessageId || '',
    record.lastSubject || '',
    record.lastError || '',
    blankToNullDate_(record.claimedAt),
    record.resumeStatus || '',
    record.needsRecovery ? true : false,
    record.sendStage || '',
    record.gmailApiMessageId || '',
    record.recoveryTries || 0,
    record.senderEmail || ''
  ];
}

function saveSystem_(record) {
  var sheet = sheetByName_(SYSTEM_SHEET);
  var values = systemValues_(record);
  if (record.rowNumber) {
    sheet.getRange(record.rowNumber, 1, 1, SYSTEM_HEADERS.length).setValues([values]);
    return record;
  }
  sheet.appendRow(values);
  record.rowNumber = sheet.getLastRow();
  return record;
}

function emptySystem_(leadId) {
  return {
    rowNumber: 0,
    leadId: leadId,
    setId: '',
    lastMessageId: '',
    lastSubject: '',
    lastError: '',
    claimedAt: null,
    resumeStatus: '',
    needsRecovery: false,
    sendStage: '',
    gmailApiMessageId: '',
    recoveryTries: 0,
    senderEmail: ''
  };
}

function systemFor_(ctx, leadId) {
  if (!leadId) return emptySystem_('');
  if (!ctx.systemByLeadId[leadId]) ctx.systemByLeadId[leadId] = emptySystem_(leadId);
  return ctx.systemByLeadId[leadId];
}

function senderValues_(sender) {
  return [
    sender.email || '',
    sender.displayName || '',
    sender.active ? true : false,
    sender.dailyLimit || '',
    sender.sentToday || 0,
    sender.quotaDate || '',
    blankToNullDate_(sender.lastUsedAt),
    sender.authStatus || '',
    sender.inboxPageToken || '',
    blankToNullDate_(sender.syncFinishedAt),
    sender.syncQuery || ''
  ];
}

function saveSender_(sender) {
  var sheet = sheetByName_(SENDERS_SHEET);
  var values = senderValues_(sender);
  if (sender.rowNumber) {
    sheet.getRange(sender.rowNumber, 1, 1, SENDER_HEADERS.length).setValues([values]);
    return sender;
  }
  sheet.appendRow(values);
  sender.rowNumber = sheet.getLastRow();
  return sender;
}

function ensureSender_(email) {
  email = String(email || '').trim().toLowerCase();
  var senders = loadSenders_();
  for (var i = 0; i < senders.length; i++) {
    if (senders[i].email.toLowerCase() === email) return senders[i];
  }
  var settings = getSettings();
  var sender = {
    rowNumber: 0,
    email: email,
    displayName: email.split('@')[0],
    active: true,
    dailyLimit: settings.dailyLimitDefault,
    sentToday: 0,
    quotaDate: todayString_(settings.timezone, new Date()),
    lastUsedAt: null,
    authStatus: 'Needs auth',
    inboxPageToken: '',
    syncFinishedAt: null,
    syncQuery: ''
  };
  return saveSender_(sender);
}

function markSenderAuthorized_(email) {
  var sender = ensureSender_(email);
  sender.authStatus = 'Authorized';
  sender.active = true;
  saveSender_(sender);
  return sender;
}

function markSenderAuth_(email, status) {
  var senders = loadSenders_();
  for (var i = 0; i < senders.length; i++) {
    if (senders[i].email.toLowerCase() === String(email).toLowerCase()) {
      senders[i].authStatus = status;
      saveSender_(senders[i]);
      return senders[i];
    }
  }
  return null;
}

function resetSenderQuotas_(settings, senders) {
  var today = todayString_(settings.timezone, new Date());
  senders.forEach(function (sender) {
    if (asDayString_(sender.quotaDate, settings.timezone) !== today) {
      sender.sentToday = 0;
      sender.quotaDate = today;
      saveSender_(sender);
    }
  });
}

function saveInboxRows_(rows) {
  var sheet = sheetByName_(INBOX_SHEET);
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, INBOX_HEADERS.length).clearContent();
  }
  if (!rows.length) return;
  var values = rows.map(function (row) {
    return [
      row.senderEmail || '',
      row.messageId || '',
      row.threadId || '',
      blankToNullDate_(row.listedAt),
      row.state || ''
    ];
  });
  sheet.getRange(2, 1, values.length, INBOX_HEADERS.length).setValues(values);
  for (var i = 0; i < rows.length; i++) rows[i].rowNumber = i + 2;
}

function appendLog_(leadId, action, sender, detail) {
  var sheet = spreadsheet_().getSheetByName(LOG_SHEET);
  if (!sheet) return;
  var text = String(detail || '').slice(0, 500);
  sheet.appendRow([new Date(), leadId || '', action || '', sender || '', text]);
  var lastRow = sheet.getLastRow();
  if (lastRow > 5001) {
    sheet.deleteRows(2, lastRow - 5001);
  }
}

function normalizedStatus_(lead, system) {
  var status = lead.status || 'Pending';
  if (status === 'Error') {
    return (system && system.resumeStatus) || 'Pending';
  }
  return status;
}

function isEmailAddress_(value) {
  return /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(String(value || '').trim());
}

function personalize_(text, lead) {
  return String(text || '')
    .replace(/\{\{\s*name\s*\}\}/gi, lead.name || '')
    .replace(/\{\{\s*businessName\s*\}\}/gi, lead.businessName || '')
    .replace(/\{\{\s*websiteUrl\s*\}\}/gi, lead.websiteUrl || '')
    .replace(/\{\{\s*email\s*\}\}/gi, lead.email || '');
}

function addGap_(fromDate, days, jitterHours) {
  var ms = days * 24 * 60 * 60 * 1000;
  var jitter = 0;
  if (jitterHours > 0) jitter = Math.floor(Math.random() * jitterHours * 60 * 60 * 1000);
  return new Date(fromDate.getTime() + ms + jitter);
}

function allowedTriggerMinutes_(value) {
  var allowed = [1, 5, 10, 15, 30];
  var target = Number(value) || 10;
  var best = allowed[0];
  var bestDiff = Math.abs(target - best);
  for (var i = 1; i < allowed.length; i++) {
    var diff = Math.abs(target - allowed[i]);
    if (diff < bestDiff) {
      best = allowed[i];
      bestDiff = diff;
    }
  }
  return best;
}

function saveSetting_(key, value) {
  var sheet = sheetByName_(SETTINGS_SHEET);
  var grid = readGrid_(sheet);
  for (var i = 1; i < grid.values.length; i++) {
    if (String(cell_(grid.values[i], grid.index, 'Key')).trim() === key) {
      var col = grid.index.Value + 1;
      sheet.getRange(i + 1, col).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value]);
}

function formatHeaderRow_(sheet, width) {
  var range = sheet.getRange(1, 1, 1, width);
  range.setFontWeight('bold');
  range.setBackground('#f3f3f3');
  sheet.setFrozenRows(1);
}

function ensureHeader_(sheet, headerName, afterHeader) {
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (cell) {
    return String(cell).trim();
  });
  if (headers.indexOf(headerName) !== -1) return;
  var after = headers.indexOf(afterHeader);
  if (after === -1) {
    sheet.getRange(1, headers.length + 1).setValue(headerName);
    formatHeaderRow_(sheet, headers.length + 1);
    return;
  }
  sheet.insertColumnAfter(after + 1);
  sheet.getRange(1, after + 2).setValue(headerName);
  formatHeaderRow_(sheet, sheet.getLastColumn());
}

function ensureSheet_(name, headers) {
  var ss = spreadsheet_();
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0 || String(sheet.getRange(1, 1).getValue()).trim() === '') {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  formatHeaderRow_(sheet, headers.length);
  return sheet;
}

function textColumn_(sheet, headerName) {
  var col = columnIndex_(sheet, headerName);
  if (!col) return;
  sheet.getRange(2, col, Math.max(sheet.getMaxRows() - 1, 1), 1).setNumberFormat('@');
}

function dateTimeColumn_(sheet, headerName) {
  var col = columnIndex_(sheet, headerName);
  if (!col) return;
  sheet.getRange(2, col, Math.max(sheet.getMaxRows() - 1, 1), 1).setNumberFormat('yyyy-mm-dd hh:mm');
}
