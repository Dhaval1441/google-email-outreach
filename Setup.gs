/**
 * Spreadsheet menu, tab setup, and the single timer.
 * Deploy as a web app (execute as you, anyone can access) before authorizing senders.
 * The trigger must be installed by the same account that authorizes the senders.
 */

function onOpen() {
  var paused = false;
  try {
    paused = isPaused(getSettings());
  } catch (err) {
    paused = false;
  }
  SpreadsheetApp.getUi()
    .createMenu('Outreach')
    .addItem('Setup sheets', 'setupSheets')
    .addItem('Save OAuth client', 'saveOAuthClient')
    .addItem('Authorize sender', 'authorizeSender')
    .addItem('Install trigger', 'installQueueTrigger')
    .addItem('Run one slot now', 'runOneSlotNow')
    .addItem(paused ? 'Resume' : 'Pause', 'togglePause')
    .addToUi();
}

function setupSheets() {
  var ss = spreadsheet_();
  var defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && !ss.getSheetByName(LEADS_SHEET) && defaultSheet.getLastRow() === 0) {
    defaultSheet.setName(LEADS_SHEET);
  }
  var leads = ensureSheet_(LEADS_SHEET, LEAD_HEADERS);
  ensureHeader_(leads, 'Country', 'Website URL');
  var senders = ensureSheet_(SENDERS_SHEET, SENDER_HEADERS);
  var templates = ensureSheet_(TEMPLATES_SHEET, TEMPLATE_HEADERS);
  var settings = ensureSheet_(SETTINGS_SHEET, SETTINGS_HEADERS);
  ensureSheet_(LOG_SHEET, LOG_HEADERS);
  var system = ensureSheet_(SYSTEM_SHEET, SYSTEM_HEADERS);
  var inbox = ensureSheet_(INBOX_SHEET, INBOX_HEADERS);

  writeMissingSettings_(settings);
  saveSetting_('SendWindowStart', '09:00');
  saveSetting_('SendWindowEnd', '20:00');
  saveSetting_('SendOnWeekends', false);
  writeSampleTemplates_(templates);
  writeGuide_(settings);
  applySheetFormats_(leads, senders, templates, settings, system, inbox);

  ss.setSpreadsheetTimeZone('Asia/Kolkata');
  ss.setActiveSheet(leads);
  system.hideSheet();
  inbox.hideSheet();

  SpreadsheetApp.getUi().alert(
    'Sheets are ready.\n\nNext:\n1. Save OAuth client\n2. Deploy this script as a web app (execute as you)\n3. Authorize each sender\n4. Replace the SAMPLE templates\n5. Install the trigger'
  );
}

function saveOAuthClient() {
  var ui = SpreadsheetApp.getUi();
  var idResult = ui.prompt(
    'OAuth client id',
    'Paste the Google Cloud OAuth client id.',
    ui.ButtonSet.OK_CANCEL
  );
  if (idResult.getSelectedButton() !== ui.Button.OK) return;
  var secretResult = ui.prompt(
    'OAuth client secret',
    'Paste the client secret. It is stored in script properties, not on the sheet.',
    ui.ButtonSet.OK_CANCEL
  );
  if (secretResult.getSelectedButton() !== ui.Button.OK) return;
  var props = PropertiesService.getScriptProperties();
  props.setProperty('CLIENT_ID', idResult.getResponseText().trim());
  props.setProperty('CLIENT_SECRET', secretResult.getResponseText().trim());
  ui.alert('Saved the OAuth client. Redirect URI for that client:\n\n' + redirectUri_());
}

function authorizeSender() {
  var ui = SpreadsheetApp.getUi();
  if (typeof OAuth2 === 'undefined') {
    ui.alert('Add the OAuth2 library first.\n\nLibrary id:\n1B7FSrk5Zi6L1rSxxTDgDEUsPzlukDsi4KGuTMorsTQHhGBzBkMun4iDF\n\nVersion 43.');
    return;
  }
  if (!clientConfigured_()) {
    ui.alert('Save the OAuth client first.');
    return;
  }
  var response = ui.prompt(
    'Authorize sender',
    'Enter the Gmail address that will send mail.',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;
  var email = response.getResponseText().trim().toLowerCase();
  if (!isEmailAddress_(email)) {
    ui.alert('Enter a valid email address.');
    return;
  }
  var sender = ensureSender_(email);
  sender.authStatus = 'Needs auth';
  saveSender_(sender);
  PropertiesService.getUserProperties().setProperty('PENDING_AUTH_EMAIL', email);
  var service = getSenderService_(email);
  service.reset();
  var url = service.getAuthorizationUrl();
  var html = HtmlService.createHtmlOutput(
    '<p>Sign in as <b>' + escapeHtml_(email) + '</b> and allow send and read access.</p>' +
      '<p><a href="' + escapeHtml_(url) + '" target="_blank" rel="noopener">Authorize ' + escapeHtml_(email) + '</a></p>' +
      '<p>The OAuth client must include this redirect URI:</p><p>' + escapeHtml_(redirectUri_()) + '</p>' +
      '<p>Deploy the script as a web app first (execute as you, access anyone), or this link will not finish.</p>'
  ).setWidth(460).setHeight(260);
  ui.showModalDialog(html, 'Authorize sender');
}

function installQueueTrigger() {
  var ui = SpreadsheetApp.getUi();
  var settings = getSettings();
  var minutes = allowedTriggerMinutes_(settings.timerMinutes);
  var existing = ScriptApp.getProjectTriggers();
  for (var i = 0; i < existing.length; i++) {
    if (existing[i].getHandlerFunction() === 'processQueue') {
      ScriptApp.deleteTrigger(existing[i]);
    }
  }
  ScriptApp.newTrigger('processQueue').timeBased().everyMinutes(minutes).create();
  ui.alert(
    'Installed one trigger. It runs processQueue every ' +
      minutes +
      ' minutes.\n\nInstall it while you are signed in as the sheet owner. That is the account whose sender tokens the trigger can use.'
  );
}

function togglePause() {
  var ui = SpreadsheetApp.getUi();
  var settings = getSettings();
  var paused = !settings.pause;
  saveSetting_('Pause', paused);
  ui.alert(paused ? 'Paused. The timer will not send until you resume.' : 'Resumed. The timer will take the next slot in the send window.');
}

function writeMissingSettings_(sheet) {
  var grid = readGrid_(sheet);
  var present = {};
  for (var i = 1; i < grid.values.length; i++) {
    var key = String(cell_(grid.values[i], grid.index, 'Key') || '').trim();
    if (key) present[key] = true;
  }
  var missing = [];
  defaultSettingsList_().forEach(function (pair) {
    if (!present[pair[0]]) missing.push(pair);
  });
  if (!missing.length) return;
  var start = Math.max(sheet.getLastRow(), 1) + 1;
  sheet.getRange(start, 1, missing.length, 2).setValues(missing);
}

function writeSampleTemplates_(sheet) {
  if (sheet.getLastRow() > 1) return;
  var rows = sampleTemplates_().map(function (item) {
    return [item.setId, item.stage, item.subject, item.body, true];
  });
  sheet.getRange(2, 1, rows.length, TEMPLATE_HEADERS.length).setValues(rows);
}

function writeGuide_(sheet) {
  var current = String(sheet.getRange(1, 4).getValue() || '').trim();
  if (current && current !== 'How this runs') return;
  var lines = [
    'How this runs',
    '1. In Google Cloud, enable the Gmail API and create an OAuth web client.',
    '2. Add this redirect URI on that client: ' + redirectUri_(),
    '3. Outreach → Save OAuth client. The secret stays in script properties.',
    '4. Deploy this script as a web app: execute as you, access anyone.',
    '5. Add every sender Gmail as a test user on the consent screen.',
    '6. Outreach → Authorize sender, once per Gmail account.',
    '7. On Leads, fill Lead ID, Name, Email ID, Business Name, Website URL, and Country.',
    'Country decides the clock. Saturday and Sunday in that country are never mailed, including follow-ups.',
    'Mail goes only from 9:00am to 8:00pm local time. USA, Dubai, and European countries are recognized.',
    'At night, and on Sunday evening, the script only checks for replies. It does not send. One check at a time.',
    'USA uses Eastern time unless you write USA Central, USA Mountain, or USA Pacific.',
    '8. Replace the SAMPLE rows on Templates. Keep the Stage names: initial, fu1, fu2, fu3.',
    '9. Placeholders: {{name}} {{businessName}} {{websiteUrl}} {{email}}',
    '10. Outreach → Install trigger. Use the same Google account that authorized the senders.',
    'Each tick sends at most one email. Leads are never mailed in a burst.',
    'Follow-ups run after 3 days, then 4, then 5, on the same Gmail Thread ID, then Status becomes Closed.',
    'A reply fills Reply Date and stops that lead.',
    'Testing-mode refresh tokens expire after 7 days. Authorize sender again when AuthStatus says Needs auth.'
  ];
  var values = lines.map(function (line) {
    return [line];
  });
  sheet.getRange(1, 4, values.length, 1).setValues(values);
  sheet.setColumnWidth(4, 520);
}

function applySheetFormats_(leads, senders, templates, settings, system, inbox) {
  textColumn_(leads, 'Lead ID');
  textColumn_(leads, 'Email ID');
  textColumn_(leads, 'Country');
  textColumn_(leads, 'Gmail Thread ID');
  dateTimeColumn_(leads, 'Sent Date');
  dateTimeColumn_(leads, 'Reply Date');
  dateTimeColumn_(leads, 'Follow-up Date');
  textColumn_(senders, 'Email');
  textColumn_(senders, 'QuotaDate');
  textColumn_(senders, 'InboxPageToken');
  textColumn_(senders, 'SyncQuery');
  dateTimeColumn_(senders, 'LastUsedAt');
  dateTimeColumn_(senders, 'SyncFinishedAt');
  checkboxColumn_(senders, 'Active');
  checkboxColumn_(templates, 'Active');
  textColumn_(system, 'Lead ID');
  textColumn_(system, 'LastMessageId');
  textColumn_(system, 'GmailApiMessageId');
  textColumn_(system, 'SenderEmail');
  dateTimeColumn_(system, 'ClaimedAt');
  textColumn_(inbox, 'Message ID');
  textColumn_(inbox, 'Thread ID');
  dateTimeColumn_(inbox, 'Listed At');
  checkboxSetting_(settings, 'SendOnWeekends');
  checkboxSetting_(settings, 'Pause');
  plainSetting_(settings, 'SendWindowStart');
  plainSetting_(settings, 'SendWindowEnd');
  plainSetting_(settings, 'FollowUpGapDays');
  plainSetting_(settings, 'Timezone');
}

function checkboxColumn_(sheet, headerName) {
  var col = columnIndex_(sheet, headerName);
  if (!col) return;
  var rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  sheet.getRange(2, col, 500, 1).setDataValidation(rule);
}

function settingRow_(sheet, key) {
  var grid = readGrid_(sheet);
  for (var i = 1; i < grid.values.length; i++) {
    if (String(cell_(grid.values[i], grid.index, 'Key')).trim() === key) return i + 1;
  }
  return 0;
}

function checkboxSetting_(sheet, key) {
  var row = settingRow_(sheet, key);
  if (!row) return;
  var rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  sheet.getRange(row, 2).setDataValidation(rule);
}

function plainSetting_(sheet, key) {
  var row = settingRow_(sheet, key);
  if (!row) return;
  sheet.getRange(row, 2).setNumberFormat('@');
}

function sampleTemplates_() {
  return [
    {
      setId: 1,
      stage: 'initial',
      subject: '[SAMPLE] Quick question for {{businessName}}',
      body: 'Hi {{name}},\n\nI was looking at {{websiteUrl}} and had a short question about how {{businessName}} handles new client work.\n\nWe help businesses with outreach and follow-up. If you are exploring that, I can share a few specific ideas.\n\nWould you be open to a short reply?'
    },
    {
      setId: 1,
      stage: 'fu1',
      subject: '[SAMPLE] Following up, {{name}}',
      body: 'Hi {{name}},\n\nI wrote a few days ago about {{businessName}}. I know inboxes get busy.\n\nIf this is not useful, no problem. If it is, reply and I will send the details.'
    },
    {
      setId: 1,
      stage: 'fu2',
      subject: '[SAMPLE] One more note for {{businessName}}',
      body: 'Hi {{name}},\n\nI will keep this short. I am still happy to share how we could help {{businessName}} if it is useful.\n\nIf now is a bad time, you can ignore this and I will send one last note later.'
    },
    {
      setId: 1,
      stage: 'fu3',
      subject: '[SAMPLE] Last note from me',
      body: 'Hi {{name}},\n\nThis is my last email about this. If you want to talk through {{websiteUrl}} and what we could take on, just reply.\n\nOtherwise I will close the loop here.'
    },
    {
      setId: 2,
      stage: 'initial',
      subject: '[SAMPLE] {{name}}, a note about {{websiteUrl}}',
      body: 'Hi {{name}},\n\nI came across {{businessName}} and {{websiteUrl}}.\n\nIf you are looking for help with client outreach, I can outline what that would look like for you.\n\nIs that worth a reply?'
    },
    {
      setId: 2,
      stage: 'fu1',
      subject: '[SAMPLE] Checking back with {{businessName}}',
      body: 'Hi {{name}},\n\nWanted to bump this in case my note about {{businessName}} was buried.\n\nHappy to share a short outline if it is relevant.'
    },
    {
      setId: 2,
      stage: 'fu2',
      subject: '[SAMPLE] Still open if useful, {{name}}',
      body: 'Hi {{name}},\n\nI have one more check-in after this. If outreach support is not something {{businessName}} needs, you can leave this unread.\n\nIf it is, tell me and I will send specifics.'
    },
    {
      setId: 2,
      stage: 'fu3',
      subject: '[SAMPLE] Closing the loop on {{businessName}}',
      body: 'Hi {{name}},\n\nThis is the last follow-up from me. I will not email again unless you reply.\n\nIf you want to revisit {{websiteUrl}} later, just answer this thread.'
    }
  ];
}
