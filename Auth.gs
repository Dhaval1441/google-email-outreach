/**
 * One OAuth2 service per sender Gmail.
 * Tokens live in the sheet owner's user properties. Install the trigger as that same account.
 */

var GMAIL_SCOPES = [
  'https://www.googleapis.com/auth/gmail.send',
  'https://www.googleapis.com/auth/gmail.readonly'
].join(' ');

function redirectUri_() {
  return 'https://script.google.com/macros/d/' + ScriptApp.getScriptId() + '/usercallback';
}

function clientConfigured_() {
  var props = PropertiesService.getScriptProperties();
  return !!(props.getProperty('CLIENT_ID') && props.getProperty('CLIENT_SECRET'));
}

function getSenderService_(email) {
  if (typeof OAuth2 === 'undefined') {
    throw new Error('Add the OAuth2 library (id 1B7FSrk5Zi6L1rSxxTDgDEUsPzlukDsi4KGuTMorsTQHhGBzBkMun4iDF, version 43).');
  }
  var props = PropertiesService.getScriptProperties();
  var clientId = props.getProperty('CLIENT_ID');
  var clientSecret = props.getProperty('CLIENT_SECRET');
  if (!clientId || !clientSecret) {
    throw new Error('Save the OAuth client from the Outreach menu first.');
  }
  email = String(email || '').trim().toLowerCase();
  return OAuth2.createService('gmail:' + email)
    .setAuthorizationBaseUrl('https://accounts.google.com/o/oauth2/v2/auth')
    .setTokenUrl('https://oauth2.googleapis.com/token')
    .setClientId(clientId)
    .setClientSecret(clientSecret)
    .setCallbackFunction('authCallback')
    .setPropertyStore(PropertiesService.getUserProperties())
    .setCache(CacheService.getUserCache())
    .setScope(GMAIL_SCOPES)
    .setParam('login_hint', email)
    .setParam('access_type', 'offline')
    .setParam('prompt', 'consent')
    .setParam('include_granted_scopes', 'true');
}

function senderAccessToken_(email) {
  var service = getSenderService_(email);
  if (!service.hasAccess()) {
    markSenderAuth_(email, 'Needs auth');
    return null;
  }
  return service.getAccessToken();
}

function authCallback(request) {
  var email = PropertiesService.getUserProperties().getProperty('PENDING_AUTH_EMAIL');
  if (!email) {
    return HtmlService.createHtmlOutput(
      'No sender is waiting. Use Outreach → Authorize sender in the spreadsheet, then try again.'
    );
  }
  var service = getSenderService_(email);
  var authorized = false;
  try {
    authorized = service.handleCallback(request);
  } catch (err) {
    return HtmlService.createHtmlOutput('Authorization failed: ' + escapeHtml_(err.message));
  }
  if (!authorized) {
    return HtmlService.createHtmlOutput('Authorization was denied for ' + escapeHtml_(email) + '.');
  }
  try {
    markSenderAuthorized_(email);
  } catch (err) {
    return HtmlService.createHtmlOutput(
      'Authorized, but the Senders tab could not be updated: ' + escapeHtml_(err.message)
    );
  }
  return HtmlService.createHtmlOutput(
    'Authorized ' + escapeHtml_(email) + '. You can close this tab and return to the spreadsheet.'
  );
}

function doGet() {
  return HtmlService.createHtmlOutput(
    'Outreach is running. Use the spreadsheet menu to authorize each sender Gmail.'
  );
}

function escapeHtml_(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}
