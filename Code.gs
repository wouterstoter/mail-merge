function selectDraft() {
  var drafts = GmailApp.getDrafts();
  var subjects = drafts.map((a,i) => (i + 1) + ". " + a.getMessage().getSubject())
  var ui = SpreadsheetApp.getUi();
  var draft = ui.prompt("Pick a draft","Pick a draft e-mail to use as a template. You have the following draft emails in your GMail:\n" + subjects.join("\n") + "\nPlease put the number or subject of the template below",ui.ButtonSet.OK_CANCEL)
  if (draft.getSelectedButton() == ui.Button.OK) {
    try {
      draft = draft.getResponseText();
      draft = isNaN(Number(draft)) ? drafts[subjects.indexOf(draft)] : drafts[Number(draft) - 1]
    } catch(e) {
      throw new Error("Oops - can't find Gmail draft");
    }
  } else {
    return
  }
  PropertiesService.getScriptProperties().setProperty('draft', draft.getId());
  return draft;
}
function getDraft() {
  var draft = PropertiesService.getScriptProperties().getProperty('draft');
  if (typeof draft === 'string') {
    draft = GmailApp.getDraft(draft)
  }
  if (!draft) {
    draft = selectDraft();
  }
  var msg = draft.getMessage();
  return {
    recipient: msg.getTo(),
    subject: msg.getSubject(),
    body: msg.getPlainBody(),
    options: {
      htmlBody: msg.getBody(),
      bcc: msg.getBcc(),
      cc: msg.getCc(),
      from: msg.getFrom(),
      name: null,
      replyTo: msg.getReplyTo(),
      noReply: false,
      attachments: msg.getAttachments()
    }
  };
}
function getValues() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const dataRange = sheet.getDataRange();
  var f = function(a,n) {return !sheet.isRowHiddenByFilter(n + 2) && !sheet.isRowHiddenByUser(n + 2)}
  var values = dataRange.getDisplayValues();
  var heads = values.shift();
  values = values.filter(f);
  var _values = dataRange.getValues().slice(1).filter(f);
  var obj = values.map((r,n) => (heads.reduce((o, k, i) => (o["_"+k] = _values[n][i] || '', o[k] = r[i] || '', o), {})));
  return obj
}
function fillTemplates() {
  var draft = getDraft();
  var data = getValues();
  data = data.map(a => {
    var d = JSON.parse(JSON.stringify(draft));
    d.subject = evaluateTxt(d.subject,a);
    d.body = evaluateTxt(d.body,a);
    d.options.htmlBody = evaluateHTML(d.options.htmlBody,a)
    d.recipient = a.recipient || d.recipient;
    for (o in d.options) {
      if (o != "htmlBody") d.options[o] = a[o] || d.options[o];
    }
    return d;
  })
  return data;
}
function sendEmails() {
  var templates = fillTemplates();
  var quota = MailApp.getRemainingDailyQuota();
  var ui = SpreadsheetApp.getUi();
  if (quota == 0) {
    ui.alert("You reached your quota for today, so you cannot send any more e-mails with this account.")
  } else if (templates.length > quota) {
    var cont = ui.alert(`You are trying to send ${templates.length} email(s), but can only send ${quota} more email(s) today. Continue sending the first ${quota} email(s)?`,ui.ButtonSet.YES_NO)
    if (cont.getSelectedButton() == ui.Button.NO) return;
    templates = templates.slice(0,quota)
  }
  for (t = 0; t < templates.length; ++t) {
    MailApp.sendEmail(...Object.values(templates[t]))
  }
}
function previewEmails() {
  var t = HtmlService.createTemplateFromFile("preview")
  t.templates = fillTemplates();
  SpreadsheetApp.getUi()
    .showModalDialog(t.evaluate(), 'Preview');
}
/** 
 * Creates the menu item "Mail Merge" for user to run scripts on drop-down.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Mail Merge')
      .addItem('Select Draft','selectDraft')
      .addItem('Preview Emails','previewEmails')
      .addItem('Send Emails','sendEmails')
      .addItem('Help','test')
      .addToUi();
}

function HtmlToTxt(x) {
  return XmlService.parse('<d>' + x.replace(/<[^>]*>/g,"") + '</d>').getRootElement().getText()
}
function TxtToHtml(x) {
  var xml = XmlService.parse('<d></d>');
  xml.getRootElement().setText(x);
  return XmlService.getRawFormat().format(xml).match(/(?<=<d>)[^]*(?=<\/d>)/)[0];
}
function HtmlToTemplate(x) {
  x = x.replace(/&lt;\?/g,"<?").replace(/\?&gt;/g,"?>")
  x = x.replace(/<br[^>]*>/g,"\r\n").replace(/&#xD;/g,"\r").replace(/&#xA;/g,"\n")
  x = x.replace(/(?<=<\?).*?(?=\?>)/g,HtmlToTxt);
  return HtmlService.createTemplate(x);
}
function evaluateTxt(x,data) {
  return HtmlToTxt(evaluateHTML(TxtToHtml(x),data))
}
function evaluateHTML(x,data) {
  var template = HtmlToTemplate(x);
  for (d in data) {
    template[d] = data[d]
  }
  return template.evaluate().getContent()
}
