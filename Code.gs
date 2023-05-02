// Copyright Wouter Stoter 2023
//
// Licensed under the Apache License, Version 2.0 (the "License"); you may not
// use this file except in compliance with the License.  You may obtain a copy
// of the License at
//
//     https://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS, WITHOUT
// WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.  See the
// License for the specific language governing permissions and limitations under
// the License.

/**
 * @OnlyCurrentDoc
*/

/**
 * Lets the user pick one of their current drafts in GMail
 * @return {GmailDraft} Draft object of the selected draft
*/
function selectDraft() {
  var drafts = GmailApp.getDrafts(); //Get the drafts
  var subjects = drafts.map((a,i) => (i + 1) + ". " + a.getMessage().getSubject()); //Get the subject lines and number them
  var ui = SpreadsheetApp.getUi();
  var draft = ui.prompt("Pick a draft","Pick a draft e-mail to use as a template. You have the following draft emails in your GMail:\n" + subjects.join("\n") + "\nPlease put the number or subject of the template below",ui.ButtonSet.OK_CANCEL)
  if (draft.getSelectedButton() == ui.Button.OK) {
    try {
      draft = draft.getResponseText();
      draft = isNaN(Number(draft)) ? drafts[subjects.indexOf(draft)] : drafts[Number(draft) - 1]
    } catch(e) {
      throw new Error("Cannot find email draft");
    }
  } else {
    return
  }
  PropertiesService.getScriptProperties().setProperty('draft', draft.getId()); // Save the draft ID in document properties, so the user does not need to pick it again unless they specifically click the option
  return draft;
}
/**
 * Retreive currently saved draft, or make the user select a draft if none present
 * @return {Object} Object containing the relevant properties of the selected draft, neccissary to send the email
*/
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
/**
 * Retreive the data from the current spreadsheet to be used in the draft
 * @return {Array} Array containing an object for each row
*/
function getValues() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]; //Get the first sheet of the spreadsheet
  const dataRange = sheet.getDataRange();
  var f = function(a,n) {return !sheet.isRowHiddenByFilter(n + 2) && !sheet.isRowHiddenByUser(n + 2)} //Filter out hidden rows
  var values = dataRange.getDisplayValues();    
  var heads = values.shift();
  values = values.filter(f);
  var _values = dataRange.getValues().slice(1).filter(f);
  var obj = values.map((r,n) => (heads.reduce((o, k, i) => (o["_"+k] = _values[n][i] || '', o[k] = r[i] || '', o), {})));
  return obj
}
/**
 * Process the email draft using Google App Script to add the variables
 * @return {Array} Array of objects with for each email the most important properties
*/
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
/**
 * Send the email to the specified recipients
*/
function sendEmails() {
  var templates = fillTemplates();
  var quota = MailApp.getRemainingDailyQuota();
  var ui = SpreadsheetApp.getUi();
  // Check user quota for emails to send
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
/**
 * Show the user a preview of the email
*/
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
/**
 * Convert HTML to plain text
 * @param x {String} html code
*/
function HtmlToTxt(x) {
  return XmlService.parse('<d>' + x.replace(/<[^>]*>/g,"") + '</d>').getRootElement().getText()
}
/**
 * Convert plain text to HTML
 * @param x {String} plain text
*/
function TxtToHtml(x) {
  var xml = XmlService.parse('<d></d>');
  xml.getRootElement().setText(x);
  return XmlService.getRawFormat().format(xml).match(/(?<=<d>)[^]*(?=<\/d>)/)[0];
}

/**
 * Convert plain text to HTML
 * @param x {String} html code
 * @return {Template} Google app script template object of email
*/
function HtmlToTemplate(x) {
  x = x.replace(/&lt;\?/g,"<?").replace(/\?&gt;/g,"?>")
  x = x.replace(/<br[^>]*>/g,"\r\n").replace(/&#xD;/g,"\r").replace(/&#xA;/g,"\n")
  x = x.replace(/(?<=<\?).*?(?=\?>)/g,HtmlToTxt);
  return HtmlService.createTemplate(x);
}
/**
 * Evaluate app script commands in plain text
 * @param x {String} plain text containing app script commands
 * @param data {Object} Variables to be used in commands
 * @return {Template} evaluated plain text
*/
function evaluateTxt(x,data) {
  return HtmlToTxt(evaluateHTML(TxtToHtml(x),data))
}
/**
 * Evaluate app script commands in HTML
 * @param x {String} html code containing app script commands
 * @param data {Object} Variables to be used in commands
 * @return {Template} html code with evaluated app script
*/
function evaluateHTML(x,data) {
  var template = HtmlToTemplate(x);
  for (d in data) {
    template[d] = data[d]
  }
  return template.evaluate().getContent()
}
