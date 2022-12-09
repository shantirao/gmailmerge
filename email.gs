// use +"<img src='cid:logo'>" to send images. URLs for included images are hardcoded here 
var media = {
  header: '', //https://path.to/image.png
  logo: ''
  };

var sentColumn = 'Sent';

// Authorization problems? switch to a GCP-managed project so script.run will work
// https://developers.google.com/apps-script/guides/cloud-platform-projects#switching_to_a_different_standard_gcp_project
// Resources -> Cloud Platform project ... 18041258489
// in that project, enable the "spreadsheets" api
// https://developers.google.com/apps-script/guides/cloud-platform-projects
function onOpen() {
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
    .createMenu('Email')
    .addItem('Mailmerge setup', 'mailMergeUI')
    .addItem('Preview send all', 'previewSend')
    .addItem('Send all now', 'sendAll')
    .addSeparator()
    .addItem('Reset sent', 'resetSent')
    .addToUi();
}

function Table(name) //lighter version of https://script.google.com/home/projects/1zJWXp8Rj5bgTx7KOggmBM-VxJGnzKexAAwGnWbGr-OtU3qtCM5Pai6T5/edit
{
  this.sheet = name ? SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name) : SpreadsheetApp.getActiveSheet();
  this.data = this.sheet.getDataRange().getValues();
  this.header = this.data.shift();
  this.length = this.data.length;
}
Table.prototype.column = function(n){var j = this.header.indexOf(n); if (j===-1) {this.header.push(n); for(var i=0;i<this.data.length; i++)this.data[i].push('');j=this.header.length-1; this.sheet.getRange(1, j+1).setValue(n);} return j;}
Table.prototype.get=function(i,j){var d = this.data[i]; if (typeof j==='undefined') {var r={}; for (var j=0;j<this.header.length; j++) r[this.header[j]]=d[j]; return r;} if (typeof j==='string') return d[this.column(j)]; return d[j];}
Table.prototype.set=function(i,j,v) {if (typeof j==='string') j = this.column(j); this.sheet.getRange(i+2, j+1).setValue(this.data[i][j] = v); return this;}

var table = null;
var mediaCache = {}
function getMediaCache() //look for /img.*src='cid:(.*?)'/
{
  for (var n in media) 
  {
    if (!(n in mediaCache) && media[n])
      mediaCache[n] = UrlFetchApp.fetch(media[n]).getBlob().setName(n);
  }
  return mediaCache;
}

function getUnreadEmails() 
{
  return GmailApp.getInboxUnreadCount();
}

function test()
{
 Logger.log( render(0,"${First Name} ${Last Name}"));
}

function previewSend() {sendAll(true);}

function property(name,x)
{
  var p = PropertiesService.getDocumentProperties();
  if (x) p.setProperty(name,x);
   return p.getProperty(name);
  //return x;
}
function resetSent()
{
  if (!table) table = new Table();
  var jSent = table.column(sentColumn);
  for (var i=0; i< table.length; i++)
  table.set(i,jSent,'');
}

function sendAll(preview)
{
  if (!table) table = new Table();
  var count = 0;
  var quota = getQuota();
  var filter = property('filter');
  if (table.header.indexOf(filter) === -1) filter = null;

  for (var i=0; i< table.length; i++)
  {
    if (filter && !table.get(i,filter)) continue;
    if (table.get(i,sentColumn)) continue;
    count++;
  }
  
  var ui = SpreadsheetApp.getUi();
  if (preview) 
    table.column('SendPreview');
  else
  {
    var response = ui.prompt('Are you sure?', 'Type SEND to send '+count+' emails.', ui.ButtonSet.YES_NO);
    if (response.getSelectedButton() !== ui.Button.YES)  return;
    if (response.getResponseText() !== 'SEND') return;
  }
  
  var subject = property('subject');
  var message = property('message');
  var to = property('to');
  for (var i=0; i< table.length && quota; i++)
  {
    if (preview) table.set(i,'SendPreview',false);
    if (filter && !table.get(i,filter)) continue;
    if (table.get(i,sentColumn)) continue;
    if (preview) table.set(i,'SendPreview',true);
    else if (!sendOne(i,to,subject,message)) break;
    quota--;
  }
}

function sendOne(i,to,subject,message)
{
  if (!table) table = new Table();
  var r = table.get(i);
  Logger.log(i,' ',r[to],' ',subject,' ',message);
  try{
    if (sendMail(r[to], null, stringReplace(subject,r),stringReplace(message,r).replace(/\n+/g,'<p>') ))
    {
      table.set(i,sentColumn,new Date);
      return true;
    }
  }
  catch (err)
  {
    Logger.log(err);
    return false;
  }
}

function render(i,text)
{
  if (!table) table = new Table();
  return stringReplaceRow(text,table,i).replace(/\n+/g,'<p>');
}

function getQuota() {return MailApp.getRemainingDailyQuota();}

function tableGet(i,field)
{
  if (!table) table = new Table();
  return table.get(i,field);
}

function tableCount()
{
  if (!table) table = new Table();
  return table.length;
}
function tableHeader()
{
  if (!table) table = new Table();
  return table.header;
}
function stringReplace(s,o) //string,object
{
  if (typeof s !== 'string') return s;
  s = s.replace(/\$\{(.*?)\}/g,function(a,b){return b in o ? o[b] : a;})
  //s = s.replace(/\$(\w+)/g,function(a,b){return b in o ? o[b] : a;}); //replace only if found
return s;
}
function stringReplaceRow(s,table,row) //string,object
{
  if (typeof s !== 'string') return s;
  s = s.replace(/\$\{(.*?)\}/g,function(a,b){return table.header.indexOf(b) !== -1 ? table.get(row,b) : a;})
  //s = s.replace(/\$(\w+)/g,function(a,b){return b in o ? o[b] : a;}); //replace only if found
return s;
}

function testReplace()
{
  var t = new Table()
  var s = stringReplaceRow("${Email}",t,0);
  Logger.log(s);
}

function sendMail(to, cc, subject, message ) 
{
  var msg = {
    to: to,
    subject: subject,
    htmlBody: message,
    inlineImages: getMediaCache() 
  };
  if (cc) msg.cc = cc;
  if (!getQuota()) return false;

  // Logger.log(msg);
  MailApp.sendEmail(msg);
  return true;
}


/**
 * Shows a custom dialog.
 */
function mailMergeUI() {
  var html = HtmlService.createHtmlOutputFromFile('mailmerge')
    .setWidth(700)
    .setHeight(500);
  SpreadsheetApp.getUi().showModelessDialog(html,'Mailmerge')
}
