// Creates a Google Document version of the Google Form Response based on a formatted template document with placeholder text
// and emails a notification and a copy of the form response to the reporter and their supervisors (based on location)
function IncidentReportTemplateCopy(responseSheet, row){
  //EDIT URLS
  var ss = responseSheet;  
  //DOCS
  var templateId = 'GoogleDocId-xxXXxxxXXXx'; //ID of document template (contains placeholders)
  var folder = DriveApp.getFolderById("XXXXX_DriveFolderId"); //FOLDER WHERE Document copy will be saved
  //DOC
  var url = 22; // Column where the Google Doc URL is recorded.
  var dataRange = ss.getRange(row, 1, 1, 23);
  var dataVals = dataRange.getValues();

  var userEm = dataVals[0][2];
  var em = userEm.toString();
  var share = dataVals[0][22].toString(); //creates email to list, includes email from report plus supervisors/managers from library selected
  var emShare = share.split(',');
  var emailList = share + ',' + em;
  //Check Settings
  var settings = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  var ccEm = settings.getRange("Settings!D2").getValue(); // get the emails listed in Settings sheet to which all Incident Reports are sent - CC's to all reports
  var emailTemplate = settings.getRange("F2").getValue(); //HTML email template with placeholders
  
  //Get template
  var template = DriveApp.getFileById(templateId);
  //get date submitted & format for file name
  var date = dataVals[0][0]; 
  var formatDate = Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), "yyyy MMM d, h:mm a");
  //get date and time of incident and format
  var incidentDate = dataVals[0][4];
  var fIncidentDate = Utilities.formatDate(new Date(incidentDate), Session.getScriptTimeZone(), "EEE, MMMMM d, yyyy");
  var incidentTime = dataVals[0][5];
  var fIncidentTime = Utilities.formatDate(new Date(incidentTime), Session.getScriptTimeZone(), "h:mm a");

  //If the form response hasn't been turned into a Google Doc (i.e. new incident report)
  if (!dataVals[0][21]) {
    //Create a Google Doc version of the form response for easier review
    var name = formatDate + " - " + dataVals[0][3] + " - " + dataVals[0][1] + " - Incident Report"; // Date - Library Branch - Name of Person Reporting - Incident Report
    var newReport = template.makeCopy(name, folder)
    var file = DocumentApp.openById(newReport.getId());
    //replace templated values
    var body = file.getBody();
        body.replaceText("%Timestamp%", " " + dataVals[0][0]);
        body.replaceText("%ReporterName%", " " + dataVals[0][1]);
        body.replaceText("%EmailList%", " " + emailList);
        body.replaceText("%Library%", " " + dataVals[0][3]);
        body.replaceText("%Date of Incident%", " " + fIncidentDate);
        body.replaceText("%Time of Incident%", " " + fIncidentTime);
        body.replaceText("%IncidentLocation%", " " + dataVals[0][6]);
        body.replaceText("%StaffInvolved%", " " + dataVals[0][7]);
        body.replaceText("%InvolvedPatrons%", " " + dataVals[0][8]);
        body.replaceText("%Age%", " " + dataVals[0][9]);
        body.replaceText("%MinorAge%", " " + dataVals[0][10]);
        body.replaceText("%ParentName%", " " + dataVals[0][11]);
        body.replaceText("%ContactedParties%", " " + dataVals[0][12]);
        body.replaceText("%ContactInfo%", " " + dataVals[0][13]);
        body.replaceText("%Emergency%", " " + dataVals[0][14]);
        body.replaceText("%IncidentDescription%", " " + dataVals[0][15]);
        body.replaceText("%IncidentType%", " " + dataVals[0][16]);
        body.replaceText("%FollowupComments%", dataVals[0][17]);
        body.replaceText("%LibrarianComments%", dataVals[0][18]);
        body.replaceText("%LibrarianInitials/Date%", dataVals[0][19]);
        body.findText("%ReportURL%").getElement().asText().setText("Edit Report").setLinkUrl(dataVals[0][20]).setFontSize(14).setBold(true).setBackgroundColor('#fff2cc');
    file.addEditors(emShare); //add 'edit' permissions supervisors
    file.saveAndClose();
    ss.getRange(row, url).setValue(file.getUrl()); //store url to the Google Doc in the form response spreadsheet
    //send an email notification and copy of incident report to reported & their supervisor
    var email = emailTemplate;
    var subject = "New Incident Report - " + dataVals[0][3];
          email = email.replace('%NAME%', dataVals[0][1]);
          email = email.replace('%RESPONSEDATE%', dataVals[0][0]);
          email = email.replace('%DATE%', fIncidentDate);
          email = email.replace('%TIME%', fIncidentTime);
          email = email.replace('%LIBRARY%', dataVals[0][3]);
          email = email.replace('%INCIDENTREPORT%', dataVals[0][20]);
          email = email.replace('%LOCATION%', dataVals[0][6]);
          email = email.replace('%STAFF%', dataVals[0][7]);
          email = email.replace('%CONTACTS%', dataVals[0][8]);
          email = email.replace('%AGE%', dataVals[0][9]);
          email = email.replace('%MINORAGE%', dataVals[0][10]);
          email = email.replace('%PARENT%', dataVals[0][11]);
          email = email.replace('%OTHERCONTACT%', dataVals[0][12]);
          email = email.replace('%CONTACTINFO%', dataVals[0][13]);
          email = email.replace('%EMERGENCY%', dataVals[0][14]);
          email = email.replace('%DESCRIPTION%', dataVals[0][15]);
          email = email.replace('%INCIDENTTYPE%', dataVals[0][16]);
          email = email.replace('%LIBCOMMENTS%', dataVals[0][18]);
          email = email.replace('%FOLLOWUP%', dataVals[0][17]);
          email = email.replace('%DOCURL%', file.getUrl());
      MailApp.sendEmail(emailList, subject, "", {htmlBody: email, cc: ccEm});
  } else if (dataVals[0][21]) {
    // If the Google Doc for this form response has already been created and needs to be edited/updated
      var templateDoc = DocumentApp.openById(templateId);
      var templateBdyAll = templateDoc.getBody().getTables(); //get the original template - contains 2 tables
      var templateBdy = templateBdyAll[0].copy(); //copy the first table, which is the template for form responses
      var reportUrl = ss.getRange(row, url).getValue(); //get the url for the Google Doc of the original form response
      var reportId = reportUrl.split("id="); // split the url to get the ID for the Document, the value indexed at [1]
      var report = DocumentApp.openById(reportId[1]);
      var tables = report.getBody().getTables(); // get tables from the form response Doc - contains 2 tables
      var notes = tables[1].copy(); //copy the 2nd table, which is the notes section
        tables[0].removeFromParent() // remove the 1st table (original form responses)
        report.getBody().clear();
        report.getBody().appendTable(templateBdy); // add back in the form template (1st table)
        report.getBody().appendTable(notes); // add back in the notes from the original form response doc
    // replace the placeholder values with the updated form responses
      var body = report.getBody();
          body.replaceText("%Timestamp%", " " + dataVals[0][0]);
          body.replaceText("%ReporterName%", " " + dataVals[0][1]);
          body.replaceText("%EmailList%", " " + emailList);
          body.replaceText("%Library%", " " + dataVals[0][3]);
          body.replaceText("%Date of Incident%", " " + fIncidentDate);
          body.replaceText("%Time of Incident%", " " + fIncidentTime);
          body.replaceText("%IncidentLocation%", " " + dataVals[0][6]);
          body.replaceText("%StaffInvolved%", " " + dataVals[0][7]);
          body.replaceText("%InvolvedPatrons%", " " + dataVals[0][8]);
          body.replaceText("%Age%", " " + dataVals[0][9]);
          body.replaceText("%MinorAge%", " " + dataVals[0][10]);
          body.replaceText("%ParentName%", " " + dataVals[0][11]);
          body.replaceText("%ContactedParties%", " " + dataVals[0][12]);
          body.replaceText("%ContactInfo%", " " + dataVals[0][13]);
          body.replaceText("%Emergency%", " " + dataVals[0][14]);
          body.replaceText("%IncidentDescription%", " " + dataVals[0][15]);
          body.replaceText("%IncidentType%", " " + dataVals[0][16]);
          body.replaceText("%FollowupComments%", dataVals[0][17]);
          body.replaceText("%LibrarianComments%", dataVals[0][18]);
          body.replaceText("%LibrarianInitials/Date%", dataVals[0][19]);
          body.findText("%ReportURL%").getElement().asText().setText("Edit Report").setLinkUrl(dataVals[0][20]).setFontSize(14).setBold(true).setBackgroundColor('#fff2cc');
          report.saveAndClose();
      // send an email to reporter and their supervisors
      var email = emailTemplate;
      var subject = "Incident Report Updated - " + dataVals[0][3];
          email = email.replace('%NAME%', dataVals[0][1]);
          email = email.replace('%RESPONSEDATE%', dataVals[0][0]);
          email = email.replace('%DATE%', fIncidentDate);
          email = email.replace('%TIME%', fIncidentTime);
          email = email.replace('%LIBRARY%', dataVals[0][3]);
          email = email.replace('%INCIDENTREPORT%', dataVals[0][20]);
          email = email.replace('%LOCATION%', dataVals[0][6]);
          email = email.replace('%STAFF%', dataVals[0][7]);
          email = email.replace('%CONTACTS%', dataVals[0][8]);
          email = email.replace('%AGE%', dataVals[0][9]);
          email = email.replace('%MINORAGE%', dataVals[0][10]);
          email = email.replace('%PARENT%', dataVals[0][11]);
          email = email.replace('%OTHERCONTACT%', dataVals[0][12]);
          email = email.replace('%CONTACTINFO%', dataVals[0][13]);
          email = email.replace('%EMERGENCY%', dataVals[0][14]);
          email = email.replace('%DESCRIPTION%', dataVals[0][15]);
          email = email.replace('%INCIDENTTYPE%', dataVals[0][16]);
          email = email.replace('%LIBCOMMENTS%', dataVals[0][18]);
          email = email.replace('%FOLLOWUP%', dataVals[0][17]);
          email = email.replace('%DOCURL%', dataVals[0][21]);
      MailApp.sendEmail(emailList, subject, "", {htmlBody: email, cc: ccEm});
      }   
  }
//set installable trigger EDIT>Current Project's Triggers | Run: getUrl ; Events: From spreadhseet - On form submit
function getUrl(e) {
  var responseSheet = e.range.getSheet();
  var row = e.range.getRow();
  var time = responseSheet.getRange(row, 1).getValue(); //get timestamp
  var timeStamp = new Date(time);
  var responseColumn = 21; // Column where the edit form response URL is recorded.
  var googleFormId = 'GoogleFormId_xXXXXxyXyyxy'; //form ID from form edit URL
  // Get the Google Form linked to the response
  var googleForm = FormApp.openById(googleFormId);
  // Get the form response based on the timestamp
  var formResponse = googleForm.getResponses(timeStamp).pop();
  // Get the Form response URL and add it to the Google Spreadsheet
  var responseUrl = formResponse.getEditResponseUrl();
  responseSheet.getRange(row, responseColumn).setValue(responseUrl);
  // Make a Doc Copy & Email
    IncidentReportTemplateCopy(responseSheet, row);
}