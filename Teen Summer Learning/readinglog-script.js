// Adding vars using the "Question" title from the google form
var tEmail = "Participant's E-mail";
var tName = "Participant's First and Last Name";
var tMinutes = "Minutes spent Reading";
var ss = SpreadsheetApp.getActiveSpreadsheet();
var dataSheet = ss.getSheetByName('Form Responses 1');
var emailTempl = dataSheet.getRange("A1").getValue(); //HTML email template

//RUNS on form submit
    function onFormSubmit(e) {
        Logger.log('on Form Submit');
        Logger.log(e);
        Logger.log(e.namedValues[tEmail]);
        var em = e.namedValues[tEmail][0];
        var nm = e.namedValues[tName][0];
        var minR = e.namedValues[tMinutes][0];
            Logger.log(em);
        var msg = emailTempl;
    //Replace placeholder text with values from the submitted form
            msg = msg.replace('%NAME%', nm);
            msg = msg.replace('%MINUTES%', minR);
            Logger.log('sendToMail');
        MailApp.sendEmail(em, "Way to go! Your reading time has been logged", "", {htmlBody: msg, cc: "maconteens@fontanalib.org"});
            Logger.log('Sent:' + msg);
}
/**
 * Creates onFormSubmit trigger.
 * ***May be able to remove this function and set trigger manually via menu "Edit>Current Project's Triggers> Add a new Trigger"***
 */
function myFunction(){
  var ss = SpreadsheetApp.getActive();
  var a = ScriptApp.newTrigger("onFormSubmit");
  var b = a.forSpreadsheet(sheet);
  var c = b.onFormSubmit();
  var d = c.create();
}