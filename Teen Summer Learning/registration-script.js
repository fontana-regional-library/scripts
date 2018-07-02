// MailChimp API Key from https://us3.admin.mailchimp.com/account/api/
var API_KEY = 'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx-us##';
// get the right server (us##) & list id (found under list settings- "Unique id for list NAME" NOT in list URL)
var mc_base_url = 'https://us##.api.mailchimp.com/3.0/lists/XXXXXXyyyyy/members';

// Pick a subscription type for users added via this script: 
// subscribed (no opt in confirmation)
// unsubscribed 
// cleaned 
// pending (triggers double optin) 
// transactional
var status = 'subscribed';

// Adding vars using the "Question" title from the google form
var fieldEmail = "Participant's E-mail";
var parOptin = "Email Newsletter";
var parName = "First & Last Name";
var guarEmail = "Parent E-mail ";
var ss = SpreadsheetApp.getActiveSpreadsheet();
var dataSheet = ss.getSheetByName('Form Responses 1');
var emailTemplate = dataSheet.getRange("A1").getValue(); //HTML email template
/**
 * Uses the MailChimp 3.0 API to add a subscriber to a list.
 */
    function sendToMailChimp(em){
        Logger.log('sendToMailChimp');
        var payload = {
            "email_address": em,
            "status": status
        };
    //replace username@fontanalib.org with your mailchimp username
        var headers = {
            "Authorization": 'Basic ' + Utilities.base64Encode('USERNAME@fontanalib.org:' + API_KEY, Utilities.Charset.UTF_8)
        };
        var options = {
            "method": "POST",
            "payload": JSON.stringify(payload),
            "headers": headers,
            "muteHttpExceptions" : true
        };
        Logger.log(options);
    // send data to MailChimp
        try {
            var response = UrlFetchApp.fetch(mc_base_url,options);
        // Log response
            if(response.getResponseCode() === 200) {
                Logger.log('Success'); // It worked!
                Logger.log(response);
            } else{
                Logger.log('Issues');
                Logger.log(response);
            }
            } catch (err) {
                Logger.log('Error');
                Logger.log(err);
            }
        }
/**
 * Runs on Form Submit
 * Emails Receipt, Checks for Adding to MailChimp
 */
    function onFormSubmit(e) {
        var em = e.namedValues[fieldEmail][0];
        var optinA = e.namedValues[parOptin][0];
        var gEm = e.namedValues[guarEmail][0];
        var nm = e.namedValues[parName][0];
        var address = em + "," + gEm;
    //Check if Opt-in to receiving emails and email address exists, send to MailChimp
        if (optinA == "Yes" && em.length){
            sendToMailChimp(em);
        }else{
            Logger.log('Error: couldnt find an email address in submission');
    }
        var email = emailTemplate;
    // fills in placeholder text in email and sends registration receipt
        email = email.replace('%NAME%', nm);
    MailApp.sendEmail(address, "Thank You for Registering for Teen Summer Learning Program at the Library", "", {htmlBody: email, cc: "maconteens@fontanalib.org"});    
    }
/**
 * Creates onFormSubmit trigger.
 * ***May be able to remove this function and set trigger manually via menu "Edit>Current Project's Triggers> Add a new Trigger"***
 */
function myFunction(){
  // Was separated line by line for debugging purposes.
  var sheet = SpreadsheetApp.getActive();
  var a = ScriptApp.newTrigger("onFormSubmit");
  var b = a.forSpreadsheet(sheet);
  var c = b.onFormSubmit();
  var d = c.create();
}