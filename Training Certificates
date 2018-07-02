////////////////////////////////////////////////////////////////////////////////////
// Training certificate spreadsheet imports Google Form entries from separate sheet
// where a column "Training certificate sent" is not marked
////////////////////////////////////////////////////////////////////////////////////
//creates menu to begin emails/certificate creation
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Training Certificates')
      .addItem('Create & Email', 'TrainingCertificate')
      .addToUi();
}
function TrainingCertificate(){
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dataSheet = ss.getSheetByName('Certificate Data'); //name of spreadsheet with directory matched info (pulls in supervisor names/emails with vlookup)
    var numRows = dataSheet.getRange("'Certificate Data'!C1").getValue(); //formula counts non-blank rows 
    var emailTemplate = dataSheet.getRange("G3").getValue(); // HTML email template pasted into cell (must replace ADDRESS with the URL below - other placeholders replaced programatically)
    //This digital Certificate of Completion is awarded to %NAME% for completing "%TITLE%" on %DATE%.
    //For a complete listing of completed trainings, please visit <a href="http://ADDRESS">your profile in the Staff Portal<a> or the <a href="http://ADDRESS">Staff Directory</a>. 
    //<br/><br/><img style="width:100%;max-width:650px;" src="cid:%BLOB%">
	var dataRange = dataSheet.getRange(3, 1, numRows, 5);
	var objects = dataRange.getValues();
	var template = DriveApp.getFileById('10a7NCVh1eJ8h8gGmDG21woR1nz0YgqFgEfC-H32-QuU'); // the ID of the Google Slides file - TEMPLATE for certificates
	var folder = DriveApp.getFolderById("1Vi6abq3j9lEfvnhLpGEb33J5aXGgg8SA"); // the ID of the Folder where to save the PDF certificates

	for (var i=0; i < numRows; i++){
        var col = objects[i];
        var date = col[2];
        var fDate = Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), "MMMM d, YYYY"); //formats date for use in email
        var titleDate = Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), "YYMMdd"); //formats dats for use in file name
        var title = col[1]; //title of the training course
        var stEmail = col[0]; // trainee's email
        var stName = col[3];  //trainee's name
        var suEmail = col[4]; // supervisor's email
        var copyName = titleDate + " _ " + stName + " _ " + title; //name of file
        var certificate = template.makeCopy(copyName, folder); //creates new slide/certificate
        var fileid = certificate.getId(); 
        var file = SlidesApp.openById(fileid);
  // Create the text merge (replaceAllText) requests for this presentation.
        var slides = file.getSlides()[0];
            slides.replaceAllText("<<name>>",stName,true);
            slides.replaceAllText("<<title>>",title,true);
            slides.replaceAllText("<<date>>",fDate,true);
      file.saveAndClose();
//Get Image
        var url = 'https://docs.google.com/presentation/d/' + fileid + '/export/png?id=' + fileid + '&pageid=' + slides.getObjectId(); 
        var options = {
            headers: {
                Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
            }
        };
        var imgName = copyName + "img"
        var response = UrlFetchApp.fetch(url, options);
        var image = response.getAs(MimeType.PNG);
        image.setName(imgName);
        var thumbnail = folder.createFile(image);
        var tbId = thumbnail.getId();
        var tbUrl = 'cid:' + tbId;
        var address = '\"' + stEmail + "," + suEmail + '\"'; // concat supervisor email & trainee emails for mailto
         // create pdf version of certificates
        var pdf = DriveApp.getFileById(fileid).getBlob();
        folder.createFile(pdf);
        var attach = DriveApp.getFileById(fileid).getAs("application/pdf"); 
        // create email subject, replace body text in template to personalize, add images inline, & attach certificate files
        var emailSubject = "Training Completion Certificate - " + stName + " - " + title;
        var email = emailTemplate;
            email = email.replace('%DATE%', fDate);
            email = email.replace('%TITLE%', title);
            email = email.replace('%NAME%', stName);
            email = email.replace('%BLOB%', tbId);
        var inlineImages = {};
            inlineImages[thumbnail.getId()] = thumbnail.getBlob();
        MailApp.sendEmail(stEmail + "," + suEmail, emailSubject, "", {htmlBody: email, inlineImages:inlineImages, attachments: [attach, thumbnail]});
        //delete the personalized copy of the google slides template from drive folder
        DriveApp.getFileById(fileid).setTrashed(true);
        //delete the image copy of the certificate from drive folder
        DriveApp.getFileById(tbId).setTrashed(true);
    }
}
