<h1>Incident Reports with Google Apps</h1>
<h2>Collecting Incident Reports</h2>
<p>A Google Form collects data about any incident (accidents, post-emergency reports, behavioral incidents, etc). The form could be setup to automatically collect email address, but to encourage quick reporting from any computer the information about the person submitting the report is user submitted.</p>
<h2>Storing Responses in Google Sheets</h2>
<p>The form responses are stored in a Google Sheet, which includes a 2nd sheet that stores "settings." For us, this simply includes a list of supervisors for each branch which should be sent a copy of the submitted information, along with a list of staff who should be cc'd on all reports (HR and Library Director) and the HTML email template with placeholders.</p>
<p>A formula in the Form Responses sheet populates a column/cell with the list of supervisors for the reported incident who should be notified (vlookup of library location):<br/>
<code>=IFERROR(ARRAYFORMULA(VLOOKUP(D2:D,Settings!A2:B9,2,FALSE)),"")</code></p>
<img src="https://github.com/fontana-regional-library/scripts/blob/master/Incident%20Reports/imgs/report-settings.PNG?raw=true"/>
<h3>Script to send notifications and Create Google Doc</h3>
<p>The <a href="https://github.com/fontana-regional-library/scripts/blob/master/Incident%20Reports/incidentreportscript.js">script</a> is added to the google sheet and is triggered whenever a form is submitted or edited. (run "myFunction" to intiate the script)</p>
<p>The onFormSumit function adds the URL to edit the form response (included in the email notifications so staff can add more info or clarifications).</p>
<p>The IncidentReportTemplateCopy function creates a Google Doc with the form responses and sends email notification when a new report has been submitted or when an existing report has been updated.</p>
<h4>The Google Doc Template</h4>
<img src="https://github.com/fontana-regional-library/scripts/blob/master/Incident%20Reports/imgs/report-template-ex.png?raw=true"/>
<p>The Google Doc template is laid out in two tables. The first table is the template for the form responses. The 2nd table is a space for other notes (long-term followup, adding links to other related incident reports, etc.)</p>
<p>When a form is submitted, a copy of the template is created and added to a specified folder in Google Drive. The script replaces the placeholder text with the form data. A link to the newly created Google Doc is added to the Form Responses sheet with the submitted form data.</p>
<p>When a form is edited, the Google Doc created when the form was submitted is edited. The table that is the "form responses template" is deleted and re-written with the edited form data, while the "notes" table remains the same. When edited, a version history in google docs is created, so that old responses can also be seen or reviewed.</p>
<h4>Email Notifications</h4>
<p>The <a href="https://github.com/fontana-regional-library/scripts/blob/master/Incident%20Reports/email-template.html">email template</a> also contains placeholders that are replaced with the form responses. This email is sent to the person who submitted the report, as well as the supervisors for the location where the incident occurred, HR, and the library director. The emails contains a copy of the submitted form, a link to edit the form responses, as well as a link to the created Google Doc.</p>
<img src="https://github.com/fontana-regional-library/scripts/blob/master/Incident%20Reports/imgs/report-email.PNG?raw=true"/>

