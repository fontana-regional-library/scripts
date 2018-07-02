<h1>Creating Training Certificates with Google Apps Script</h1>
<h2>Collecting Data with Google Forms</h2>
<p>Data is collected with a Google Form. This form is completed by staff when they attend a training session or webinar &amp; asks for the name of the training, when the training was completed, web address, provider, and feedback about training quality &amp; goals.</p>
<p>This collected form data is the "Master repository" for training data.</p>
<p>This form collects email address automatically &amp; is restricted to domain users. Repsonse receipts are sent on request, and users can submit another form after submission.</p>
<img src="https://github.com/fontana-regional-library/scripts/blob/master/Training%20Certificates/imgs/certificate-data.jpg?raw=true"/>
<h2>Preparing Training Data for use in Certficate Creation</h2>
<p>A separate google sheet imports data from the master form response sheet, including staff email, the title of the training, and the date the training was completed, if a certificate has not be sent (using NULL as indicator that certificate hasn't been sent)</p>
<p>Cell A3:<br/>
<pre><code>=(QUERY(IMPORTRANGE("XXXxxxXXXXxxxxxxxxxFileID-ofMasterFormResponseSheet","Form Responses 1!B2:P"),"SELECT Col1, Col2, Col3 WHERE Col15 IS NULL"))</code></pre></p>
<p>Lookup staff name &amp; supervisor's email based on trainee's email address:<br/>
    <pre><code>=IFERROR(ARRAYFORMULA(VLOOKUP(A3:A,directory!A2:C,2,FALSE)),"")</code></pre><br/><br/>
    <pre><code>=IFERROR(ARRAYFORMULA(VLOOKUP(A3:A,directory!A2:C,3,FALSE)),"")</code></pre></p>
    <img src="https://github.com/fontana-regional-library/scripts/blob/master/Training%20Certificates/imgs/certificate-data.JPG?raw=true"/>
<h3>Staff Directory Sheet</h3>
<p>While the form automatically collects email address upon submission, we tried to streamline data collection &amp; avoid input repetition by creating a staff directory sheet that lists the unique emails from the master sheet, and then is populated manually with user's name &amp; supervisor email address.</p>
<img src="https://github.com/fontana-regional-library/scripts/blob/master/Training%20Certificates/imgs/directory.PNG?raw=true"/>
<p>Cell G3 - Email Template</p>
<pre><code>This digital Certificate of Completion is awarded to %NAME% for completing "%TITLE%" on %DATE%. For a complete listing of completed trainings, please visit &lt;a href="ADDRESS/URL"&gt;your profile in the Staff Activity Portal&lt;a&gt; or the &lt;a href="ADDRESS/URL"&gt;Staff Directory&lt;/a&gt;. &lt;br/&gt;&lt;br/&gt;&lt;img style="width:100%;max-width:650px;" src="cid:%BLOB%"&gt;</code></pre>
<h3>Certificate Template</h3>
<p>Create a google slides file to design your certificate (You can use something like Canva, Photoshop, etc to design your "Background image") and insert placeholder textboxes where name, training title, date, etc should go.</p>
<img src="https://github.com/fontana-regional-library/scripts/blob/master/Training%20Certificates/imgs/certificate.PNG?raw=true"/>
<p>Using Google Apps scripts, we will programatically make a copy of this template, fill out the user data, save as an image file/pdf and email to the trainee and their supervisor.</p>
<p>We also have created a folder in Google Drive where all the generated pdf certificates are saved.</p>
<h2>Google Apps Script</h2>
<p>With the google form, google sheets, google slide certificate template, etc ready, we can use the <a href="https://github.com/fontana-regional-library/scripts/blob/master/Training%20Certificates/Training%20Certificate%20sheet%20script.js">Google Apps Script to semi-automate certificate creation</a>.</p>
<p>Using the script (in the menu on your certificate spreadsheet, go to "Tools>Script Editor"), we create a "Task Menu" which creates a new menu item on our spreadsheet. When clicking "Create and Email," the script will loop through the data, create a new certificate, save a pdf and image version of their certificate, email their certificate to them and their supervisor, and then delete some files for cleanup. (we could script this to happen on form submit, but we've added the button trigger to allow for proofreading before certificates are sent).</p>