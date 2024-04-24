# AppScript-Email-Generator
This is a generic copy of the Google App Script project I wrote for the University of Alabama in Huntsville [Enhanced Teaching and Learning Center (ETLC)](https://www.uah.edu/etl)

## Functionality
This project has been used to mass-send personalized emails containing embedded URLs, embedded images, and Google Drive attachments. In addition, emails sent by the generator can be automatically tagged per user specification.

## Using the Project
The Apps Script code was written as an extension of a Google Sheet. To do this, open a copy of the spreadsheet -> Extensions -> Apps Script, and import the code there.

Once a copy of the sheet is in use, a sheet contains all pertinent instructions as well as a place for the test email. Instructions for use have also been included below.

### Instructions
	1. Input list of recipients and their respecive emails in the "Recipients" tab
	2. Write the email body in the "Email Writeup" tab
		Use Column A to denote what type of email element is included (types listed at the bottom of this tab)
		Use Column B to include the associated text
			URLs to websites will automatically be converted to clickable links in the email once sent
			Separate each body paragraph into individual rows
		Use {name} as a placeholder to be replaced with the recipient's name in the email body
		Add any other variables to be replaced in the email using the same curly bracket notation as the {name} placeholder in the "Recipients" tab
			Include the respective added variables in the same tab
			Separate variables must be created for different capitalization schemes.
	3. [OPTIONAL] Select "Test Send" from the "Email Functions" menu to send the email to the address listed in C2 of this tab
		A sample email generated from the row 2 of the "Recipients" tab will be sent to the test email
	4. Select "Send Emails" from the "Email Functions" menu to send emails to all recipients and log send time

### Email Element Types
    SUBJECT; Subject of email. Default email signatures are not included in the email.
    BODY: Body paragraph of email. All instances will have one blank line space between them.
    SIGNATURE: Signature of email. The first instance will be spaced from the rest of the body, and all following instances will be spaced immediately below.
    BLANK: Additional blank line
    IMAGE: Image to be embeded within the email.
    ATTACHMENT: Attachments to the email. Every instance (Variables can be used) must be a Google Drive link, and will be converted to a pdf.
    CC: Email to carbon copy. If multiple emails are used, they must be separated by commas (Variables can be used) Only the last instance of CC will be used.
    BCC: Email to blind carbon copy. If multiple emails are used, they must be separated by commas (Variables can be used) Only the last instance of CC will be used.
    LABEL: Name of the label to catagorize the email. 

### Other Important Information
    The recipients sheet contains a column of when an email was sent to that recipent. If a value exists in that cell, the corresponding email will not be sent.
    Currently, embeded links can only be sent using a variable (A variable can be the same for all recipients). The format for a link variable is {varName1}_{varName2}.
        Ex: {spreadsheetlink}_{url} (Row 1 declaration) {the Google search engine}_{google.com} (recipient declaration)
    Embeded images must have 3 fields in the email writeup, and 1 in the recipients. Fields 2 and 3 specifiy width and height respectively, and can be blank.
        Ex (Email Writeup): {IMG1}_{600}_{400}
        Ex (Email Writeup): {IMG1}_{}_{400}
        Ex (Email Writeup): {IMG1}_{600}_{}
        Ex (Recipients): {IMG1} (Drive links for values)
    In order to label an email, the label must be first exist or be created from the Gmail client. Variables cannot be used.
    If a variable does not have a value for a recipient, it will be treated as blank text.


## Miscellaneous Information
1. [Apps Script reference page](https://developers.google.com/apps-script/reference)
2. If an error occurs, check file permissions.
3. The script will need to be authorized, this prompt will occur upon running the file.
4. If an email is not sending, check if it has already been marked as sent in the recipients sheet (or disable duplicate checking).
