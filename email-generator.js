/*
Project: Email Generator
Author: Joshua Bays, BSCPE UAH Class of 2026
  jb0401@uah.edu
  joshua_bays@protonmail.com
Description:
  This project takes formatted data from a spreadsheet, and creates a personalized email for a list of provided recipients.
  Emails can include variables that will be replaced
*/

/* Add menu items */
function onOpen() {
  let ui = SpreadsheetApp.getUi();
  let menu = ui.createMenu("Email Functions");
  menu.addItem("Test Send", "test_email").addToUi(); /* Add button for test send function */
  menu.addItem("Send Emails", "send_emails").addToUi(); /* Add button for seand all function */
}

/* Send one test email (1st row of recipient data) */
function test_email(){
  generate_emails("test");
}

/* Send all emails */
function send_emails(){
  generate_emails("all");
}

/* Function used to send emails, includes option for a test send */
function generate_emails(option){
  /* Open the sheet */
  let baseSheet = SpreadsheetApp.getActiveSpreadsheet();
  // let baseSheet = SpreadsheetApp.openById(spreadSheetId); /* Alternative option to open by ID */
  
  /* Create the tempalate values */
  let emailBody = ""; /* Main body and signature */
  let emailSubject = ""; /* Subject line */
  let emailAttach = []; /* List of attachments */
  let emailImages = {}; /* List of images */
  let emailLabel = null;
  let emailCC = "";
  let emailBCC = "";
  let sigBreak = 0; /* Check if line break before signature line has been created */

  /* Open the template writeup and get the data values */
  let writeupValues = baseSheet.getSheetByName("Email Writeup").getDataRange().getDisplayValues();
  let rowCt = writeupValues.length;
  
  /* Read through the entire writeup and create the template */
  for(let i = 0; i < rowCt; i++){
    switch(writeupValues[i][0]){
      case "SUBJECT":
        /* Create the subject */
        emailSubject = writeupValues[i][1];
        break;
      case "BODY":
        if(emailBody != ""){
          /* Add a line break before any body paragraph past the first */
          emailBody += "<br>";
        }
        /* Add a line break after any body paragraph */
        emailBody += writeupValues[i][1] + "<br>";
        break;
      case "SIGNATURE":
        if(!sigBreak){
          /* Add 2 line breaks before the first line of the signature */
          emailBody += "<br><br>";
          sigBreak = 1;
        }
        emailBody += writeupValues[i][1] + "<br>";
        break;
      case "BLANK":
        emailBody += "<br><br>"; /* Add line breaks */
        break;
      case "ATTACHMENT":
        if(writeupValues[i][1].search(/{[^}]+}/) >= 0){
          /* If the attachment contains a variable, add it to be replaced by the recipient data */
          emailAttach.push(writeupValues[i][1].substring(writeupValues[i][0].search(/{[^]+}/)));
        }else{
          /* Otherwise, take the id from the link and attach it as a blob (pdf) */
          emailAttach.push(DriveApp.getFileById(writeupValues[i][1].match(/[-\w]{25,}(?!.*[-\w]{25,})/)[0]).getBlob());
        }
        break;
      case "IMAGE":
        let w = writeupValues[i][1].match(/{[0-9]*}/g)[0].slice(1, -1);
        let h = writeupValues[i][1].match(/{[0-9]*}/g)[1].slice(1, -1);
        emailBody += "<br><img style='display:block; left-margin:auto; right-margin:auto' src='cid:"+"img"+writeupValues[i][1].match(/{[^}]*}/g)[0]+"'width='"+w+"'height='"+h+"'></br>";
        emailImages["img"+writeupValues[i][1].match(/{[^}]*}/g)[0]] = null;
        break;
      case "LABEL":
        let labelList = GmailApp.getUserLabels();
        for(let j = 0; j < labelList.length; j++){
          if(labelList[j].getName() == writeupValues[i][1]){
            emailLabel = GmailApp.getUserLabelByName(writeupValues[i][1]);
          }
        }
        break;
      case "CC":
        emailCC = writeupValues[i][1];
        break;
      case "BCC":
        emailBCC = writeupValues[i][1];
        break;
    }
  }
  /* Remove last newline character */
  if(emailBody[emailBody.length - 1] == '<br>'){ emailBody.slice(0, -1); }
  /* Get the recipient data */
  let recipientsSheet = baseSheet.getSheetByName("Recipients");
  let recipientsValues = recipientsSheet.getDataRange().getDisplayValues();
  rowCt = recipientsValues.length;
  
  /* Make the email customized for each recipient */
  for(let i = 1; i < rowCt; i++){
    /* Skip if an email has already been marked as sent */
    if(option != "test" && recipientsValues[i][1] != ""){
      continue;
    }
    
    /* Create recipient-specific elements that start as copys of the template values */
    let recipientBody = emailBody;
    let recipientSubject = emailSubject;
    let recipientCC = emailCC;
    let recipientBCC = emailBCC;
    let recipientImages = emailImages;
    let imgReplace = 0;
    
    /* Get the recipient's email from the recipient data */
    let recipientEmail = recipientsValues[i][0];
    
    /* Replace the variables of recipient elements with the recipient data */
    for(let j = 2; j < recipientsValues[i].length; j++){
      imgReplace = 0;
      
      /* Replace images */
      for(let k = 0; k < Object.keys(recipientImages).length; k++){
        if("img"+recipientsValues[0][j] == Object.keys(recipientImages)[k]){
          recipientImages["img"+recipientsValues[0][j]] = DriveApp.getFileById(recipientsValues[i][j].match(/[-\w]{25,}(?!.*[-\w]{25,})/)[0]).getBlob();
          imgReplace = 1;
        }
      }
      
      /* Update subject and body */
      if(!imgReplace){
        /* Replace any links. Links are formatted as {varName1}_{varName2} (Ex: {spreadsheetlink}_{url} (Row 1 declaration) {UAH's website}_{uah.edu} (recipient declaration)) */
        if(recipientsValues[0][j].search(/{[^}]*}_{[^}]*}/) > -1){
          recipientBody = recipientBody.replaceAll(recipientsValues[0][j], "<a href=\""+recipientsValues[i][j].match(/_{[^}]*}/)[0].slice(2, -1)+"\">"+recipientsValues[i][j].slice(1,  recipientsValues[i][j].search("}"))+"</a>");
        }
        recipientBody = recipientBody.replaceAll(recipientsValues[0][j], recipientsValues[i][j]);
        recipientSubject = recipientSubject.replaceAll(recipientsValues[0][j], recipientsValues[i][j]);
        recipientCC = recipientCC.replaceAll(recipientsValues[0][j], recipientsValues[i][j]);
        recipientBCC = recipientBCC.replaceAll(recipientsValues[0][j], recipientsValues[i][j]);
      }

      /* Replace the variables within the attachment lists with blobs taken from the recipient data */
      for(let k = 0; k < emailAttach.length; k++){
        /* Check if an attachment item is a variable */
        if(recipientsValues[0][j] == emailAttach[k].toString()){
          /* Replace the variable with a blob generated with from the id found within the link in recipient data */
          emailAttach[k] = DriveApp.getFileById(recipientsValues[i][j].match(/[-\w]{25,}(?!.*[-\w]{25,})/)[0]).getBlob();
        }
      }
    }
    
    /* If the test option was passed, replace the recipient email with the specified test email */
    if(option == "test"){ recipientEmail = baseSheet.getRange("Instructions!C2").getValue(); }

    /* Send the email. */
    /* The sendEmail function should ignore any empty lists or strings, but if a problem comes up, look here. */
    MailApp.sendEmail({
      to: recipientEmail,
      subject: recipientSubject,
      cc: recipientCC,
      bcc: recipientBCC,
      htmlBody: recipientBody,
      attachments: emailAttach,
      inlineImages: emailImages,
    });
    
    /* If the test option was passed, do not send any more emails and do not log any emails as sent*/
    if(option == "test"){ break; }
    
    if(emailLabel != null){
      Utilities.sleep(5000);
      let emailList = GmailApp.search("to:"+recipientEmail+" subject:"+recipientSubject, 0, 1);
      emailList[0].addLabel(emailLabel);
    }
    
    /* Log the email as sent */
    recipientsSheet.getRange(i+1, 2).setValue(new Date());
  }
}

