function processEmail(processType='draftCombinedNew') {
  const sheetName = "MailApp"
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();

  const mail_subject = sheet.getRange("H2").getValue();
  const mail_salutation = sheet.getRange("H3").getValue();
  const mail_body_para1 = sheet.getRange("H4").getValue();
  const mail_body_para2 = sheet.getRange("H5").getValue();
  const mail_body_para3 = sheet.getRange("H6").getValue();
  const post_script = sheet.getRange("H7").getValue();
  
  // Template for email subject and body
  var mailBody = prepareBody(mail_salutation, mail_body_para1, mail_body_para2, mail_body_para3, post_script)
  if (processType == 'draftSeparateNew')
    separateNewEmails(data, mail_subject, mailBody, sendMailFlag=false, draftMailFlag=true);
  else if (processType == 'sendSeparateNew')
    separateNewEmails(data, mail_subject, mailBody, sendMailFlag=true, draftMailFlag=false);
  else if (processType == 'draftCombinedNew')
    combinedNewEmails(data, mail_subject, mailBody, sendMailFlag=false, draftMailFlag=true);
  else if (processType == 'sendCombinedNew')
    combinedNewEmails(data, mail_subject, mailBody, sendMailFlag=true, draftMailFlag=false);
  else if (processType == 'draftSeparateForward')
    separateForwardEmails(data, mail_subject, mailBody, sendMailFlag=false, draftMailFlag=true)
  else if (processType == 'sendSeparateForward')
    separateForwardEmails(data, mail_subject, mailBody, sendMailFlag=true, draftMailFlag=false)
  else if (processType == 'draftCombinedForward')
    combinedForwardEmails(data, mail_subject, mailBody, sendMailFlag=false, draftMailFlag=true)
  else if (processType == 'sendCombinedForward')
    combinedForwardEmails(data, mail_subject, mailBody, sendMailFlag=true, draftMailFlag=false)
}

// For Combined New Mails
function combinedNewEmails(data, mail_subject, mailBody, sendMailFlag=false, draftMailFlag=false){
  var emailList = [];
  var flatsConsidered = [];
  for (let i = 1; i < data.length; i++) {
    const flatNo = data[i][0]    
    const sendmail = data[i][4];
    if (sendmail.toLowerCase() == "yes" && (data[i][2] || data[i][3])){        
      if (data[i][2])
        emailList.push(data[i][2]) // email1
      if (data[i][3])
        emailList.push(data[i][3]); // email2
      
      flatsConsidered.push(flatNo) // add flat
    }
  }
  // Replace placeholders in subject and body
  var subject = mail_subject;
  var body = mailBody;

  // Attach Signature
  body = body+getSignature();
  
  var uniqueEmailSet = new Set(emailList); // take unique list
  emailList = Array.from(uniqueEmailSet)
  // Create draft (use 'to' to populate the draft email address field)
  if (emailList.length > 0)
  {
      var msg = "";
      recipientBatches = splitEmailsIntoBatches(emailList)  ;    
      recipientBatches.forEach((recipientBatch, index) => {
        if (draftMailFlag){
          GmailApp.createDraft(recipientBatch.join(","), subject, "", { htmlBody: body });
          msg = "Mail drafted for "+(emailList.length)+" emails (flats:"+flatsConsidered.length+") in "+recipientBatches.length+" batches. ["+flatsConsidered.join(",")+"]";
        }
        if (sendMailFlag){
          //GmailApp.sendEmail(recipientBatch.join(","), subject, "", { htmlBody: body }); // IMPORTANT !! Be cautious while uncommenting. It will send mail to recipients
          msg = "Mail sent to "+(emailList.length)+" emails (flats:"+flatsConsidered.length+") in "+recipientBatches.length+" batches. ["+flatsConsidered.join(",")+"]"          
        }
      });      
      Logger.log(msg)
  }    
}

// For Separate New Mails
function separateNewEmails(data, mail_subject, mailBody, sendMailFlag=false, draftMailFlag=false){
  for (let i = 1; i < data.length; i++) {
    const variablesToReplace = {};
    variablesToReplace["{{FlatNo}}"] = data[i][0];
    variablesToReplace["{{OwnerNames}}"] = data[i][1];
    var emailList = [];
    if (data[i][2])
      emailList.push(data[i][2]) // email1
    if (data[i][2])
      emailList.push(data[i][3]); // email2
    
    const sendmail = data[i][4];
    var msgToLog = "{{OwnerNames}}, Flat no. {{FlatNo}}";
    if (sendmail.toLowerCase() == "yes"){
      // Replace placeholders in subject and body
      var subject = mail_subject;
      var body = mailBody;

      for (var variable in variablesToReplace) {
        
        subject = subject.replaceAll(variable, variablesToReplace[variable]);
        body = body.replaceAll(variable, variablesToReplace[variable]);
        msgToLog = msgToLog.replaceAll(variable, variablesToReplace[variable]);
      }
      // Attach Signature
      body = body+getSignature()
      
      // Create draft (use 'to' to populate the draft email address field)
      if (emailList.length > 0)
      {
          var recipients = emailList.join(",");  // Convert array to comma-separated string
          if (draftMailFlag){
            GmailApp.createDraft(recipients, subject, "", { htmlBody: body });
            Logger.log("Mail drafted for "+msgToLog);            
          }
          if (sendMailFlag){
            //GmailApp.sendEmail(recipients, subject, "", { htmlBody: body }); // IMPORTANT !! Be cautious while uncommenting. It will send mail to recipients
            Logger.log("Mail sent to "+msgToLog);
          }
      }
    }
  }
}

// For Combined Forward Mails
function combinedForwardEmails(data, mail_subject, mailBody, sendMailFlag = false, draftMailFlag = false) {
  var threads = GmailApp.search('subject:"' + mail_subject + '"');

  if (threads.length > 0) {
    var latestThread = threads[0];
    var messages = latestThread.getMessages();
    var latestMessage = messages[messages.length - 1];

    // Attach Signature
    body = mailBody + getSignature()
    var originalBody = latestMessage.getBody();
    var combinedBody = body + originalBody;
    var draftSubject = "Fwd: " + mail_subject;

    // Get Emails
    var emailList = [];
    var flatsConsidered = [];
    for (let i = 1; i < data.length; i++) {
      const flatNo = data[i][0]
      const sendmail = data[i][4];
      if (sendmail.toLowerCase() == "yes" && (data[i][2] || data[i][3])) {
        if (data[i][2])
          emailList.push(data[i][2]) // email1
        if (data[i][2])
          emailList.push(data[i][3]); // email2

        flatsConsidered.push(flatNo) // add flat
      }
    }
    var uniqueEmailSet = new Set(emailList); // take unique list
    emailList = Array.from(uniqueEmailSet)
    if (emailList.length > 0) {
      var msg = "";
      recipientBatches = splitEmailsIntoBatches(emailList);
      recipientBatches.forEach((recipientBatch, index) => {
        if (draftMailFlag) {
          GmailApp.createDraft(recipientBatch.join(","), draftSubject, "",{ htmlBody: combinedBody });
          msg = "Mail drafted for "+(emailList.length)+" email ids (flats:"+flatsConsidered.length+") in "+recipientBatches.length+" batches. ["+flatsConsidered.join(",")+"]";
        }
        else if (sendMailFlag) {
          //GmailApp.sendEmail(recipientBatch.join(","), draftSubject, "",{ htmlBody: combinedBody});
          msg = "Mail sent to "+(emailList.length)+" email ids (flats:"+flatsConsidered.length+") in "+recipientBatches.length+" batches. ["+flatsConsidered.join(",")+"]"          
        }
      });      
      Logger.log(msg)
    }
  } else {
    Logger.log("No email found with the given subject.");
  }
}

// For Separate Forward Mails
function separateForwardEmails(data, mail_subject, mailBody, sendMailFlag=false, draftMailFlag=false) {
  const variablesToReplace = {};
  for (let i = 1; i < data.length; i++) {
    variablesToReplace["{{FlatNo}}"] = data[i][0];
    variablesToReplace["{{OwnerNames}}"] = data[i][1];
    var emailList = [];
    if (data[i][2])
      emailList.push(data[i][2]) // email1
    if (data[i][2])
      emailList.push(data[i][3]); // email2

    const sendmail = data[i][4];
    var msgToLog = "{{OwnerNames}}, Flat no. {{FlatNo}}";
    if (sendmail.toLowerCase() == "yes") {
      // Replace placeholders in subject and body
      var subject = mail_subject;
      var body = mailBody;
      for (var variable in variablesToReplace) {

        subject = subject.replaceAll(variable, variablesToReplace[variable]);
        body = body.replaceAll(variable, variablesToReplace[variable]);
        msgToLog = msgToLog.replaceAll(variable, variablesToReplace[variable]);
      }
      var threads = GmailApp.search('subject:"' + subject + '"');
      if (threads.length > 0) {
        var latestThread = threads[0];
        var messages = latestThread.getMessages();
        var latestMessage = messages[messages.length - 1];

        // Combine custom message with the original
        var originalBody = latestMessage.getBody();
        var combinedBody = body + getSignature() + originalBody;
        var draftSubject = "Fwd: " + subject;
        
        // Create draft (use 'to' to populate the draft email address field)
        if (emailList.length > 0) {
          var recipients = emailList.join(",");  // Convert array to comma-separated string
          if (draftMailFlag) {
            GmailApp.createDraft(recipients, draftSubject, "", { htmlBody: combinedBody });
            Logger.log("Mail drafted for " + msgToLog);
          }
          if (sendMailFlag) {
            //GmailApp.sendEmail(recipients, draftSubject, "", { htmlBody: combinedBody }); // IMPORTANT !! Be cautious while uncommenting. It will send mail to recipients
            Logger.log("Mail sent to " + msgToLog);
          }
        }
      }
      else {
        Logger.log("No email found with the given subject for "+msgToLog);
      }      
    }    
  }  
}

function getSignature(){
  var signature = `<font color="#666666" face="arial, sans-serif">
    Thanks &amp; Regards,<br>
    Regent Park CHSL Committee,<br>
    Baner, Pune-411045
  </p>
  `;

  return signature;
}

function prepareBody(mail_salutation, mail_body_para1, mail_body_para2, mail_body_para3, post_script){
  var htmlBodyTemplate = `
    <div style="font-family: Arial, sans-serif; font-size: 14px;">`+mail_salutation;
      
      if (mail_body_para1 != ""){
        htmlBodyTemplate +=  "<p>"+mail_body_para1+"<p>";
      }
      if (mail_body_para2 != ""){
        htmlBodyTemplate +=  "<p>"+mail_body_para2+"<p>";
      }
      if (mail_body_para3 != ""){
        htmlBodyTemplate +=  "<p>"+mail_body_para3+"<p>";
      }
      if (post_script != ""){
        htmlBodyTemplate +=  "<p><i><strong>PS. </strong>"+post_script+"<i><p>";
      }
    
    htmlBodyTemplate += "</div>";
  ;

  return htmlBodyTemplate;
}

function splitEmailsIntoBatches(emailList) {
  const batchSize = 50; // Set your desired batch size
  const batches = [];
  
  for (let i = 0; i < emailList.length; i += batchSize) {
    const batch = emailList.slice(i, i + batchSize);
    batches.push(batch);
  }
  return batches;
}