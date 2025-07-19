function processEmail(processType='draftCombined') {
  const sheetName = "EmailApp"
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
  if (processType == 'draftSeparate')
    separateEmails(data, mail_subject, mailBody, false, true);
  else if (processType == 'sendSeparate')
    separateEmails(data, mail_subject, mailBody, true, false);
  else if (processType == 'draftCombined')
    combinedEmails(data, mail_subject, mailBody, false, true);
  else if (processType == 'sendCombined')
    combinedEmails(data, mail_subject, mailBody, true, false);
}

// For Combined Mails
function combinedEmails(data, mail_subject, mailBody, sendMailFlag=false, draftMailFlag=false){
  var emailList = [];
  var flatsConsidered = []
  for (let i = 1; i < data.length; i++) {
    const flatNo = data[i][0]    
    const sendmail = data[i][4];
    if (sendmail.toLowerCase() == "yes" && (data[i][2] || data[i][3])){        
      if (data[i][2])
        emailList.push(data[i][2]) // email1
      if (data[i][2])
        emailList.push(data[i][3]); // email2
      
      flatsConsidered.push(flatNo) // add flat
    }
  }
  // Replace placeholders in subject and body
  var subject = mail_subject;
  var body = mailBody;

  // Attach Signature
  body = body+getSignature()
  
  // Create draft (use 'to' to populate the draft email address field)
  if (emailList.length > 0)
  {
      var recipients = emailList.join(",");  // Convert array to comma-separated string
      if (draftMailFlag){
        GmailApp.createDraft(recipients, subject, "", { htmlBody: body });
        Logger.log("Mail drafted for "+(flatsConsidered.length)+" members. ["+flatsConsidered.join(",")+"]");
      }
      if (sendMailFlag){
        //GmailApp.sendEmail(recipients, subject, "", { htmlBody: body }); // IMPORTANT !! Be cautious while uncommenting. It will send mail to recipients
        Logger.log("Mail sent to "+(flatsConsidered.length)+" members. ["+flatsConsidered.join(",")+"]");
      }
  }    
}

// For Separate Mails
function separateEmails(data, mail_subject, mailBody, sendMailFlag=false, draftMailFlag=false){
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