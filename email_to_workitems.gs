// ===============================================
// TOKEN MANAGER
// ===============================================

function acessTokenGenerator() {
  var props = PropertiesService.getScriptProperties();
  var acessToken = props.getProperty("Acess_Token");
  var expiryTime = props.getProperty("Expiry_Time");

  if (acessToken && expiryTime && new Date().getTime() < parseInt(expiryTime)) {
    Logger.log("✅ Using Existing AcessToken");
    return acessToken;
  }

  Logger.log("🔄 Getting new AcessToken...");

  var refreshToken  = "1000.ae8748657d6a6ac3ea3f250fac0b5f3a.d8a89ddb22714c7fd1e27366c7723404";
  var clientId      = "1000.ZQNHDF0O4II870AQ5KSWE3SIPC2V3O";
  var clientSeceret = "4dc5aa9e5a5ba3261ba459a4f3d6fa66969f322d26";

  var response = UrlFetchApp.fetch("https://accounts.zoho.in/oauth/v2/token", {
    method: "post",
    payload: {
      refresh_token : refreshToken,
      client_id     : clientId,
      client_secret : clientSeceret,
      grant_type    : "refresh_token"
    },
    muteHttpExceptions: true
  });

  var data     = JSON.parse(response.getContentText());
  var newToken = data.access_token;

  if (!newToken) {
    Logger.log("🔴 Failed to get access token: " + response.getContentText());
    return null;
  }

  var expiry = new Date().getTime() + (55 * 60 * 1000); // 55 mins
  props.setProperty("Acess_Token", newToken);
  props.setProperty("Expiry_Time", expiry.toString());

  Logger.log("✅ New AcessToken saved");
  return newToken;
}


// ===============================================
// LABEL MANAGER
// ===============================================

function checkLabel(labelName) {
  var label = GmailApp.getUserLabelByName(labelName);
  if (!label) {
    label = GmailApp.createLabel(labelName);
    Logger.log("✅ Label created: " + labelName);
  }
  return label;
}
// ==============================================
// Acknowledgment Logic
// ==============================================

COMPANY_NAME = "CormSquare";
SUPPORT_EMAIL = "anycompany@gmail.com";

function sendAcknowledgemt(thread,from,subject,ticketId,ccHeader){

var nameMatch  = from.match(/^([^<]+)/);
var senderName = nameMatch ? nameMatch[1].trim() : "there";
 var plainBody =
 "Dear" + senderName +",\n\n"+
 "Thank you for reaching out. Your request has been received and a ticket has be created on behalf of the request sucessfully.\n\n"+
 "Ticket Id : " +ticketId+"\n"+
 "Subject : " +subject+"\n\n"+
 "Our team will review your request and get back to you as soon as possible.\n" +
    "Please keep your Ticket ID handy for any future follow-ups.\n\n" +
    "Warm regards,\n" +
    COMPANY_NAME + " Support Team\n" +
    SUPPORT_EMAIL;

  var htmlBody =
    "<div style='font-family:Arial,sans-serif;max-width:600px;color:#333;'>" +
      "<p>Dear <strong>" + senderName + "</strong>,</p>" +
      "<p>Thank you for reaching out to us. Your request has been received and a ticket has been raised successfully.</p>" +

      "<div style='background:#f4f6f8;border-left:4px solid #4a90e2;padding:16px 20px;margin:20px 0;border-radius:4px;'>" +
        "<p style='margin:0;font-size:13px;color:#666;'>YOUR TICKET ID</p>" +
        "<p style='margin:8px 0 4px;font-size:22px;font-weight:bold;color:#333;'> " + ticketId + "</p>" +
        "<p style='margin:0;font-size:13px;color:#555;'>Subject: <em>" + subject + "</em></p>" +
      "</div>" +

      "<p>Our team will review your request and get back to you as soon as possible.<br/>" +
      "Please keep your Ticket ID handy for any future follow-ups.</p>" +

      "<p>Warm regards,<br/>" +
        "<strong>" + COMPANY_NAME + " Support Team</strong><br/>" +
        "<a href='mailto:" + SUPPORT_EMAIL + "'>" + SUPPORT_EMAIL + "</a>" +
      "</p>" +

      "<hr style='border:none;border-top:1px solid #eee;margin:20px 0;'/>" +
      "<p style='font-size:11px;color:#aaa;'>This is an automated acknowledgment. Our support team will reach out to you shortly.</p>" +
    "</div>";
  
    var replyOptions = {
    htmlBody : htmlBody,
    name     : COMPANY_NAME + " Support",
    replyTo  : SUPPORT_EMAIL
  };

  if (ccHeader && ccHeader.trim() !== "") {
    replyOptions.cc = ccHeader;   // Gmail accepts the raw "Cc" header string directly
    Logger.log("📨 Sending acknowledgment CC to: " + ccHeader);
  }

  thread.reply(plainBody,replyOptions);


 // thread.reply(plainBody, {
 //   htmlBody : htmlBody,
 //   name : COMPANY_NAME + "Support",
  //  replyTo : SUPPORT_EMAIL
  //});

}

// ==============================================
// Generate Ticket Id
// ==============================================

function generateTicketId(itemId){
  
if (!itemId) {

  var datePart= Utilities.formatDate(new Date(), Session.getScriptTimeZone(),"yyyyMMdd");
  var randomPart = Math.floor(1000 +Math.random() * 9000);
  return "TKT-"+datePart+"-"+randomPart;
}
else {
  return "TKT - " + itemId;
}
}


// ===============================================
// WORK ITEM CREATOR
// ===============================================

function createWorkItemsFromUnreadEmails() {
 
  // --- Config ---
  var accessToken    = acessTokenGenerator();
  var teamId         = "60059184003";
  var projectId      = "48644000000173084";
  var sprintId       = "48644000000173090";
  var projItemTypeId = "48644000000173134";
  var projPriorityId = "48644000000173128";
 // var epicId         = "48644000000294099";

  var DEFAULT_EPIC_ID = "48644000000329039"


  if (!accessToken) {
    Logger.log("🔴 Aborting: No access token available");
    return;
  }

  // --- Label Setup ---
  var label = checkLabel("Work Item Created");

  // --- Fetch only unread emails WITHOUT the label ---
  var threads = GmailApp.search("is:unread -label:Work-Item-Created", 0, 50);
  Logger.log("📧 Threads to process: " + threads.length);

  if (threads.length === 0) {
    Logger.log("✅ No new emails to process");
    return;
  }

  // --- Process Each Thread ---
  for (var i = 0; i < threads.length; i++) {
    try {

      var messages = threads[i].getMessages();
      var message  = messages[messages.length - 1]; // get latest message in thread

      var subject = message.getSubject() || "[No Subject]";
      var body    = message.getBody();
      var from    = message.getFrom();
      var cc      = message.getHeader("Cc") || message.getHeader("CC") || "";

      Logger.log("CC:"+cc);

      Logger.log("📨 Processing (" + (i + 1) + "/" + threads.length + "): " + subject);
      var rawAttachments = message.getAttachments();
      Logger.log("📎 Raw attachments count: " + rawAttachments.length);
      var epicId = resolveEpicCc(message) || DEFAULT_EPIC_ID;
      Logger.log("📌 Using Epic ID: " + epicId + " for: " + subject);

      body = body.replace(/<style[\s\S]*?<\/style>/gi, "").replace(/<script[\s\S]*?<\/script>/gi, "");

      // filter attchmenrs before API Call 
      var attachments = getFilteredAttachments(message);

      // --- Build API URL ---
      var url = "https://sprintsapi.zoho.in/zsapi/team/"
        + teamId      + "/projects/"
        + projectId   + "/sprints/"
        + sprintId    + "/item/"
        + "?projitemtypeid=" + projItemTypeId
        + "&projpriorityid=" + projPriorityId
        + "&epicid=" + epicId ;

      // --- Build Payload ---
      var payload = {
        name        : subject,
        description : "<b>From:</b> " + from + "<br><br>" + body
      };

      // --- API Call ---
      var options = {
        method            : "post",
        headers           : { Authorization: "Zoho-oauthtoken " + accessToken },
        payload           : payload,
        muteHttpExceptions: true
      };

      var response     = UrlFetchApp.fetch(url, options);
      var responseCode = response.getResponseCode();
      var responseData = JSON.parse(response.getContentText());

      Logger.log("📬 API Response: " + JSON.stringify(responseData));

      // --- On Success: Mark Read + Apply Label ---
      if (responseCode === 200 || responseCode === 201) {

       // var zohoItemId = responseData && responseData.data && responseData.data[0]
       //                  ? responseData.data[0].globalKey || responseData.data[0].addedItemId
      //                   : null;



       // ── Extract Zoho Item ID for attachment upload ──
        var zohoItemId = null;
        if (responseData.itemNo) {
          zohoItemId = responseData.itemNo;
        } else if (responseData.data && responseData.data[0]) {
          zohoItemId = responseData.data[0].globalKey || responseData.data[0].itemId;
        }

         // ── Also get the raw addedItemId for the attachment API ──
        var addedItemId = responseData.addedItemId || zohoItemId;

        var ticketId = generateTicketId(zohoItemId);

        // ── Upload attachments to the created work item ──
        if (addedItemId) {
          uploadAttachmentsToWorkItem(accessToken, teamId, projectId, sprintId, addedItemId, attachments);
        } else {
          Logger.log("⚠ Could not determine item ID for attachment upload");
        }

        
        // var zohoItemId = responseData.itemNo;
       // var ticketId = generateTicketId(zohoItemId);

      // ✅ Pass CC header so acknowledgment reaches all CC'd recipients

        var ccHeader = message.getHeader("Cc") || message.getHeader("CC") || "";
        sendAcknowledgemt(threads[i], from, subject, ticketId, ccHeader);
         
      //  sendAcknowledgemt(threads[i],from,subject,ticketId);

        threads[i].markRead();       // ✅ mark as read
        threads[i].addLabel(label);  // ✅ label so it's never processed again
        
        Logger.log("✅ Email has been Acknowledged"+subject+"|"+ticketId);
        Logger.log("✅ Done: " + subject);

      } else {
        Logger.log("❌ API Error for: " + subject + " | Code: " + responseCode);
      }

      Utilities.sleep(1000); // avoid rate limiting

    } catch (e) {
      Logger.log("🔴 Exception on email " + (i + 1) + ": " + e.message);
    }
  }

  Logger.log("🏁 All unread emails processed.");
}
