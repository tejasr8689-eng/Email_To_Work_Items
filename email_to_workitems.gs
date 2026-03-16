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

  var refreshToken  = "1000.e8e73c814aed4c76bceacb718c5619d4.d3668a0de7378f8d1c995861947a60cd";
  var clientId      = "1000.NJ2TE4B9P34HUITK0TREF2DGEKTE3X";
  var clientSeceret = "0290f3beea020388220d28dfbd5a7cde64dd2fc5cf";

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


// ===============================================
// WORK ITEM CREATOR
// ===============================================

function createWorkItemsFromUnreadEmails() {

  // --- Config ---
  var accessToken    = acessTokenGenerator();
  var teamId         = "60066513512";
  var projectId      = "52688000000008845";
  var sprintId       = "52688000000008851";
  var projItemTypeId = "52688000000008895";
  var projPriorityId = "52688000000008889";

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

      var subject = message.getSubject();
      var body    = message.getPlainBody();
      var from    = message.getFrom();

      Logger.log("📨 Processing (" + (i + 1) + "/" + threads.length + "): " + subject);

      // --- Build API URL ---
      var url = "https://sprintsapi.zoho.in/zsapi/team/"
        + teamId      + "/projects/"
        + projectId   + "/sprints/"
        + sprintId    + "/item/"
        + "?projitemtypeid=" + projItemTypeId
        + "&projpriorityid=" + projPriorityId;

      // --- Build Payload ---
      var payload = {
        name        : subject,
        description : "From: " + from + "\n\n" + body
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
        threads[i].markRead();       // ✅ mark as read
        threads[i].addLabel(label);  // ✅ label so it's never processed again
        Logger.log("✅ Done: " + subject);
      } else {
        Logger.log("❌ API Error for: " + subject + " | Code: " + responseCode);
      }

      Utilities.sleep(300); // avoid rate limiting

    } catch (e) {
      Logger.log("🔴 Exception on email " + (i + 1) + ": " + e.message);
    }
  }

  Logger.log("🏁 All unread emails processed.");
}
