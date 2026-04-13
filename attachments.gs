function getFilteredAttachments(message) {
  var attachments = message.getAttachments();
  var filtered = [];

  for (var i = 0; i < attachments.length; i++) {
    var att = attachments[i];
    var name = att.getName().toLowerCase();
    var size = att.getSize();
    var contentType = att.getContentType().toLowerCase();

    // ── Skip tiny images (logos/signatures are usually < 5KB) ──
    if (size < 5 * 1024 && contentType.indexOf("image/") === 0) {
      Logger.log("⏭ Skipped (likely logo/signature): " + att.getName() + " | " + size + " bytes");
      continue;
    }

    // ── Skip known signature image formats by name patterns ──
    var signaturePatterns = ["logo", "signature", "sig", "banner", "footer", "badge", "icon", "avatar"];
    var isSignatureImage = signaturePatterns.some(function(p) { return name.indexOf(p) !== -1; });
    if (isSignatureImage && contentType.indexOf("image/") === 0) {
      Logger.log("⏭ Skipped (name matched signature pattern): " + att.getName());
      continue;
    }

    // ── Skip inline images (embedded in email body, not real attachments) ──
    if (att.isGoogleType && att.isGoogleType()) continue; // skip Google-native types

    filtered.push(att);
    Logger.log("📎 Queued attachment: " + att.getName() + " | " + size + " bytes | " + contentType);
  }

  Logger.log("📎 Attachments to upload: " + filtered.length + " of " + attachments.length);
  return filtered;
}


function uploadAttachmentsToWorkItem(accessToken, teamId, projectId, sprintId, itemId, attachments) {
  if (!attachments || attachments.length === 0) {
    Logger.log("📎 No attachments to upload");
    return;
  }

  for (var i = 0; i < attachments.length; i++) {
    try {
      var att     = attachments[i];
      var blobs = att.copyBlob();
      var url     = "https://sprintsapi.zoho.in/zsapi/team/" + teamId
                  + "/projects/" + projectId
                  + "/sprints/" + sprintId
                  + "/item/" + itemId + "/attachments/";

      var options = {
        method            : "post",
        headers           : { Authorization: "Zoho-oauthtoken " + accessToken },
        payload           : { action: "attachment",
                              uploadfile: blobs},   // GmailAttachment is directly accepted by UrlFetchApp
        muteHttpExceptions: true
      };

      var response = UrlFetchApp.fetch(url, options);
      var code     = response.getResponseCode();

      if (code === 200 || code === 201) {
        Logger.log("✅ Uploaded attachment: " + att.getName());
      } else {
        Logger.log("❌ Failed to upload: " + att.getName() + " | Code: " + code + " | " + response.getContentText());
      }

      Utilities.sleep(500); // small delay between uploads

    } catch (e) {
      Logger.log("🔴 Exception uploading attachment " + attachments[i].getName() + ": " + e.message);
    }
  }
}
