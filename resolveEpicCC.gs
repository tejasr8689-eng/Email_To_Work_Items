function resolveEpicCc(message) {

  var CC_TO_EPIC_MAP = {
  "degreepuc1234@gmail.com" : "48644000000329039",   // EPIC A
  "tejasr8689@gmail.com"    : "48644000000294099",   // EPIC B
  "tejasr23mca@rnsit.ac.in" : "48644000000329041",   // EPIC C
  };

  var ccHeader = message.getHeader("Cc") || message.getHeader("CC") || "";
  //var cc   = message.getHeader("Cc") || message.getHeader("CC") || "";
  if (!ccHeader) return null;

  var ccEmails = ccHeader.toLowerCase().match(/[\w.-]+@[\w.-]+\.\w+/g) || [];
  for (var email in CC_TO_EPIC_MAP) {
    if (ccEmails.indexOf(email.toLowerCase()) !== -1) {
      Logger.log("🎯 Matched CC: " + email + " → Epic: " + CC_TO_EPIC_MAP[email]);
      return CC_TO_EPIC_MAP[email];
    }
  }
 return null; // no match
}
