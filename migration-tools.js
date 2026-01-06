function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Migration Tools')
    .addItem('1. Build/Reset Dashboard', 'setupDashboard')
    .addSeparator()
    .addItem('2. Validate New URLs (Check Col C)', 'validateNewUrls')
    .addItem('3. Test Redirects (Check Col A)', 'auditRedirects')
    .addToUi();
}

/**
 * Sets up the spreadsheet structure and applies Conditional Formatting.
 * WARNING: CLEARS ALL DATA in 'Config' and 'Migration Audit'.
 */
function setupDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var response = ui.alert('Reset Dashboard?', 'This will clear all current data. Are you sure?', ui.ButtonSet.YES_NO);
  if (response == ui.Button.NO) return;
  
  // --- 1. Setup Config Tab ---
  var configSheet = ss.getSheetByName("Config");
  if (!configSheet) { configSheet = ss.insertSheet("Config"); }
  configSheet.clear(); 
  
  configSheet.getRange("A1").setValue("Migration Configuration").setFontSize(14).setFontWeight("bold");
  configSheet.getRange("A3").setValue("New Base URL (including https://):").setFontWeight("bold");
  configSheet.getRange("B3").setValue("https://example.com").setBackground("#FFF2CC");
  configSheet.getRange("C3").setValue("Ensure no trailing slash");
  
  configSheet.getRange("A5").setValue("Instructions:").setFontWeight("bold");
  configSheet.getRange("A6").setValue("1. Enter your new site's base URL in cell B3 above.");
  configSheet.getRange("A7").setValue("2. Go to the 'Migration Audit' tab.");
  configSheet.getRange("A8").setValue("3. Paste your OLD site URLs into Column A (starting at A2).");
  configSheet.getRange("A9").setValue("4. Columns B and C will calculate automatically.");
  
  configSheet.setColumnWidth(1, 400); 
  configSheet.setColumnWidth(2, 300);

  // --- 2. Setup Migration Audit Tab ---
  var auditSheet = ss.getSheetByName("Migration Audit");
  if (!auditSheet) { auditSheet = ss.insertSheet("Migration Audit"); }
  
  auditSheet.clear(); 
  auditSheet.clearConditionalFormatRules(); // Clear old formatting rules to avoid duplicates

  var headers = [["Original URL (Old)", "Extracted Slug", "Expected New URL", "HTTP Status", "Actual Destination / Final URL", "Review"]];
  auditSheet.getRange("A1:F1").setValues(headers).setFontWeight("bold").setBackground("#EFEFEF").setBorder(true, true, true, true, true, true);
  auditSheet.setFrozenRows(1);
  
  // Formulas
  auditSheet.getRange("B2").setFormula('=MAP(A2:A, LAMBDA(url, IF(url="", "", REGEXEXTRACT(url, "[^/]+/?$"))))');
  auditSheet.getRange("C2").setFormula('=MAP(B2:B, LAMBDA(slug, IF(slug="", "", IF(RIGHT(Config!$B$3,1)="/", LEFT(Config!$B$3, LEN(Config!$B$3)-1), Config!$B$3) & "/" & slug)))');
  var matchFormula = '=MAP(C2:C, E2:E, LAMBDA(exp, act, IF(exp="", "", IF(act="", "Pending...", IF(REGEXREPLACE(exp, "/$", "") = REGEXREPLACE(act, "/$", ""), "✅ Match", "⚠️ Mismatch")))))';
  auditSheet.getRange("F2").setFormula(matchFormula);

  // --- 3. Add Conditional Formatting ---
  var range = auditSheet.getRange("F2:F1000");
  
    // Rule 2: Red/Amber for Mismatch
  var ruleMismatch = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains("Mis")
    .setBackground("#f4c7c3") // Light Red
    .setFontColor("#c90000")
    .setRanges([range])
    .build();
    
  // Rule 1: Green for Match
  var ruleMatch = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains("Ma")
    .setBackground("#b7e1cd") // Light Green
    .setFontColor("#0b5394")
    .setRanges([range])
    .build();



  var rules = auditSheet.getConditionalFormatRules();
  rules.push(ruleMatch);
  rules.push(ruleMismatch);
  auditSheet.setConditionalFormatRules(rules);

  ui.alert('Dashboard Reset Complete!');
}

function validateNewUrls() { processUrls(3); }
function auditRedirects() { processUrls(1); }

function processUrls(colIndex) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Migration Audit");
  if (!sheet) return;

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { SpreadsheetApp.getUi().alert("No data found."); return; }
  
  sheet.getRange(2, 4, lastRow - 1, 2).clearContent();

  var urls = sheet.getRange(2, colIndex, lastRow - 1, 1).getValues();
  var configHost = ss.getSheetByName("Config").getRange("B3").getValue();
  if (configHost.slice(-1) === "/") configHost = configHost.slice(0, -1);

  var outputValues = [];

  for (var i = 0; i < urls.length; i++) {
    var url = urls[i][0];
    if (url == "") { outputValues.push(["", ""]); continue; }

    try {
      var options = { 'muteHttpExceptions': true, 'followRedirects': false };
      var response = UrlFetchApp.fetch(url, options);
      var code = response.getResponseCode();
      var headers = response.getAllHeaders();
      var location = "";

      if (code == 200) { location = url; } 
      else if (code >= 300 && code < 400) {
        var rawLoc = headers['Location'] || headers['location'] || "";
        if (rawLoc.startsWith("/")) { location = configHost + rawLoc; } 
        else if (rawLoc.startsWith("http")) { location = rawLoc; } 
        else { location = rawLoc; }
      } 
      else { location = "Error Page"; }

      outputValues.push([code, location]);
    } catch (e) { outputValues.push(["Error", e.message]); }
  }
  sheet.getRange(2, 4, lastRow - 1, 2).setValues(outputValues);
  SpreadsheetApp.getUi().alert("Check Complete!");
}
