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
 * Sets up the spreadsheet structure and applies Robust Conditional Formatting.
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
  auditSheet.clearConditionalFormatRules(); 

  var headers = [["Original URL (Old)", "Extracted Slug", "Expected New URL", "HTTP Status", "Actual Destination / Final URL", "Review"]];
  auditSheet.getRange("A1:F1").setValues(headers).setFontWeight("bold").setBackground("#EFEFEF").setBorder(true, true, true, true, true, true);
  auditSheet.setFrozenRows(1);
  
  // Formulas
  auditSheet.getRange("B2").setFormula('=MAP(A2:A, LAMBDA(url, IF(url="", "", REGEXEXTRACT(url, "[^/]+/?$"))))');
  auditSheet.getRange("C2").setFormula('=MAP(B2:B, LAMBDA(slug, IF(slug="", "", IF(RIGHT(Config!$B$3,1)="/", LEFT(Config!$B$3, LEN(Config!$B$3)-1), Config!$B$3) & "/" & slug)))');
  
  // Updated Formula: Uses "Mismatched" as requested
  var matchFormula = '=MAP(C2:C, E2:E, LAMBDA(exp, act, IF(exp="", "", IF(act="", "Pending...", IF(REGEXREPLACE(exp, "/$", "") = REGEXREPLACE(act, "/$", ""), "✅ Match", "⚠️ Mismatched")))))';
  auditSheet.getRange("F2").setFormula(matchFormula);

  // --- 3. Add Safer Conditional Formatting ---
  var range = auditSheet.getRange("F2:F1000");
  
  // Rule 1: Red/Amber for Mismatched
  // We use "whenTextEqualTo" to avoid "Match" matching inside "Mismatched"
  var ruleMismatch = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("⚠️ Mismatched")
    .setBackground("#f4c7c3") // Light Red
    .setFontColor("#c90000")
    .setRanges([range])
    .build();

  // Rule 2: Green for Match
  var ruleMatch = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("✅ Match")
    .setBackground("#b7e1cd") // Light Green
    .setFontColor("#0b5394")
    .setRanges([range])
    .build();

  var rules = auditSheet.getConditionalFormatRules();
  rules.push(ruleMismatch); // Push Red first (priority)
  rules.push(ruleMatch);
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

/**
 * Checks the HTTP status code of a URL
 * @param {string} url The URL to check
 * @return {string} The HTTP status code or error message
 * @customfunction
 */
function getStatusCode(url) {
  try {
    // Ensure the URL has a protocol
    if (!url.startsWith('http://') && !url.startsWith('https://')) {
      url = 'https://' + url;
    }
    // Send HTTP request to the URL
    var response = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
    
    // Get the status code of the response
    var statusCode = response.getResponseCode();
    
    // Custom error messages based on status code
    if (statusCode == 200) {
      return "Success (200)";
    } else if (statusCode == 404) {
      return "Not Found (404)";
    } else if (statusCode == 500) {
      return "Internal Server Error (500)";
    } else if (statusCode == 503) {
      return "Service Unavailable (503)";
    } else if (statusCode == 401) {
      return "Unauthorized (401)";
    } else if (statusCode == 403) {
      return "Forbidden (403)";
    } else if (statusCode == 502) {
      return "Bad Gateway (502)";
    } else if (statusCode == 504) {
      return "Gateway Timeout (504)";
    } else {
      return "Status Code: " + statusCode;
    }
  } catch (e) {
    // Catch any other errors (e.g., network errors)
    return "Error: " + e.message;
  }
}

/**
 * Checks the HTTP status code of URL/robots.txt
 * @param {string} url The base URL to check
 * @return {string} The HTTP status code or error message
 * @customfunction
 */
function getRobotsStatus(url) {
  try {
    // Ensure the URL has a protocol
    if (!url.startsWith('http://') && !url.startsWith('https://')) {
      url = 'https://' + url;
    }
    // Ensure URL ends with a slash
    if (!url.endsWith('/')) {
      url = url + '/';
    }
    var robotsUrl = url + 'robots.txt';
    // Send HTTP request to robots.txt
    var response = UrlFetchApp.fetch(robotsUrl, {muteHttpExceptions: true});
    
    // Get the status code of the response
    var statusCode = response.getResponseCode();
    
    // Custom error messages based on status code
    if (statusCode == 200) {
      return "Success (200)";
    } else if (statusCode == 404) {
      return "Not Found (404)";
    } else if (statusCode == 500) {
      return "Internal Server Error (500)";
    } else if (statusCode == 503) {
      return "Service Unavailable (503)";
    } else if (statusCode == 401) {
      return "Unauthorized (401)";
    } else if (statusCode == 403) {
      return "Forbidden (403)";
    } else if (statusCode == 502) {
      return "Bad Gateway (502)";
    } else if (statusCode == 504) {
      return "Gateway Timeout (504)";
    } else {
      return "Status Code: " + statusCode;
    }
  } catch (e) {
    return "Error: " + e.message;
  }
}
/**
 * Checks the HTTP status code of URL/sitemap_index.xml or URL/sitemap.xml
 * @param {string} url The base URL to check
 * @return {string} The HTTP status code or error message
 * @customfunction
 */
function getSitemapStatus(url) {
  try {
    // Ensure the URL has a protocol
    if (!url.startsWith('http://') && !url.startsWith('https://')) {
      url = 'https://' + url;
    }
    // Ensure URL ends with a slash
    if (!url.endsWith('/')) {
      url = url + '/';
    }
    // Try sitemap_index.xml first
    var sitemapUrl = url + 'sitemap_index.xml';
    var response = UrlFetchApp.fetch(sitemapUrl, {muteHttpExceptions: true});
    var statusCode = response.getResponseCode();
    
    // If sitemap_index.xml is not found, try sitemap.xml
    if (statusCode != 200) {
      sitemapUrl = url + 'sitemap.xml';
      response = UrlFetchApp.fetch(sitemapUrl, {muteHttpExceptions: true});
      statusCode = response.getResponseCode();
    }
    
    // Custom error messages based on status code
    if (statusCode == 200) {
      return "Success (200)";
    } else if (statusCode == 404) {
      return "Not Found (404)";
    } else if (statusCode == 500) {
      return "Internal Server Error (500)";
    } else if (statusCode == 503) {
      return "Service Unavailable (503)";
    } else if (statusCode == 401) {
      return "Unauthorized (401)";
    } else if (statusCode == 403) {
      return "Forbidden (403)";
    } else if (statusCode == 502) {
      return "Bad Gateway (502)";
    } else if (statusCode == 504) {
      return "Gateway Timeout (504)";
    } else {
      return "Status Code: " + statusCode;
    }
  } catch (e) {
    return "Error: " + e.message;
  }
}

/**
 * Attempts to find the sitemap URL by checking robots.txt or defaults to sitemap_index.xml or sitemap.xml
 * @param {string} url The base URL to check
 * @return {string} The sitemap URL or 'Not Found'
 * @customfunction
 */
function getSitemapUrl(url) {
  try {
    // Ensure the URL has a protocol
    if (!url.startsWith('http://') && !url.startsWith('https://')) {
      url = 'https://' + url;
    }
    // Ensure URL ends with a slash
    if (!url.endsWith('/')) {
      url = url + '/';
    }
    // Check robots.txt for Sitemap directive
    var robotsUrl = url + 'robots.txt';
    var response = UrlFetchApp.fetch(robotsUrl, {muteHttpExceptions: true});
    
    if (response.getResponseCode() == 200) {
      var content = response.getContentText();
      // Look for Sitemap directive in robots.txt
      var lines = content.split('\n');
      for (var i = 0; i < lines.length; i++) {
        if (lines[i].toLowerCase().startsWith('sitemap:')) {
          return lines[i].replace(/^sitemap:\s*/i, '').trim();
        }
      }
    }
    // If no sitemap in robots.txt, try sitemap_index.xml
    var sitemapUrl = url + 'sitemap_index.xml';
    var sitemapResponse = UrlFetchApp.fetch(sitemapUrl, {muteHttpExceptions: true});
    if (sitemapResponse.getResponseCode() == 200) {
      return sitemapUrl;
    }
    // If sitemap_index.xml is not found, try sitemap.xml
    sitemapUrl = url + 'sitemap.xml';
    sitemapResponse = UrlFetchApp.fetch(sitemapUrl, {muteHttpExceptions: true});
    if (sitemapResponse.getResponseCode() == 200) {
      return sitemapUrl;
    }
    return "Not Found";
  } catch (e) {
    return "Error: " + e.message;
  }
}

/**
 * Checks if the webpage includes a Google Tag Manager (GTM) container
 * @param {string} url The URL to check
 * @return {string} "Found", "Not Found", or error message
 * @customfunction
 */
function getGTMStatus(url) {
  try {
    // Ensure the URL has a protocol
    if (!url.startsWith('http://') && !url.startsWith('https://')) {
      url = 'https://' + url;
    }
    // Send HTTP request to the URL
    var response = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
    var statusCode = response.getResponseCode();
    
    // Check if the page is accessible
    if (statusCode != 200) {
      if (statusCode == 404) return "Not Found (404)";
      if (statusCode == 500) return "Internal Server Error (500)";
      if (statusCode == 503) return "Service Unavailable (503)";
      if (statusCode == 401) return "Unauthorized (401)";
      if (statusCode == 403) return "Forbidden (403)";
      if (statusCode == 502) return "Bad Gateway (502)";
      if (statusCode == 504) return "Gateway Timeout (504)";
      return "Status Code: " + statusCode;
    }
    
    // Get page content
    var content = response.getContentText();
    // Check for GTM script or noscript tag (GTM-XXXXXX pattern)
    if (content.match(/GTM-[A-Za-z0-9]{6,}/) || content.includes('googletagmanager.com')) {
      return "Found";
    }
    return "Not Found";
  } catch (e) {
    return "Error: " + e.message;
  }
}

/**
 * Checks if the webpage includes a Meta Pixel
 * @param {string} url The URL to check
 * @return {string} "Found", "Not Found", or error message
 * @customfunction
 */
function getMetaPixelStatus(url) {
  try {
    // Ensure the URL has a protocol
    if (!url.startsWith('http://') && !url.startsWith('https://')) {
      url = 'https://' + url;
    }
    // Send HTTP request to the URL
    var response = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
    var statusCode = response.getResponseCode();
    
    // Check if the page is accessible
    if (statusCode != 200) {
      if (statusCode == 404) return "Not Found (404)";
      if (statusCode == 500) return "Internal Server Error (500)";
      if (statusCode == 503) return "Service Unavailable (503)";
      if (statusCode == 401) return "Unauthorized (401)";
      if (statusCode == 403) return "Forbidden (403)";
      if (statusCode == 502) return "Bad Gateway (502)";
      if (statusCode == 504) return "Gateway Timeout (504)";
      return "Status Code: " + statusCode;
    }
    
    // Get page content
    var content = response.getContentText();
    // Check for Meta Pixel script (fbq('init', 'PIXEL_ID') or connect.facebook.net)
    if (content.match(/fbq\(['"]init['"],\s*['"][0-9]+['"]\)/) || content.includes('connect.facebook.net')) {
      return "Found";
    }
    return "Not Found";
  } catch (e) {
    return "Error: " + e.message;
  }
}

/**
 * Checks if the webpage includes breadcrumbs (Schema.org BreadcrumbList)
 * @param {string} url The URL to check
 * @return {string} "Found", "Not Found", or error message
 * @customfunction
 */
function getBreadcrumbsStatus(url) {
  try {
    // Ensure the URL has a protocol
    if (!url.startsWith('http://') && !url.startsWith('https://')) {
      url = 'https://' + url;
    }
    // Send HTTP request to the URL
    var response = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
    var statusCode = response.getResponseCode();
    
    // Check if the page is accessible
    if (statusCode != 200) {
      if (statusCode == 404) return "Not Found (404)";
      if (statusCode == 500) return "Internal Server Error (500)";
      if (statusCode == 503) return "Service Unavailable (503)";
      if (statusCode == 401) return "Unauthorized (401)";
      if (statusCode == 403) return "Forbidden (403)";
      if (statusCode == 502) return "Bad Gateway (502)";
      if (statusCode == 504) return "Gateway Timeout (504)";
      return "Status Code: " + statusCode;
    }
    
    // Get page content
    var content = response.getContentText();
    // Check for Schema.org BreadcrumbList (JSON-LD or microdata)
    if (content.match(/"@type"\s*:\s*"BreadcrumbList"/) || content.match(/itemtype\s*=\s*['"]http:\/\/schema.org\/BreadcrumbList['"]/)) {
      return "Found";
    }
    return "Not Found";
  } catch (e) {
    return "Error: " + e.message;
  }
}
