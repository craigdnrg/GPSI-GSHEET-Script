 /**
 * =================================================================
 * Google PageSpeed Insights (GPSI) Robust Checker - V3.2 (with Setup)
 * =================================================================
 *
 * ENSURE TO REPLACE API KEY: NULLED   to your actual API Key from GSC.
 *
 * Description: This script is architected to reliably process a large
 * number of URLs without timing out. It processes one URL at a time,
 * saves its state, and uses triggers to resume its work until the
 * queue is empty.
 *
 * v3.4 [TO BE ADDED] - Update to add in -> Need to seperate out the API key so people add their own, not too worried about limits tbf. 
 * 
 * v3.3 Update: 
 * Ensure that when the sheet is ran the cache is increased, currently set to 24 hours, although could be even longer. 
 * 
 * 
 * v3.2 Update:
 * - When starting, the script now checks the "Status" column.
 * - Any row already marked "Complete" will be skipped and not
 * added to the processing queue.
 * - Added a safety check in the processor to skip any "Complete"
 * rows it encounters.
 *
 * V3.1 Update:
 * - Added a 'Setup Sheet' menu item to automatically create and
 * format the 'PageSpeed' tab with all required headers.
 *
 * V3 Features:
 * - Resilient, trigger-based processing to avoid 6-min timeout.
 * - Start/Stop controls for the user.
 * - Clear "Queued" status for URLs awaiting processing.
 * - Efficient batch writing for each row's data.
 * 
 * V2.9 Features:
 * - Sets up the sheet for the person with the menu at top. 
 * 
 */

// --- CONFIGURATION ---
const CONFIG = {
  API_KEY: 'NULLED', // <--- IMPORTANT: PASTE YOUR API KEY HERE
  SHEET_NAME: 'PageSpeed',
  CACHE_EXPIRATION_SECONDS:   86400, //21600, // 6 hours
  TRIGGER_DELAY_MINUTES: 1,       // Delay between processing each URL
  
  // Column Mapping (Do not change unless you modify sheet structure)
  URL_COL: 1, STATUS_COL: 2, LAST_CHECKED_COL: 3, M_PERF_COL: 4, M_FCP_COL: 5,
  M_SI_COL: 6, M_LCP_COL: 7, M_TBT_COL: 8, M_CLS_COL: 9, D_PERF_COL: 10,
  D_FCP_COL: 11, D_SI_COL: 12, D_LCP_COL: 13, D_TBT_COL: 14, D_CLS_COL: 15,
  M_ISSUES_COL: 16, D_ISSUES_COL: 17, REPORT_LINK_COL: 18,
};
const SCRIPT_PROPERTIES = PropertiesService.getScriptProperties();
const TRIGGER_FUNCTION_NAME = 'processQueue';
// --- END CONFIGURATION ---


/**
 * Creates the custom menu when the spreadsheet is opened.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('PageSpeed Tools')
    .addItem('⚙️ Setup \'PageSpeed\' Sheet', 'setupSheet')
    .addSeparator()
    .addItem('▶️ Start All Checks', 'startProcessing')
    .addItem('⏹️ Stop/Reset Current Process', 'stopAndResetProcess')
    .addSeparator()
    .addItem('🗑️ Clear All Results & Cache', 'clearCacheAndResults')
    .addToUi();
}

/**
 * =================================================================
 * setupSheet
 * =================================================================
 * Creates the 'PageSpeed' sheet with all the required headers
 * and formatting.
 */
function setupSheet() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (sheet) {
    const response = ui.alert(
      `Sheet "${CONFIG.SHEET_NAME}" already exists.`,
      'Do you want to overwrite the first row with the correct headers?',
      ui.ButtonSet.YES_NO
    );
    if (response !== ui.Button.YES) {
      sheet.activate(); // Just go to the sheet
      return;
    }
  } else {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
  }

  // Clear existing content and formatting in the first row
  sheet.getRange(1, 1, 1, sheet.getMaxColumns()).clear().setFontWeight(null).setWrap(false);

  // Define headers based on CONFIG
  const headers = [];
  headers[CONFIG.URL_COL - 1] = "URL";
  headers[CONFIG.STATUS_COL - 1] = "Status";
  headers[CONFIG.LAST_CHECKED_COL - 1] = "Last Checked";
  headers[CONFIG.M_PERF_COL - 1] = "Mobile Perf";
  headers[CONFIG.M_FCP_COL - 1] = "Mobile FCP";
  headers[CONFIG.M_SI_COL - 1] = "Mobile SI";
  headers[CONFIG.M_LCP_COL - 1] = "Mobile LCP";
  headers[CONFIG.M_TBT_COL - 1] = "Mobile TBT";
  headers[CONFIG.M_CLS_COL - 1] = "Mobile CLS";
  headers[CONFIG.D_PERF_COL - 1] = "Desktop Perf";
  headers[CONFIG.D_FCP_COL - 1] = "Desktop FCP";
  headers[CONFIG.D_SI_COL - 1] = "Desktop SI";
  headers[CONFIG.D_LCP_COL - 1] = "Desktop LCP";
  headers[CONFIG.D_TBT_COL - 1] = "Desktop TBT";
  headers[CONFIG.D_CLS_COL - 1] = "Desktop CLS";
  headers[CONFIG.M_ISSUES_COL - 1] = "Mobile Top Issues";
  headers[CONFIG.D_ISSUES_COL - 1] = "Desktop Top Issues";
  headers[CONFIG.REPORT_LINK_COL - 1] = "Full Report Link";

  const headerRow = [headers];
  
  // Write headers
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues(headerRow);
  
  // Apply formatting
  headerRange.setFontWeight('bold');
  sheet.setFrozenRows(1);
  
  // Set column widths and wrapping for readability
  sheet.setColumnWidth(CONFIG.URL_COL, 300);
  sheet.setColumnWidth(CONFIG.STATUS_COL, 100);
  sheet.setColumnWidth(CONFIG.LAST_CHECKED_COL, 150);
  sheet.setColumnWidth(CONFIG.M_ISSUES_COL, 350);
  sheet.setColumnWidth(CONFIG.D_ISSUES_COL, 350);
  sheet.setColumnWidth(CONFIG.REPORT_LINK_COL, 150);

  // Apply text wrapping to issues columns (for all rows)
  sheet.getRange(1, CONFIG.M_ISSUES_COL, sheet.getMaxRows(), 1).setWrap(true);
  sheet.getRange(1, CONFIG.D_ISSUES_COL, sheet.getMaxRows(), 1).setWrap(true);
  
  // Resize other columns
  sheet.autoResizeColumns(CONFIG.M_PERF_COL, CONFIG.D_CLS_COL - CONFIG.M_PERF_COL + 1);

  sheet.activate();
  ui.alert(`Sheet "${CONFIG.SHEET_NAME}" has been set up successfully.`);
}

/**
 * Initiates the process. Cleans up old state/triggers and queues all URLs.
 */
function startProcessing() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    ui.alert(`Sheet "${CONFIG.SHEET_NAME}" not found.`, `Please run "PageSpeed Tools > ⚙️ Setup 'PageSpeed' Sheet" first.`, ui.ButtonSet.OK);
    return;
  }

  const response = ui.alert('Start Processing?', 'This will queue all non-completed URLs for checking. The script will process one URL every minute. You can close the sheet, and it will continue in the background.', ui.ButtonSet.OK_CANCEL);
  if (response !== ui.Button.OK) {
    return;
  }

  stopAndResetProcess(false); // Clean up previous runs without showing an alert
  
  // v3.2 Update: Get both URL and Status columns to check before queueing
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('No URLs found in the sheet.');
    return;
  }
  
  const dataRange = sheet.getRange(2, CONFIG.URL_COL, lastRow - 1, CONFIG.STATUS_COL);
  const data = dataRange.getValues();
  
  const urlQueue = [];
  for (let i = 0; i < data.length; i++) {
    const url = data[i][CONFIG.URL_COL - 1]; // URL is in the first col of our range
    const status = data[i][CONFIG.STATUS_COL - 1]; // Status is in the second col
    
    // Check for valid URL
    if (url && (url.startsWith('http://') || url.startsWith('https://'))) {
      // v3.2 Update: Only queue if status is not 'Complete'
      if (status !== 'Complete') {
        urlQueue.push(i + 2); // Push the row number
      }
    }
  }

  if (urlQueue.length === 0) {
    let message = 'No valid URLs found to process.';
    const validUrlsExist = data.some(row => row[CONFIG.URL_COL - 1] && (row[CONFIG.URL_COL - 1].startsWith('http://') || row[CONFIG.URL_COL - 1].startsWith('https://')));
    if (validUrlsExist) {
      message = 'No new URLs to process. All valid URLs are already marked as "Complete".';
    }
    ui.alert(message);
    return;
  }
  
  // Set the "Queued" status for all URLs to be processed (in a batch)
  const statusColumnValues = sheet.getRange(2, CONFIG.STATUS_COL, lastRow - 1, 1).getValues();
  let changed = false;
  urlQueue.forEach(rowNum => {
    const index = rowNum - 2; // Convert row number back to 0-based index
    if (statusColumnValues[index][0] !== 'Queued') {
      statusColumnValues[index][0] = 'Queued';
      changed = true;
    }
  });

  if (changed) {
    sheet.getRange(2, CONFIG.STATUS_COL, statusColumnValues.length, 1).setValues(statusColumnValues);
  }
  // End v3.2 batch update

  SCRIPT_PROPERTIES.setProperty('urlQueue', JSON.stringify(urlQueue));
  
  // Create a trigger to start the process almost immediately.
  ScriptApp.newTrigger(TRIGGER_FUNCTION_NAME)
    .timeBased()
    .after(1000) // 1 second
    .create();

  ui.alert(`Queued ${urlQueue.length} URLs for processing. The first check will begin shortly.`);
}

/**
 * This is the core processing function, run by a trigger.
 * It processes one URL from the queue and then sets a new trigger for the next one.
 */
function processQueue() {
  const queueProperty = SCRIPT_PROPERTIES.getProperty('urlQueue');
  if (!queueProperty) {
    cleanup();
    return;
  }
  
  const urlQueue = JSON.parse(queueProperty);
  if (urlQueue.length === 0) {
    cleanup();
    Logger.log('Processing complete. No more URLs in queue.');
    return;
  }

  const currentRow = urlQueue.shift(); // Get the first row from the queue
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    Logger.log(`Sheet "${CONFIG.SHEET_NAME}" not found. Stopping process.`);
    cleanup();
    return;
  }

  // v3.2 Update: Check if the row is already completed before processing
  const currentStatus = sheet.getRange(currentRow, CONFIG.STATUS_COL).getValue();
  const urlForLog = sheet.getRange(currentRow, CONFIG.URL_COL).getValue();

  if (currentStatus === 'Complete') {
    Logger.log(`Skipping row ${currentRow} (${urlForLog}). Status is already 'Complete'.`);
    
    // Save the modified queue
    SCRIPT_PROPERTIES.setProperty('urlQueue', JSON.stringify(urlQueue));
    
    // If there are more URLs, call the function again to process the next item immediately
    if (urlQueue.length > 0) {
      processQueue(); 
    } else {
      // This was the last item, so clean up.
      Logger.log('Finished processing the last URL (which was a skip).');
      cleanup();
    }
    return; // Stop execution for *this* call.
  }
  // End v3.2 Update

  // If not skipping, save the queue and proceed with processing.
  SCRIPT_PROPERTIES.setProperty('urlQueue', JSON.stringify(urlQueue));
  const url = urlForLog; // We already fetched it
  
  try {
    sheet.getRange(currentRow, CONFIG.STATUS_COL).setValue('Processing...');
    SpreadsheetApp.flush();

    const mobileResult = checkSingleUrl(url, 'MOBILE');
    const desktopResult = checkSingleUrl(url, 'DESKTOP');

    let rowData = [];
    if (mobileResult.error || desktopResult.error) {
      const mError = mobileResult.error || 'OK';
      const dError = desktopResult.error || 'OK';
      rowData = ['Error', new Date().toUTCString(),
        'N/A', 'N/A', 'N/A', 'N/A', 'N/A', 'N/A', // Mobile data
        'N/A', 'N/A', 'N/A', 'N/A', 'N/A', 'N/A', // Desktop data
        `Mobile: ${mError}`, `Desktop: ${dError}`, 'N/A'
      ];
    } else {
      const reportLink = `https://pagespeed.web.dev/report?url=${encodeURIComponent(url)}`;
      rowData = ['Complete', new Date().toUTCString(),
        mobileResult.performance, mobileResult.fcp, mobileResult.si, mobileResult.lcp, mobileResult.tbt, mobileResult.cls,
        desktopResult.performance, desktopResult.fcp, desktopResult.si, desktopResult.lcp, desktopResult.tbt, desktopResult.cls,
        mobileResult.topIssues, desktopResult.topIssues, reportLink
      ];
    }
    // Write the entire row's data in one efficient call
    sheet.getRange(currentRow, CONFIG.STATUS_COL, 1, rowData.length).setValues([rowData]);

  } catch (e) {
    sheet.getRange(currentRow, CONFIG.STATUS_COL, 1, 2).setValues([['Script Error', e.message]]);
  }
  
  // If there are more URLs, set a trigger for the next run
  if (urlQueue.length > 0) {
    ScriptApp.newTrigger(TRIGGER_FUNCTION_NAME)
      .timeBased()
      .after(CONFIG.TRIGGER_DELAY_MINUTES * 60 * 1000)
      .create();
  } else {
      Logger.log('Finished processing the last URL.');
      cleanup();
  }
}

/**
 * Deletes triggers and clears script properties. Called at the end of the process.
 */
function cleanup() {
  SCRIPT_PROPERTIES.deleteProperty('urlQueue');
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === TRIGGER_FUNCTION_NAME) {
      ScriptApp.deleteTrigger(trigger);
    }
  }
  Logger.log('Cleanup complete. All triggers and properties have been removed.');
}


/**
 * Menu function to manually stop the process.
 * @param {boolean} [showAlert=true] - Whether to show a confirmation alert.
 */
function stopAndResetProcess(showAlert = true) {
  cleanup(); // Deletes triggers and properties
  
  // Set any "Queued" or "Processing" statuses to "Stopped"
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) return; // Sheet doesn't exist, nothing to do

  if (sheet.getLastRow() > 1) {
    const statusRange = sheet.getRange(2, CONFIG.STATUS_COL, sheet.getLastRow() - 1, 1);
    const statuses = statusRange.getValues();
    let changed = false;
    for(let i=0; i<statuses.length; i++){
      if(statuses[i][0] === 'Queued' || statuses[i][0] === 'Processing...'){
        statuses[i][0] = 'Stopped';
        changed = true;
      }
    }
    if (changed) {
      statusRange.setValues(statuses);
    }
  }
  
  if (showAlert) {
    SpreadsheetApp.getUi().alert('Process stopped. All pending checks have been cancelled.');
  }
}


/**
 * Clears the script cache and all results from the sheet.
 */
function clearCacheAndResults() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('Are you sure?', 'This will STOP any running process, clear all cached results, and delete all data from column B onwards.', ui.ButtonSet.YES_NO);

  if (response == ui.Button.YES) {
    stopAndResetProcess(false); // Stop any running process first
    CacheService.getScriptCache().removeAll([]);
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) {
       ui.alert('Cache cleared. Sheet not found, so no results were cleared.');
       return;
    }

    if (sheet.getLastRow() > 1) {
      const rangeToClear = sheet.getRange(2, CONFIG.STATUS_COL, sheet.getLastRow() - 1, CONFIG.REPORT_LINK_COL - CONFIG.STATUS_COL + 1);
      rangeToClear.clearContent();
    }
    ui.alert('Cache and results have been cleared.');
  }
}

// --- CORE API AND PARSING FUNCTIONS ---

function checkSingleUrl(url, strategy) {
  const cache = CacheService.getScriptCache();
  const cacheKey = `psi_v3_${strategy}_${url}`;
  const cached = cache.get(cacheKey);
  if (cached) {
    Logger.log(`Cache hit for: ${url} (${strategy})`);
    return JSON.parse(cached);
  }
  const apiResponse = callPageSpeedApi(url, strategy);
  if (apiResponse.error) return apiResponse;
  const parsedData = parseApiResponse(apiResponse);
  if (!parsedData.error) {
    cache.put(cacheKey, JSON.stringify(parsedData), CONFIG.CACHE_EXPIRATION_SECONDS);
  }
  return parsedData;
}

function callPageSpeedApi(url, strategy) {
  const apiUrl = `https://www.googleapis.com/pagespeedonline/v5/runPagespeed?url=${encodeURIComponent(url)}&key=${CONFIG.API_KEY}&strategy=${strategy}&category=PERFORMANCE`;
  try {
    const response = UrlFetchApp.fetch(apiUrl, { 'muteHttpExceptions': true });
    if (response.getResponseCode() === 200) {
      return JSON.parse(response.getContentText());
    } else {
      const errorData = JSON.parse(response.getContentText());
      const errorMessage = (errorData.error && errorData.error.message) ? errorData.error.message : `API returned status ${response.getResponseCode()}`;
      Logger.log(`API Error for ${url} (${strategy}): ${errorMessage}`);
      return { error: errorMessage };
    }
  } catch (e) {
    Logger.log(`Fetch Error for ${url} (${strategy}): ${e.message}`);
    return { error: `Failed to fetch. Check URL or network. Details: ${e.message}` };
  }
}

function parseApiResponse(apiResponse) {
  try {
    if (!apiResponse.lighthouseResult) {
      Logger.log(`No lighthouseResult in API response: ${JSON.stringify(apiResponse)}`);
      return { error: 'Invalid API response: No lighthouseResult.' };
    }
    const lighthouse = apiResponse.lighthouseResult;
    const audits = lighthouse.audits;
    if (!audits) {
       return { error: 'Invalid API response: No audits.' };
    }
    const getAuditValue = (id) => audits[id] ? audits[id].displayValue : 'N/A';
    const result = {
      performance: (lighthouse.categories && lighthouse.categories.performance) ? Math.round(lighthouse.categories.performance.score * 100) : 'N/A',
      fcp: getAuditValue('first-contentful-paint'),
      si: getAuditValue('speed-index'),
      lcp: getAuditValue('largest-contentful-paint'),
      tbt: getAuditValue('total-blocking-time'),
      cls: getAuditValue('cumulative-layout-shift'),
      topIssues: ''
    };
    const opportunities = Object.values(audits)
      .filter(a => a.details && a.details.type === 'opportunity' && a.details.overallSavingsMs > 0)
      .sort((a, b) => b.details.overallSavingsMs - a.details.overallSavingsMs).slice(0, 5);
    
    result.topIssues = opportunities
      .map((audit, i) => {
        const savingsMs = (audit.details && audit.details.overallSavingsMs) ? Math.round(audit.details.overallSavingsMs) : 0;
        return `${i + 1}. ${audit.title} (Est. Savings: ${savingsMs} ms)`;
      })
      .join('\n');

    if (!result.topIssues) result.topIssues = "No major opportunities found.";
    return result;
  } catch (e) {
    Logger.log(`Parse Error: ${e.message}. Response was: ${JSON.stringify(apiResponse).substring(0, 500)}`);
    return { error: `Could not parse API response. ${e.message}` };
  }
}
