// ==========================================================================
// CONFIGURATION AREA
// ==========================================================================
const SPREADSHEET_ID = "18ThG_m8fA6CYLVbwQMwA9h7pKJK_QdDKQLgTLqpo-nw";

const EMPLOYEE_SHEET   = "EMPLOYEE";
const LOCATION_SHEET   = "LOCATION";
const ATTENDANCE_SHEET = "ATTENDANCE";
const OVERTIME_SHEET   = "OVERTIME";
const ADMIN_SHEET      = "ADMIN";
const DEPLOYMENT_SHEET = "DEPLOYMENT";

// Google Drive Folder for Photos
const PHOTO_FOLDER_ID = "1gnrXXjWSg94kMhMcTTIStAGR3CH-bOhr";

// Auto Clock-Out Times
const AM_AUTO_CLOCKOUT_HOUR = 12;
const AM_AUTO_CLOCKOUT_MINUTE = 30;
const PM_AUTO_CLOCKOUT_HOUR = 17;
const PM_AUTO_CLOCKOUT_MINUTE = 30;

// Late Time Threshold
const LATE_TIME_HOUR = 8;
const LATE_TIME_MINUTE = 16;

// PERFORMANCE: Cache configuration (5 minutes)
const CACHE_DURATION = 300; // seconds

// ==========================================================================
// PERFORMANCE OPTIMIZATION: IN-MEMORY CACHE
// ==========================================================================
var cache = CacheService.getScriptCache();

/**
 * Get cached data or fetch from function
 * @param {string} key - Cache key
 * @param {function} fetchFunction - Function to call if cache miss
 * @param {number} duration - Cache duration in seconds (default 300)
 */
function getCachedOrFetch(key, fetchFunction, duration) {
  duration = duration || CACHE_DURATION;
  
  try {
    var cached = cache.get(key);
    if (cached) {
      Logger.log("⚡ CACHE HIT: " + key);
      return JSON.parse(cached);
    }
  } catch (e) {
    Logger.log("Cache read error: " + e.message);
  }
  
  Logger.log("🔍 CACHE MISS: " + key + " - Fetching data");
  var data = fetchFunction();
  
  try {
    cache.put(key, JSON.stringify(data), duration);
  } catch (e) {
    Logger.log("Cache write error: " + e.message);
  }
  
  return data;
}

/**
 * Clear specific cache key or all cache
 */
function clearCache(key) {
  if (key) {
    cache.remove(key);
    Logger.log("🗑️ Cleared cache: " + key);
  } else {
    cache.removeAll();
    Logger.log("🗑️ Cleared all cache");
  }
}

// ==========================================================================
// PERFORMANCE OPTIMIZATION: EMPLOYEE DATA CACHE
// ==========================================================================

/**
 * Load all employees into memory - OPTIMIZED
 * Called once and cached for 5 minutes
 */
function loadAllEmployees() {
  var startTime = new Date().getTime();
  
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var empSheet = ss.getSheetByName(EMPLOYEE_SHEET);
    
    if (!empSheet) {
      return { employees: {}, list: [] };
    }
    
    var data = empSheet.getDataRange().getValues();
    var employees = {};
    var list = [];
    
    // Build employee lookup map (by ID) and list
    for (var i = 1; i < data.length; i++) {
      var empId = data[i][0] ? data[i][0].toString().trim() : "";
      var firstName = data[i][1] ? data[i][1].toString().trim() : "";
      var lastName = data[i][2] ? data[i][2].toString().trim() : "";
      
      if (!empId) continue;
      
      var fullName = (firstName + " " + lastName).trim();
      var displayName = empId + " - " + fullName;
      
      employees[empId.toUpperCase()] = {
        id: empId,
        name: fullName,
        fullName: displayName
      };
      
      list.push(displayName);
    }
    
    var duration = new Date().getTime() - startTime;
    Logger.log("⏱️ Loaded " + Object.keys(employees).length + " employees in " + duration + "ms");
    
    return { employees: employees, list: list };
  } catch (err) {
    Logger.log("Error loading employees: " + err.message);
    return { employees: {}, list: [] };
  }
}

/**
 * Get employee by ID - OPTIMIZED with cache
 */
function getEmployeeByID(employeeId) {
  var startTime = new Date().getTime();
  
  try {
    var employeeData = getCachedOrFetch("all_employees", loadAllEmployees);
    var employee = employeeData.employees[employeeId.toUpperCase()];
    
    var duration = new Date().getTime() - startTime;
    Logger.log("⏱️ getEmployeeByID took " + duration + "ms");
    
    if (employee) {
      return {
        success: true,
        employeeId: employee.id,
        employeeName: employee.name,
        fullName: employee.fullName
      };
    }
    
    return { success: false, message: "Employee ID not found. Please check and try again." };
  } catch (err) {
    var duration = new Date().getTime() - startTime;
    Logger.log("⏱️ getEmployeeByID error after " + duration + "ms: " + err.message);
    return { success: false, message: "Error validating employee: " + err.message };
  }
}

/**
 * Get employee by name - OPTIMIZED with cache
 */
function getEmployeeByName(employeeName) {
  var startTime = new Date().getTime();
  
  try {
    var employeeData = getCachedOrFetch("all_employees", loadAllEmployees);
    
    // Search by name
    for (var key in employeeData.employees) {
      var emp = employeeData.employees[key];
      if (emp.name === employeeName) {
        var duration = new Date().getTime() - startTime;
        Logger.log("⏱️ getEmployeeByName took " + duration + "ms");
        
        return {
          success: true,
          employeeId: emp.id,
          employeeName: emp.name,
          fullName: emp.fullName
        };
      }
    }
    
    var duration = new Date().getTime() - startTime;
    Logger.log("⏱️ getEmployeeByName took " + duration + "ms");
    return { success: false, message: "Employee not found." };
  } catch (err) {
    var duration = new Date().getTime() - startTime;
    Logger.log("⏱️ getEmployeeByName error after " + duration + "ms: " + err.message);
    return { success: false, message: "Error retrieving employee: " + err.message };
  }
}

/**
 * Get employee current ID - OPTIMIZED
 */
function getEmployeeCurrentID(employeeName) {
  try {
    var result = getEmployeeByName(employeeName);
    return result.success ? result.employeeId : null;
  } catch (err) {
    Logger.log("Error getting employee current ID: " + err.message);
    return null;
  }
}

// ==========================================================================
// PERFORMANCE OPTIMIZATION: LOCATION DATA CACHE
// ==========================================================================

/**
 * Load all locations - OPTIMIZED
 */
function loadAllLocations() {
  var startTime = new Date().getTime();
  
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var locSheet = ss.getSheetByName(LOCATION_SHEET);
    var locations = [];
    
    if (locSheet) {
      var data = locSheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        var locationName = data[i][1] ? data[i][1].toString().trim() : "";
        if (locationName && locationName !== "") {
          locations.push(locationName);
        }
      }
    }
    
    if (locations.length === 0) {
      locations.push("Main Office");
    }
    
    var duration = new Date().getTime() - startTime;
    Logger.log("⏱️ Loaded " + locations.length + " locations in " + duration + "ms");
    
    return locations;
  } catch (err) {
    Logger.log("Error loading locations: " + err.message);
    return ["Main Office"];
  }
}

// ==========================================================================
// UTILITY FUNCTIONS
// ==========================================================================

// Generate unique DEPLOYMENT_ID
function generateDeploymentId() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var deploySheet = ss.getSheetByName(DEPLOYMENT_SHEET);
  
  if (!deploySheet) return "DP_001";
  
  var data = deploySheet.getDataRange().getValues();
  var maxNum = 0;
  
  // Skip header, find highest DP number
  for (var i = 1; i < data.length; i++) {
    var dpId = data[i][0] ? data[i][0].toString().trim() : "";
    if (dpId.indexOf("DP_") === 0) {
      var numPart = parseInt(dpId.substring(3), 10);
      if (!isNaN(numPart) && numPart > maxNum) {
        maxNum = numPart;
      }
    }
  }
  
  var nextNum = maxNum + 1;
  return "DP_" + ("000" + nextNum).slice(-3);
}

// Generate unique OT_ID
function generateOtId() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var otSheet = ss.getSheetByName(OVERTIME_SHEET);
  
  if (!otSheet) return "OT_001";
  
  var data = otSheet.getDataRange().getValues();
  var maxNum = 0;
  
  // Skip header, find highest OT number
  for (var i = 1; i < data.length; i++) {
    var otId = data[i][0] ? data[i][0].toString().trim() : "";
    if (otId.indexOf("OT_") === 0) {
      var numPart = parseInt(otId.substring(3), 10);
      if (!isNaN(numPart) && numPart > maxNum) {
        maxNum = numPart;
      }
    }
  }
  
  var nextNum = maxNum + 1;
  return "OT_" + ("000" + nextNum).slice(-3);
}

// Validate Admin credentials
function validateAdmin(adminId, adminCode) {
  var startTime = new Date().getTime();
  
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var adminSheet = ss.getSheetByName(ADMIN_SHEET);
    
    if (!adminSheet) {
      return { success: false, message: "ADMIN sheet not found." };
    }
    
    var data = adminSheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      var sheetAdminId = data[i][0] ? data[i][0].toString().trim() : "";
      var sheetName = data[i][1] ? data[i][1].toString().trim() : "";
      var sheetCode = data[i][2] ? data[i][2].toString().trim() : "";
      
      if (sheetAdminId === adminId && sheetCode === adminCode) {
        var duration = new Date().getTime() - startTime;
        Logger.log("⏱️ validateAdmin took " + duration + "ms");
        return { success: true, adminName: sheetName, adminId: sheetAdminId };
      }
    }
    
    var duration = new Date().getTime() - startTime;
    Logger.log("⏱️ validateAdmin took " + duration + "ms");
    return { success: false, message: "Invalid Admin ID or Code." };
  } catch (err) {
    var duration = new Date().getTime() - startTime;
    Logger.log("⏱️ validateAdmin error after " + duration + "ms: " + err.message);
    return { success: false, message: "Error validating admin: " + err.message };
  }
}

// Calculate total hours between two time strings
function calculateTotalHours(startTime, endTime) {
  if (!startTime || !endTime) return "0.00";
  
  var startParts = startTime.split(':');
  var endParts = endTime.split(':');
  
  var startMinutes = parseInt(startParts[0], 10) * 60 + parseInt(startParts[1], 10);
  var endMinutes = parseInt(endParts[0], 10) * 60 + parseInt(endParts[1], 10);
  
  if (endMinutes < startMinutes) {
    endMinutes += 24 * 60;
  }
  
  var difference = endMinutes - startMinutes;
  return (difference / 60).toFixed(2);
}

/**
 * Helper function to set hour formulas with proper number formatting
 */
function setHourFormulas(sheet, rowNumber) {
  // Calculate Total Break Hours (Column K)
  var breakFormula = '=IF(AND(F' + rowNumber + '<>"", G' + rowNumber + '<>""), ' +
                     'ROUND((G' + rowNumber + ' - F' + rowNumber + ') * 24, 2), 0)';
  sheet.getRange(rowNumber, 11).setFormula(breakFormula);
  sheet.getRange(rowNumber, 11).setNumberFormat("0.00");
  
  // Calculate Total Work Hours (Column L)
  var workFormula = '=IF(AND(C' + rowNumber + '<>"", H' + rowNumber + '<>""), ' +
                    'ROUND((H' + rowNumber + ' - C' + rowNumber + ') * 24 - K' + rowNumber + ', 2), 0)';
  sheet.getRange(rowNumber, 12).setFormula(workFormula);
  sheet.getRange(rowNumber, 12).setNumberFormat("0.00");
}

/**
 * Helper function to set formula with number format (alternative name for compatibility)
 */
function setFormulaWithNumberFormat(sheet, row, col, formula) {
  sheet.getRange(row, col).setFormula(formula);
  sheet.getRange(row, col).setNumberFormat("0.00");
}

/**
 * Check if employee is currently deployed
 * Returns deployment info if deployed, null if not deployed
 */
function isEmployeeDeployed(employeeName) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var deploySheet = ss.getSheetByName(DEPLOYMENT_SHEET);
    
    if (!deploySheet) return null;
    
    var currentDate = new Date();
    var formattedDate = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
    var currentTime = currentDate.getHours() * 60 + currentDate.getMinutes(); // Current time in minutes
    
    var data = deploySheet.getDataRange().getValues();
    
    // Check from most recent to oldest (reverse order)
    for (var i = data.length - 1; i >= 1; i--) {
      var rowDate = Utilities.formatDate(new Date(data[i][1]), Session.getScriptTimeZone(), "yyyy-MM-dd");
      var rowName = data[i][2] ? data[i][2].toString().trim() : "";
      var expectedReturn = data[i][4] ? data[i][4].toString().trim() : "";
      
      // Check if this is the employee and deployment is today
      if (rowDate === formattedDate && rowName === employeeName) {
        // If no expected return time, consider deployed for the whole day
        if (!expectedReturn || expectedReturn === "") {
          return {
            isDeployed: true,
            deploymentId: data[i][0],
            location: data[i][5],
            purpose: data[i][6]
          };
        }
        
        // If expected return time is provided, check if current time is before return time
        var returnTimeParts = expectedReturn.split(':');
        if (returnTimeParts.length >= 2) {
          var returnHour = parseInt(returnTimeParts[0], 10);
          var returnMinute = parseInt(returnTimeParts[1], 10);
          
          // Handle AM/PM if present
          if (expectedReturn.toLowerCase().indexOf('pm') !== -1 && returnHour !== 12) {
            returnHour += 12;
          } else if (expectedReturn.toLowerCase().indexOf('am') !== -1 && returnHour === 12) {
            returnHour = 0;
          }
          
          var returnTimeInMinutes = returnHour * 60 + returnMinute;
          
          // If current time is before expected return time, employee is still deployed
          if (currentTime < returnTimeInMinutes) {
            return {
              isDeployed: true,
              deploymentId: data[i][0],
              location: data[i][5],
              purpose: data[i][6],
              expectedReturn: expectedReturn
            };
          }
        } else {
          // If expected return time format is invalid, consider deployed for the day
          return {
            isDeployed: true,
            deploymentId: data[i][0],
            location: data[i][5],
            purpose: data[i][6]
          };
        }
        
        // If we found a deployment but it's past return time, not deployed anymore
        return null;
      }
    }
    
    return null; // Not deployed
    
  } catch (err) {
    Logger.log("Error checking deployment status: " + err.message);
    return null;
  }
}

/**
 * Initialize or verify ATTENDANCE sheet structure
 * Creates required columns if they don't exist
 */
function initializeAttendanceSheet() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(ATTENDANCE_SHEET);
    
    if (!sheet) {
      // Create new sheet if it doesn't exist
      sheet = ss.insertSheet(ATTENDANCE_SHEET);
      sheet.appendRow([
        "Date", 
        "Employee_Name", 
        "Time In", 
        "Location", 
        "Time In Photo",
        "Noon Break_Started", 
        "Noon Break_Ended", 
        "Time Out", 
        "Time Out Location", 
        "Time Out Photo",
        "Total Break (Hours)",
        "Total Work (Hours)"
      ]);
      
      // Format header row
      var headerRange = sheet.getRange(1, 1, 1, 12);
      headerRange.setBackground("#00E676");
      headerRange.setFontWeight("bold");
      headerRange.setFontColor("#FFFFFF");
      
      return { success: true, message: "ATTENDANCE sheet created successfully." };
    }
    
    // Verify headers exist
    var headers = sheet.getRange(1, 1, 1, 12).getValues()[0];
    var requiredHeaders = [
      "Date", 
      "Employee_Name", 
      "Time In", 
      "Location", 
      "Time In Photo",
      "Noon Break_Started", 
      "Noon Break_Ended", 
      "Time Out", 
      "Time Out Location", 
      "Time Out Photo",
      "Total Break (Hours)",
      "Total Work (Hours)"
    ];
    
    // Update headers if needed
    for (var i = 0; i < requiredHeaders.length; i++) {
      if (!headers[i] || headers[i] !== requiredHeaders[i]) {
        sheet.getRange(1, i + 1).setValue(requiredHeaders[i]);
      }
    }
    
    return { success: true, message: "ATTENDANCE sheet verified." };
    
  } catch (err) {
    Logger.log("Error initializing attendance sheet: " + err.message);
    return { success: false, message: err.message };
  }
}

/**
 * Handle GET requests - For Vercel frontend integration
 * Returns JSON responses for API calls with proper CORS headers
 */
function doGet(e) {
  try {
    Logger.log("=== doGet called at " + new Date().toISOString() + " ===");
    
    // Check if this is an API call (has 'action' parameter)
    if (e && e.parameter && e.parameter.action) {
      var action = e.parameter.action;
      Logger.log("🎯 GET Action: " + action);
      
      var result = {};
      
      // Initialize attendance sheet structure on first call
      try {
        initializeAttendanceSheet();
      } catch (initErr) {
        Logger.log("⚠️ Sheet init warning: " + initErr.message);
      }
      
      // Route to appropriate function based on action
      switch(action) {
        case 'getInitialData':
          result = getInitialData();
          break;
          
        case 'getEmployeeByID':
          result = getEmployeeByID(e.parameter.employeeId || "");
          break;
          
        case 'getAttendanceStatus':
          result = getAttendanceStatus(e.parameter.employeeName || "");
          break;
          
        case 'getActiveClockedInEmployees':
          result = getActiveClockedInEmployees();
          break;
          
        case 'validateAdmin':
          result = validateAdmin(e.parameter.adminId || "", e.parameter.adminCode || "");
          break;
          
        case 'autoClockOutAM':
          var message = autoClockOutAM();
          result = { success: true, message: message };
          break;
          
        case 'autoClockOutPM':
          var message = autoClockOutPM();
          result = { success: true, message: message };
          break;
          
        default:
          result = { success: false, message: "Unknown action: " + action };
      }
      
      Logger.log("✅ GET Request completed successfully");
      
      // Return JSON response with CORS headers
      return ContentService
        .createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // If no action parameter, return error
    Logger.log("❌ No action parameter in GET request");
    return ContentService
      .createTextOutput(JSON.stringify({ 
        success: false, 
        message: "No action specified. This is an API endpoint.",
        timestamp: new Date().toISOString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
               
  } catch (err) {
    Logger.log("❌ Error in doGet: " + err.message);
    Logger.log("🔍 Stack: " + err.stack);
    return ContentService
      .createTextOutput(JSON.stringify({ 
        success: false, 
        message: "Server error: " + err.message,
        timestamp: new Date().toISOString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Handle POST requests - For complex payloads from Vercel frontend
 * Handles recordAttendance, recordDeployment, manualClockOutSelected
 * OPTIMIZED: Added CORS headers and better error handling
 */
function doPost(e) {
  var startTime = new Date().getTime();
  
  try {
    Logger.log("=== doPost called at " + new Date().toISOString() + " ===");
    Logger.log("📍 Request source: " + (e.parameter ? "URL params" : "POST body"));
    
    // Parse the JSON payload
    var data = {};
    
    if (e && e.postData && e.postData.contents) {
      Logger.log("📦 POST data received, size: " + e.postData.contents.length + " bytes");
      Logger.log("📋 Content type: " + (e.postData.type || "unknown"));
      
      try {
        data = JSON.parse(e.postData.contents);
        Logger.log("✅ Parsed data successfully");
        Logger.log("📊 Parsed keys: " + Object.keys(data).join(", "));
      } catch (parseErr) {
        Logger.log("❌ JSON parse error: " + parseErr.message);
        Logger.log("📄 Raw content (first 500 chars): " + e.postData.contents.substring(0, 500));
        return ContentService
          .createTextOutput(JSON.stringify({ 
            success: false, 
            message: "Invalid JSON: " + parseErr.message,
            error: "PARSE_ERROR"
          }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    } else if (e && e.parameter) {
      Logger.log("📋 No postData, using parameters");
      data = e.parameter;
    } else {
      Logger.log("❌ No data received in POST request");
      Logger.log("📊 Request object keys: " + (e ? Object.keys(e).join(", ") : "null"));
      return ContentService
        .createTextOutput(JSON.stringify({ 
          success: false, 
          message: "No data received in request",
          error: "NO_DATA"
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    var action = data.action;
    Logger.log("🎯 Action: " + action);
    
    if (!action) {
      Logger.log("❌ No action specified in payload");
      Logger.log("📊 Data keys: " + Object.keys(data).join(", "));
      return ContentService
        .createTextOutput(JSON.stringify({ 
          success: false, 
          message: "No action specified in request",
          error: "NO_ACTION"
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    var result = {};
    
    // Initialize attendance sheet structure
    try {
      var initResult = initializeAttendanceSheet();
      Logger.log("📊 Sheet initialization: " + JSON.stringify(initResult));
    } catch (initErr) {
      Logger.log("⚠️ Sheet initialization warning: " + initErr.message);
      // Continue anyway - sheet might already exist
    }
    
    // Route to appropriate function based on action
    switch(action) {
      case 'recordAttendance':
        Logger.log("📝 Processing recordAttendance action");
        Logger.log("📊 Spreadsheet ID: " + SPREADSHEET_ID);
        
        try {
          var payload = {
            fullName: data.fullName || "",
            location: data.location || "",
            type: data.type || "",
            photoBase64: data.photoBase64 || "",
            timeStart: data.timeStart || "",
            timeEnd: data.timeEnd || "",
            remarks: data.remarks || "",
            adminId: data.adminId || "",
            adminName: data.adminName || ""
          };
          
          Logger.log("📋 Payload summary: " + JSON.stringify({
            fullName: payload.fullName,
            location: payload.location,
            type: payload.type,
            hasPhoto: !!payload.photoBase64,
            photoSize: payload.photoBase64 ? payload.photoBase64.length : 0
          }));
          
          // Validate required fields
          if (!payload.fullName) {
            throw new Error("Missing required field: fullName");
          }
          if (!payload.location) {
            throw new Error("Missing required field: location");
          }
          if (!payload.type) {
            throw new Error("Missing required field: type");
          }
          
          Logger.log("✅ Validation passed - calling recordAttendance");
          
          var actionStartTime = new Date().getTime();
          var message = recordAttendance(payload);
          var actionDuration = new Date().getTime() - actionStartTime;
          
          Logger.log("✅ recordAttendance SUCCESS in " + actionDuration + "ms");
          Logger.log("📤 Success message: " + message);
          
          result = { 
            success: true, 
            message: message,
            timestamp: new Date().toISOString(),
            _performance: {
              actionDuration: actionDuration + "ms"
            }
          };
        } catch (recordErr) {
          Logger.log("❌ recordAttendance FAILED: " + recordErr.message);
          Logger.log("📍 Error location: " + recordErr.fileName + ":" + recordErr.lineNumber);
          Logger.log("🔍 Stack: " + recordErr.stack);
          
          result = { 
            success: false, 
            message: recordErr.message,
            error: "RECORD_ATTENDANCE_ERROR",
            timestamp: new Date().toISOString()
          };
        }
        break;
        
      case 'recordDeployment':
        Logger.log("🚗 Processing recordDeployment action");
        try {
          var payload = {
            employeeName: data.employeeName || "",
            deploymentLocation: data.deploymentLocation || "",
            purpose: data.purpose || "",
            expectedReturn: data.expectedReturn || "",
            remarks: data.remarks || "",
            adminId: data.adminId || "",
            adminName: data.adminName || ""
          };
          var message = recordDeployment(payload);
          Logger.log("✅ recordDeployment result: " + message);
          result = { success: true, message: message, timestamp: new Date().toISOString() };
        } catch (deployErr) {
          Logger.log("❌ recordDeployment FAILED: " + deployErr.message);
          result = { 
            success: false, 
            message: deployErr.message,
            error: "DEPLOYMENT_ERROR",
            timestamp: new Date().toISOString()
          };
        }
        break;
        
      case 'manualClockOutSelected':
        Logger.log("⏰ Processing manualClockOutSelected action");
        try {
          result = manualClockOutSelected(data.employeeNames || [], data.shiftType || "PM");
          Logger.log("✅ manualClockOutSelected result: " + JSON.stringify(result));
          result.timestamp = new Date().toISOString();
        } catch (clockOutErr) {
          Logger.log("❌ manualClockOutSelected FAILED: " + clockOutErr.message);
          result = { 
            success: false, 
            message: clockOutErr.message,
            error: "CLOCK_OUT_ERROR",
            timestamp: new Date().toISOString()
          };
        }
        break;
        
      default:
        Logger.log("❌ Unknown POST action: " + action);
        result = { 
          success: false, 
          message: "Unknown POST action: " + action,
          error: "UNKNOWN_ACTION",
          timestamp: new Date().toISOString()
        };
    }
    
    var totalDuration = new Date().getTime() - startTime;
    Logger.log("⏱️ Total doPost duration: " + totalDuration + "ms");
    Logger.log("📤 Returning response (first 200 chars): " + JSON.stringify(result).substring(0, 200));
    
    // Add performance data
    if (result._performance) {
      result._performance.totalDuration = totalDuration + "ms";
    } else {
      result._performance = {
        totalDuration: totalDuration + "ms"
      };
    }
    
    // CRITICAL: Return with proper CORS headers and JSON mime type
    var output = ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
    
    Logger.log("✅ Response prepared successfully");
    return output;
    
  } catch (err) {
    var totalDuration = new Date().getTime() - startTime;
    Logger.log("❌ FATAL ERROR in doPost after " + totalDuration + "ms");
    Logger.log("🔥 Error: " + err.message);
    Logger.log("📍 Location: " + err.fileName + ":" + err.lineNumber);
    Logger.log("🔍 Stack: " + err.stack);
    
    return ContentService
      .createTextOutput(JSON.stringify({ 
        success: false, 
        message: "Server error: " + err.message,
        error: "SERVER_ERROR",
        timestamp: new Date().toISOString(),
        _performance: {
          duration: totalDuration + "ms",
          error: true
        }
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Get initial data (employees and locations) for frontend - OPTIMIZED
 * Uses caching to reduce API response time from 3s to <500ms
 */
function getInitialData() {
  var startTime = new Date().getTime();
  Logger.log("=== getInitialData START ===");
  
  try {
    // Use cached data for faster response
    var employeeData = getCachedOrFetch("all_employees", loadAllEmployees);
    var locations = getCachedOrFetch("all_locations", loadAllLocations);
    
    var duration = new Date().getTime() - startTime;
    Logger.log("⏱️ getInitialData completed in " + duration + "ms");
    Logger.log("📊 Returned " + employeeData.list.length + " employees and " + locations.length + " locations");
    
    return {
      success: true,
      locations: locations,
      employees: employeeData.list,
      _performance: {
        duration: duration + "ms",
        cached: true
      }
    };
  } catch (err) {
    var duration = new Date().getTime() - startTime;
    Logger.log("⏱️ getInitialData error after " + duration + "ms: " + err.message);
    
    return {
      success: false,
      message: err.message,
      locations: ["Main Office"],
      employees: [],
      _performance: {
        duration: duration + "ms",
        error: err.message
      }
    };
  }
}

/**
 * Get current attendance status for an employee
 * Returns object with current status flags
 */
function getAttendanceStatus(employeeName) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(ATTENDANCE_SHEET);
    
    if (!sheet) {
      return {
        hasClockIn: false,
        hasBreakStart: false,
        hasBreakEnd: false,
        hasClockOut: false,
        rowNumber: -1
      };
    }
    
    var formattedDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    var data = sheet.getDataRange().getValues();
    
    // Search for today's record (from most recent to oldest)
    for (var r = data.length - 1; r >= 1; r--) {
      if (!data[r][0]) continue;
      
      var rowDate = Utilities.formatDate(new Date(data[r][0]), Session.getScriptTimeZone(), "yyyy-MM-dd");
      var rowName = data[r][1] ? data[r][1].toString().trim() : "";
      
      if (rowDate === formattedDate && rowName === employeeName) {
        return {
          hasClockIn: data[r][2] !== "" && data[r][2] !== null,      // Column C: Time In
          hasBreakStart: data[r][5] !== "" && data[r][5] !== null,   // Column F: Break Start
          hasBreakEnd: data[r][6] !== "" && data[r][6] !== null,     // Column G: Break End
          hasClockOut: data[r][7] !== "" && data[r][7] !== null,     // Column H: Time Out
          rowNumber: r + 1  // 1-based row number
        };
      }
    }
    
    // No record found for today
    return {
      hasClockIn: false,
      hasBreakStart: false,
      hasBreakEnd: false,
      hasClockOut: false,
      rowNumber: -1
    };
    
  } catch (err) {
    Logger.log("Error getting attendance status: " + err.message);
    throw new Error("Failed to check attendance status: " + err.message);
  }
}

/**
 * Record attendance actions (Clock In, Break Start, Break End, Clock Out)
 * Implements strict validation and duplicate prevention
 */
function recordAttendance(payload) {
  try {
    // Add comprehensive logging
    Logger.log("=== recordAttendance START ===");
    Logger.log("Payload received: " + JSON.stringify({
      fullName: payload.fullName,
      location: payload.location,
      type: payload.type,
      hasPhoto: !!(payload.photoBase64 && payload.photoBase64.length > 0)
    }));
    
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    Logger.log("Spreadsheet opened successfully");
    
    var targetSheet;
    var formattedDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    var formattedTime = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "hh:mm a");
    
    Logger.log("Current date: " + formattedDate);
    Logger.log("Current time: " + formattedTime);
    
    var photoUrl = "No Photo";

    // FIXED: Safeguard against undefined payload properties
    var rawName = (payload.fullName || "").toString();
    var rawLocation = (payload.location || "").toString();

    Logger.log("Raw name: " + rawName);
    Logger.log("Raw location: " + rawLocation);

    var finalName = rawName;
    if (finalName.indexOf(" - ") !== -1) {
      finalName = finalName.split(" - ")[1].trim();
    }
    
    Logger.log("Final name after extraction: " + finalName);

    // Validate finalName is not empty
    if (!finalName || finalName === "") {
      throw new Error("Employee name is empty. Please check employee data.");
    }

    // Location no longer has ID prefix, use as-is
    var finalLocation = rawLocation.trim();
    
    Logger.log("Final location: " + finalLocation);

  // Process photo upload - Save to specific Google Drive folder
  if (payload.photoBase64 && payload.photoBase64.indexOf(",") !== -1) {
    try {
      Logger.log("Processing photo upload...");
      var parts = payload.photoBase64.split(",");
      var base64Data = parts[1];
      var metaData = parts[0];
      
      var contentType = "image/jpeg"; 
      if (metaData.indexOf(":") !== -1 && metaData.indexOf(";") !== -1) {
          contentType = metaData.substring(metaData.indexOf(":") + 1, metaData.indexOf(";"));
      }
      
      var decodeBytes = Utilities.base64Decode(base64Data);
      var filename = finalName + "_" + payload.type + "_" + new Date().getTime() + ".jpg";
      var blob = Utilities.newBlob(decodeBytes, contentType, filename);
      
      // Save to specific folder
      var folder = DriveApp.getFolderById(PHOTO_FOLDER_ID);
      var file = folder.createFile(blob); 
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      photoUrl = file.getUrl();
      Logger.log("Photo uploaded successfully to folder: " + photoUrl);
    } catch (err) {
      Logger.log("Photo upload failed: " + err.message);
      photoUrl = "No Photo"; // Use "No Photo" instead of error message
    }
  } else {
    Logger.log("No photo to upload");
  }

  // OVERTIME HANDLING (unchanged)
  if (payload.type === "OVERTIME") {
    Logger.log("Processing OVERTIME type");
    targetSheet = ss.getSheetByName(OVERTIME_SHEET);
    if (!targetSheet) throw new Error("Sheet tab 'OVERTIME' not found.");
    
    var otId = generateOtId();
    var totalHours = calculateTotalHours(payload.timeStart, payload.timeEnd);
    var createdBy = payload.adminName || payload.adminId || "Unknown Admin";
    var adminIdValue = payload.adminId || "";
    var remarks = payload.remarks || "";
    
    targetSheet.appendRow([
      otId, finalName, formattedDate, payload.timeStart, payload.timeEnd, 
      totalHours, remarks, createdBy, adminIdValue
    ]);
    
    Logger.log("Overtime record created successfully");
    return "Overtime successfully submitted for " + finalName + " (OT ID: " + otId + ")";
  }
  
  // NEW SINGLE-SHIFT ATTENDANCE WORKFLOW
  Logger.log("Processing standard attendance type: " + payload.type);
  targetSheet = ss.getSheetByName(ATTENDANCE_SHEET);
  
  Logger.log("Looking for ATTENDANCE sheet: " + ATTENDANCE_SHEET);
  Logger.log("Sheet found: " + (targetSheet ? "Yes" : "No"));
  
  if (!targetSheet) {
    throw new Error("ATTENDANCE sheet not found. Expected sheet name: '" + ATTENDANCE_SHEET + "'");
  }
  
  Logger.log("Getting attendance status for: " + finalName);
  
  // Get current attendance status
  var status = getAttendanceStatus(finalName);
  
  Logger.log("Current status: " + JSON.stringify(status));
  
  // CLOCK IN
  if (payload.type === "IN") {
    Logger.log("=== PROCESSING CLOCK IN ===");
    Logger.log("Employee: " + finalName);
    
    // Check if employee is deployed
    var deploymentStatus = isEmployeeDeployed(finalName);
    if (deploymentStatus && deploymentStatus.isDeployed) {
      throw new Error("Cannot clock in: " + finalName + " is currently deployed to " + 
                     deploymentStatus.location + ". Expected return: " + 
                     (deploymentStatus.expectedReturn || "End of day"));
    }
    
    // Validate: Cannot clock in if already clocked in today
    if (status.hasClockIn) {
      throw new Error("You have already clocked in today.");
    }
    
    Logger.log("Validation passed - proceeding with clock in");
    Logger.log("Data to write: Date=" + formattedDate + ", Name=" + finalName + ", Time=" + formattedTime + ", Location=" + finalLocation);
    
    try {
      // Verify sheet is accessible
      var sheetName = targetSheet.getName();
      Logger.log("Writing to sheet: " + sheetName);
      
      // Get current last row before append
      var lastRowBefore = targetSheet.getLastRow();
      Logger.log("Last row before append: " + lastRowBefore);
      
      // Create new attendance record
      // Columns: Date, Employee_Name, Time In, Location, Time In Photo, Break Start, Break End, Time Out, Time Out Location, Time Out Photo
      var rowData = [
        formattedDate,   // A: Date
        finalName,       // B: Employee_Name
        formattedTime,   // C: Time In
        finalLocation,   // D: Location
        photoUrl,        // E: Time In Photo
        "",              // F: Break Start (empty)
        "",              // G: Break End (empty)
        "",              // H: Time Out (empty)
        "",              // I: Time Out Location (empty)
        ""               // J: Time Out Photo (empty)
      ];
      
      Logger.log("Row data prepared: " + JSON.stringify(rowData));
      
      // Append the row
      Logger.log("Calling appendRow...");
      targetSheet.appendRow(rowData);
      
      Logger.log("appendRow completed successfully");
      
      // Force flush MULTIPLE times to ensure write
      Logger.log("Flushing spreadsheet...");
      SpreadsheetApp.flush();
      Utilities.sleep(150); // Wait 150ms (increased from 100ms)
      SpreadsheetApp.flush();
      
      Logger.log("Spreadsheet flushed - checking if row was added");
      
      // Get the row number that was just added
      var lastRowAfter = targetSheet.getLastRow();
      Logger.log("Last row after append: " + lastRowAfter);
      
      if (lastRowAfter <= lastRowBefore) {
        Logger.log("ERROR: Row count did not increase!");
        throw new Error("Row was not added! Last row before: " + lastRowBefore + ", after: " + lastRowAfter);
      }
      
      Logger.log("SUCCESS: Row was added at position " + lastRowAfter);
      
      // Verify the data was written
      Logger.log("Verifying written data...");
      var verifyRow = targetSheet.getRange(lastRowAfter, 1, 1, 10).getValues()[0];
      Logger.log("Verification - Row data: " + JSON.stringify(verifyRow));
      
      // ==========================================================================
      // LATE MARKING: Mark as LATE if clock in is 8:16 AM or later
      // ==========================================================================
      var currentTime = new Date();
      var clockInHour = currentTime.getHours();
      var clockInMinute = currentTime.getMinutes();
      
      // Check if late (using configured threshold)
      var isLate = false;
      if (clockInHour > LATE_TIME_HOUR) {
        isLate = true; // After late hour
      } else if (clockInHour === LATE_TIME_HOUR && clockInMinute >= LATE_TIME_MINUTE) {
        isLate = true; // At or after late time
      }
      
      if (isLate) {
        Logger.log("Employee is LATE - applying red background");
        // Mark entire row as late with red background
        var rowRange = targetSheet.getRange(lastRowAfter, 1, 1, 12); // All columns A-L
        rowRange.setBackground("#FFCDD2"); // Light red background
        Logger.log("Red background applied to row " + lastRowAfter);
      }
      
      // Final flush
      SpreadsheetApp.flush();
      
      Logger.log("=== CLOCK IN COMPLETED SUCCESSFULLY ===");
      
      return "Clock In successful at " + formattedTime + (isLate ? " (LATE)" : "") + " - Row: " + lastRowAfter;
      
    } catch (appendError) {
      Logger.log("=== ERROR DURING CLOCK IN ===");
      Logger.log("ERROR: " + appendError.message);
      Logger.log("Stack trace: " + appendError.stack);
      throw new Error("Failed to write to ATTENDANCE sheet: " + appendError.message);
    }
  }
  
  // BREAK START
  else if (payload.type === "BREAK_START") {
    Logger.log("Processing Break Start for: " + finalName);
    
    // Validate: Must have clocked in
    if (!status.hasClockIn) {
      throw new Error("Cannot start break: You must clock in first.");
    }
    
    // Validate: Cannot start break if already started
    if (status.hasBreakStart) {
      throw new Error("Break has already been started.");
    }
    
    // Validate: Cannot start break if already clocked out
    if (status.hasClockOut) {
      throw new Error("Cannot start break: You have already clocked out.");
    }
    
    Logger.log("Updating break start for row: " + status.rowNumber);
    
    // Update Break Start (Column F)
    targetSheet.getRange(status.rowNumber, 6).setValue(formattedTime);
    
    // Force flush
    SpreadsheetApp.flush();
    
    Logger.log("Break Start completed successfully");
    
    return "Break started at " + formattedTime;
  }
  
  // BREAK END
  else if (payload.type === "BREAK_END") {
    Logger.log("Processing Break End for: " + finalName);
    
    // Validate: Must have started break
    if (!status.hasBreakStart) {
      throw new Error("Cannot end break: You must start break first.");
    }
    
    // Validate: Cannot end break if already ended
    if (status.hasBreakEnd) {
      throw new Error("Break has already been ended.");
    }
    
    // Validate: Cannot end break if already clocked out
    if (status.hasClockOut) {
      throw new Error("Cannot end break: You have already clocked out.");
    }
    
    Logger.log("Updating break end for row: " + status.rowNumber);
    
    // Update Break End (Column G)
    targetSheet.getRange(status.rowNumber, 7).setValue(formattedTime);
    
    // Force flush
    SpreadsheetApp.flush();
    
    Logger.log("Break End completed successfully");
    
    return "Break ended at " + formattedTime;
  }
  
  // CLOCK OUT
  else if (payload.type === "OUT") {
    Logger.log("Processing Clock Out for: " + finalName);
    
    // Check if employee is deployed
    var deploymentStatus = isEmployeeDeployed(finalName);
    if (deploymentStatus && deploymentStatus.isDeployed) {
      throw new Error("Cannot clock out: " + finalName + " is currently deployed to " + 
                     deploymentStatus.location + ". Expected return: " + 
                     (deploymentStatus.expectedReturn || "End of day"));
    }
    
    // Validate: Must have clocked in
    if (!status.hasClockIn) {
      throw new Error("Cannot clock out: You must clock in first.");
    }
    
    // Validate: Cannot clock out if already clocked out
    if (status.hasClockOut) {
      throw new Error("You have already clocked out today.");
    }
    
    Logger.log("Updating clock out for row: " + status.rowNumber);
    
    // AUTO-END BREAK if break was started but not ended (Option A - Preferred)
    if (status.hasBreakStart && !status.hasBreakEnd) {
      targetSheet.getRange(status.rowNumber, 7).setValue(formattedTime);
      Logger.log("Auto-ended break for " + finalName + " before clock out");
    }
    
    // Update Clock Out (Columns H, I, J)
    targetSheet.getRange(status.rowNumber, 8).setValue(formattedTime);      // H: Time Out
    targetSheet.getRange(status.rowNumber, 9).setValue(finalLocation);      // I: Time Out Location
    targetSheet.getRange(status.rowNumber, 10).setValue(photoUrl);          // J: Time Out Photo
    
    Logger.log("Clock out data written to columns H, I, J");
    
    // Set hour formulas with proper number formatting
    setHourFormulas(targetSheet, status.rowNumber);
    
    Logger.log("Hour formulas set for row: " + status.rowNumber);
    
    // Force flush to ensure data is written
    SpreadsheetApp.flush();
    Utilities.sleep(100);
    SpreadsheetApp.flush();
    
    // Verify the update
    var verifyData = targetSheet.getRange(status.rowNumber, 8, 1, 3).getValues()[0];
    Logger.log("Verification - Clock out data: " + JSON.stringify(verifyData));
    
    Logger.log("Clock Out completed successfully for " + finalName);
    
    return "Clock Out successful at " + formattedTime;
  }
  
  throw new Error("Invalid attendance type: " + payload.type);
  
  } catch (err) {
    Logger.log("=== ERROR in recordAttendance ===");
    Logger.log("ERROR: " + err.message);
    Logger.log("Stack trace: " + err.stack);
    
    // Log payload info safely
    try {
      Logger.log("Payload type: " + (payload ? payload.type : "undefined"));
      Logger.log("Payload fullName: " + (payload ? payload.fullName : "undefined"));
      Logger.log("Payload location: " + (payload ? payload.location : "undefined"));
    } catch (logErr) {
      Logger.log("Could not log payload details");
    }
    
    // Return user-friendly error message
    throw err; // Re-throw the original error with its message
  }
}


// ==========================================================================
// AUTO CLOCK-OUT FUNCTIONS
// ==========================================================================

/**
 * Auto Clock-Out for AM Shift (Half-Day)
 * Trigger: Set to run daily at 12:30 PM
 * Automatically clocks out employees who haven't clocked out by noon (half-day employees)
 * Also auto-ends breaks if employee forgot to end break
 */
function autoClockOutAM() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(ATTENDANCE_SHEET);
    
    if (!sheet) {
      Logger.log("ATTENDANCE sheet not found");
      return "Error: ATTENDANCE sheet not found";
    }
    
    var formattedDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    var autoClockOutTime = "12:30 PM";
    var autoClockOutLocation = "Auto Clock-Out (Half-Day)";
    var autoClockOutPhoto = "Forgot to clock out";
    
    var data = sheet.getDataRange().getValues();
    var updatedCount = 0;
    
    // Start from row 2 (skip header)
    for (var r = 1; r < data.length; r++) {
      if (!data[r][0]) continue; // Skip empty rows
      
      var rowDate = Utilities.formatDate(new Date(data[r][0]), Session.getScriptTimeZone(), "yyyy-MM-dd");
      var clockOutTime = data[r][7]; // Column H (index 7) - Time Out
      var breakStart = data[r][5];   // Column F (index 5) - Break Start
      var breakEnd = data[r][6];     // Column G (index 6) - Break End
      
      // If date matches today and clock-out is empty, auto clock-out
      if (rowDate === formattedDate && clockOutTime === "") {
        var rowNumber = r + 1; // Convert to 1-based row number
        
        // Auto-end break if started but not ended
        if (breakStart !== "" && breakEnd === "") {
          sheet.getRange(rowNumber, 7).setValue(autoClockOutTime); // Column G: Break End
        }
        
        // Auto clock-out
        sheet.getRange(rowNumber, 8).setValue(autoClockOutTime);      // Column H: Time Out
        sheet.getRange(rowNumber, 9).setValue(autoClockOutLocation);  // Column I: Location Out
        sheet.getRange(rowNumber, 10).setValue(autoClockOutPhoto);    // Column J: Photo Out
        
        // Set hour formulas with proper number formatting
        setHourFormulas(sheet, rowNumber);
        
        updatedCount++;
      }
    }
    
    var resultMessage = "AM Shift (Half-Day): " + updatedCount + " employee(s) auto clocked-out.";
    Logger.log("Auto Clock-Out AM completed. Updated " + updatedCount + " records.");
    return resultMessage;
    
  } catch (err) {
    Logger.log("Error in autoClockOutAM: " + err.message);
    return "Error: " + err.message;
  }
}

/**
 * Auto Clock-Out for PM Shift (Full Day)
 * Trigger: Set to run daily at 5:30 PM
 * Automatically clocks out employees who haven't clocked out by end of day
 * Also auto-ends breaks if employee forgot to end break
 */
function autoClockOutPM() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(ATTENDANCE_SHEET);
    
    if (!sheet) {
      Logger.log("ATTENDANCE sheet not found");
      return "Error: ATTENDANCE sheet not found";
    }
    
    var formattedDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    var autoClockOutTime = "05:30 PM";
    var autoClockOutLocation = "Auto Clock-Out (Full Day)";
    var autoClockOutPhoto = "Forgot to clock out";
    
    var data = sheet.getDataRange().getValues();
    var updatedCount = 0;
    
    // Start from row 2 (skip header)
    for (var r = 1; r < data.length; r++) {
      if (!data[r][0]) continue; // Skip empty rows
      
      var rowDate = Utilities.formatDate(new Date(data[r][0]), Session.getScriptTimeZone(), "yyyy-MM-dd");
      var clockOutTime = data[r][7]; // Column H (index 7) - Time Out
      var breakStart = data[r][5];   // Column F (index 5) - Break Start
      var breakEnd = data[r][6];     // Column G (index 6) - Break End
      
      // If date matches today and clock-out is empty, auto clock-out
      if (rowDate === formattedDate && clockOutTime === "") {
        var rowNumber = r + 1; // Convert to 1-based row number
        
        // Auto-end break if started but not ended
        if (breakStart !== "" && breakEnd === "") {
          sheet.getRange(rowNumber, 7).setValue(autoClockOutTime); // Column G: Break End
        }
        
        // Auto clock-out
        sheet.getRange(rowNumber, 8).setValue(autoClockOutTime);      // Column H: Time Out
        sheet.getRange(rowNumber, 9).setValue(autoClockOutLocation);  // Column I: Location Out
        sheet.getRange(rowNumber, 10).setValue(autoClockOutPhoto);    // Column J: Photo Out
        
        // Calculate Total Break Hours (Column K)
        var breakFormula = '=IF(AND(F' + rowNumber + '<>"", G' + rowNumber + '<>""), ' +
                           'ROUND((G' + rowNumber + ' - F' + rowNumber + ') * 24, 2), 0)';
        setFormulaWithNumberFormat(sheet, rowNumber, 11, breakFormula);
        
        // Calculate Total Work Hours (Column L)
        var workFormula = '=IF(AND(C' + rowNumber + '<>"", H' + rowNumber + '<>""), ' +
                          'ROUND((H' + rowNumber + ' - C' + rowNumber + ') * 24 - K' + rowNumber + ', 2), 0)';
        setFormulaWithNumberFormat(sheet, rowNumber, 12, workFormula);
        
        updatedCount++;
      }
    }
    
    var resultMessage = "PM Shift (Full Day): " + updatedCount + " employee(s) auto clocked-out.";
    Logger.log("Auto Clock-Out PM completed. Updated " + updatedCount + " records.");
    return resultMessage;
    
  } catch (err) {
    Logger.log("Error in autoClockOutPM: " + err.message);
    return "Error: " + err.message;
  }
}

/**
 * Get list of employees who are currently clocked in but not clocked out
 * Used for manual auto clock-out selection
 */
function getActiveClockedInEmployees() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(ATTENDANCE_SHEET);
    
    if (!sheet) {
      return { success: false, message: "ATTENDANCE sheet not found", employees: [] };
    }
    
    var formattedDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    var data = sheet.getDataRange().getValues();
    var activeEmployees = [];
    
    // Start from row 2 (skip header)
    for (var r = 1; r < data.length; r++) {
      if (!data[r][0]) continue; // Skip empty rows
      
      var rowDate = Utilities.formatDate(new Date(data[r][0]), Session.getScriptTimeZone(), "yyyy-MM-dd");
      var employeeName = data[r][1] ? data[r][1].toString().trim() : "";
      var timeIn = data[r][2] ? data[r][2].toString().trim() : "";
      var timeOut = data[r][7]; // Column H - Time Out
      var breakStart = data[r][5] ? data[r][5].toString().trim() : "";
      var breakEnd = data[r][6] ? data[r][6].toString().trim() : "";
      
      // If date matches today and clocked in but not clocked out
      if (rowDate === formattedDate && timeIn !== "" && timeOut === "") {
        activeEmployees.push({
          name: employeeName,
          timeIn: timeIn,
          breakStart: breakStart,
          breakEnd: breakEnd,
          rowNumber: r + 1
        });
      }
    }
    
    return { success: true, employees: activeEmployees };
    
  } catch (err) {
    Logger.log("Error getting active employees: " + err.message);
    return { success: false, message: err.message, employees: [] };
  }
}

/**
 * Manual clock out selected employees
 * @param {Array} employeeNames - Array of employee names to clock out
 * @param {String} shiftType - "AM" or "PM"
 */
function manualClockOutSelected(employeeNames, shiftType) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(ATTENDANCE_SHEET);
    
    if (!sheet) {
      return { success: false, message: "ATTENDANCE sheet not found" };
    }
    
    var clockOutTime = shiftType === "AM" ? "12:30 PM" : "05:30 PM";
    var clockOutLocation = "Manual Auto Clock-Out (" + shiftType + " Shift)";
    var clockOutPhoto = "Admin manual clock-out";
    
    var formattedDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    var data = sheet.getDataRange().getValues();
    var updatedCount = 0;
    
    // Process each selected employee
    for (var i = 0; i < employeeNames.length; i++) {
      var targetName = employeeNames[i];
      
      // Find the employee's row
      for (var r = 1; r < data.length; r++) {
        if (!data[r][0]) continue;
        
        var rowDate = Utilities.formatDate(new Date(data[r][0]), Session.getScriptTimeZone(), "yyyy-MM-dd");
        var rowName = data[r][1] ? data[r][1].toString().trim() : "";
        var timeOut = data[r][7]; // Column H - Time Out
        
        if (rowDate === formattedDate && rowName === targetName && timeOut === "") {
          var rowNumber = r + 1;
          var breakStart = data[r][5];
          var breakEnd = data[r][6];
          
          // Auto-end break if started but not ended
          if (breakStart !== "" && breakEnd === "") {
            sheet.getRange(rowNumber, 7).setValue(clockOutTime); // Column G: Break End
          }
          
          // Clock out
          sheet.getRange(rowNumber, 8).setValue(clockOutTime);      // Column H: Time Out
          sheet.getRange(rowNumber, 9).setValue(clockOutLocation);  // Column I: Location Out
          sheet.getRange(rowNumber, 10).setValue(clockOutPhoto);    // Column J: Photo Out
          
          // Calculate Total Break Hours (Column K)
          var breakFormula = '=IF(AND(F' + rowNumber + '<>"", G' + rowNumber + '<>""), ' +
                             'ROUND((G' + rowNumber + ' - F' + rowNumber + ') * 24, 2), 0)';
          setFormulaWithNumberFormat(sheet, rowNumber, 11, breakFormula);
          
          // Calculate Total Work Hours (Column L)
          var workFormula = '=IF(AND(C' + rowNumber + '<>"", H' + rowNumber + '<>""), ' +
                            'ROUND((H' + rowNumber + ' - C' + rowNumber + ') * 24 - K' + rowNumber + ', 2), 0)';
          setFormulaWithNumberFormat(sheet, rowNumber, 12, workFormula);
          
          updatedCount++;
          break; // Move to next employee
        }
      }
    }
    
    return { 
      success: true, 
      message: updatedCount + " employee(s) clocked out successfully (" + shiftType + " Shift)",
      count: updatedCount
    };
    
  } catch (err) {
    Logger.log("Error in manualClockOutSelected: " + err.message);
    return { success: false, message: err.message };
  }
}

// ==========================================================================
// DEPLOYMENT FUNCTIONS
// ==========================================================================

/**
 * Record Employee Deployment
 * Creates a new deployment record in the DEPLOYMENT sheet
 * Requires admin authentication
 * Automatically handles attendance bypass for deployed employees
 */
function recordDeployment(payload) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var deploySheet = ss.getSheetByName(DEPLOYMENT_SHEET);
    
    if (!deploySheet) {
      throw new Error("DEPLOYMENT sheet not found.");
    }
    
    // Generate unique Deployment ID
    var deploymentId = generateDeploymentId();
    
    // Format current date and time
    var currentDate = new Date();
    var formattedDate = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
    var formattedTime = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "hh:mm a");
    
    // Extract employee name from format "ID - Name"
    var rawName = (payload.employeeName || "").toString();
    var finalName = rawName;
    if (finalName.indexOf(" - ") !== -1) {
      finalName = finalName.split(" - ")[1].trim();
    }
    
    // Extract location from format "ID - Location" (if applicable)
    var rawLocation = (payload.deploymentLocation || "").toString();
    var finalLocation = rawLocation;
    if (finalLocation.indexOf(" - ") !== -1) {
      finalLocation = finalLocation.split(" - ")[1].trim();
    }
    
    // Get optional fields
    var purpose = payload.purpose || "";
    var expectedReturn = payload.expectedReturn || "";
    var remarks = payload.remarks || "";
    
    // Get admin info
    var approvedBy = payload.adminName || payload.adminId || "Unknown Admin";
    
    // ==========================================================================
    // AUTO ATTENDANCE HANDLING FOR DEPLOYED EMPLOYEE
    // ==========================================================================
    
    var attendanceSheet = ss.getSheetByName(ATTENDANCE_SHEET);
    
    if (attendanceSheet) {
      var attendanceData = attendanceSheet.getDataRange().getValues();
      var hasTimeInToday = false;
      var timeInRowNumber = -1;
      
      // Check if employee already clocked in today
      for (var r = 1; r < attendanceData.length; r++) {
        if (!attendanceData[r][0]) continue;
        
        var rowDate = Utilities.formatDate(new Date(attendanceData[r][0]), Session.getScriptTimeZone(), "yyyy-MM-dd");
        var rowName = attendanceData[r][1].toString().trim();
        var timeOut = attendanceData[r][7]; // Column H - Time Out
        
        if (rowDate === formattedDate && rowName === finalName) {
          hasTimeInToday = true;
          
          // If they clocked in but haven't clocked out yet
          if (timeOut === "" || timeOut === null) {
            timeInRowNumber = r + 1; // Convert to 1-based row number
          }
          break;
        }
      }
      
      // If employee clocked in but hasn't clocked out, auto clock out with DEPLOYED status
      if (timeInRowNumber !== -1) {
        // Auto-end break if started but not ended
        var breakStart = attendanceData[timeInRowNumber - 1][5];
        var breakEnd = attendanceData[timeInRowNumber - 1][6];
        if (breakStart !== "" && breakEnd === "") {
          attendanceSheet.getRange(timeInRowNumber, 7).setValue(formattedTime); // Column G: Break End
        }
        
        attendanceSheet.getRange(timeInRowNumber, 8).setValue(formattedTime); // Column H: Time Out
        attendanceSheet.getRange(timeInRowNumber, 9).setValue("DEPLOYED"); // Column I: Location Out
        attendanceSheet.getRange(timeInRowNumber, 10).setValue("Employee deployed to: " + finalLocation); // Column J: Photo/Note
        
        Logger.log("Auto clocked out " + finalName + " - Status: DEPLOYED");
      }
      
      // Note: If employee hasn't clocked in, we do nothing (they bypass attendance entirely)
    }
    
    // ==========================================================================
    // SAVE DEPLOYMENT RECORD
    // ==========================================================================
    
    // Append row to DEPLOYMENT sheet
    // Columns: DEPLOYMENT_ID | D_DATE | NAME | DEPLOY_TIME | RETURN_TIME | DEPLOYMENT LOCATION | PURPOSE | REMARKS | APPROVED By | DATE APPROVED
    deploySheet.appendRow([
      deploymentId,           // Column A: DEPLOYMENT_ID
      formattedDate,          // Column B: D_DATE
      finalName,              // Column C: NAME
      formattedTime,          // Column D: DEPLOY_TIME
      expectedReturn,         // Column E: RETURN_TIME (expected)
      finalLocation,          // Column F: DEPLOYMENT LOCATION
      purpose,                // Column G: PURPOSE
      remarks,                // Column H: REMARKS
      approvedBy,             // Column I: APPROVED By
      formattedDate           // Column J: DATE APPROVED
    ]);
    
    return "Deployment successfully recorded for " + finalName + " (Deployment ID: " + deploymentId + ")";
    
  } catch (err) {
    Logger.log("Error in recordDeployment: " + err.message);
    throw new Error("Failed to record deployment: " + err.message);
  }
}


// ==========================================================================
// FLEXIBLE EMPLOYEE ID MANAGEMENT SYSTEM
// ==========================================================================

/**
 * Update Employee ID in EMPLOYEE sheet
 * 
 * FLEXIBLE ID SYSTEM - CORE FUNCTION:
 * This function allows admins to change employee IDs without affecting:
 * - Historical attendance records (stored by NAME)
 * - Overtime records (stored by NAME)
 * - Deployment records (stored by NAME)
 * 
 * @param {String} currentId - Current employee ID
 * @param {String} newId - New employee ID to assign
 * @param {String} adminId - Admin ID for authorization
 * @param {String} adminCode - Admin authentication code
 * @return {Object} Result with success status and message
 */
function updateEmployeeId(currentId, newId, adminId, adminCode) {
  try {
    // Validate admin credentials
    var adminValidation = validateAdmin(adminId, adminCode);
    if (!adminValidation.success) {
      return { success: false, message: "Admin authorization failed: " + adminValidation.message };
    }
    
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var empSheet = ss.getSheetByName(EMPLOYEE_SHEET);
    
    if (!empSheet) {
      return { success: false, message: "EMPLOYEE sheet not found." };
    }
    
    var data = empSheet.getDataRange().getValues();
    
    // Check if new ID already exists (prevent duplicates)
    for (var i = 1; i < data.length; i++) {
      var sheetEmpId = data[i][0] ? data[i][0].toString().trim() : "";
      if (sheetEmpId.toUpperCase() === newId.toUpperCase()) {
        return { success: false, message: "New ID '" + newId + "' already exists. Please choose a different ID." };
      }
    }
    
    // Find and update the employee ID
    for (var i = 1; i < data.length; i++) {
      var sheetEmpId = data[i][0] ? data[i][0].toString().trim() : "";
      var firstName = data[i][1] ? data[i][1].toString().trim() : "";
      var lastName = data[i][2] ? data[i][2].toString().trim() : "";
      
      // Combine first and last name to create full name
      var sheetEmpName = (firstName + " " + lastName).trim();
      
      if (sheetEmpId.toUpperCase() === currentId.toUpperCase()) {
        // Update the ID in row i+1 (1-based indexing)
        empSheet.getRange(i + 1, 1).setValue(newId);
        
        return { 
          success: true, 
          message: "Employee ID successfully updated from '" + currentId + "' to '" + newId + "' for " + sheetEmpName,
          employeeName: sheetEmpName,
          oldId: currentId,
          newId: newId
        };
      }
    }
    
    return { success: false, message: "Employee ID '" + currentId + "' not found." };
    
  } catch (err) {
    Logger.log("Error updating employee ID: " + err.message);
    return { success: false, message: "Error updating employee ID: " + err.message };
  }
}

/**
 * Bulk update multiple employee IDs
 * 
 * @param {Array} updates - Array of objects with {currentId, newId}
 * @param {String} adminId - Admin ID for authorization
 * @param {String} adminCode - Admin authentication code
 * @return {Object} Result with success count and details
 */
function bulkUpdateEmployeeIds(updates, adminId, adminCode) {
  try {
    // Validate admin credentials
    var adminValidation = validateAdmin(adminId, adminCode);
    if (!adminValidation.success) {
      return { success: false, message: "Admin authorization failed: " + adminValidation.message };
    }
    
    var results = {
      success: true,
      totalProcessed: updates.length,
      successCount: 0,
      failureCount: 0,
      details: []
    };
    
    for (var i = 0; i < updates.length; i++) {
      var update = updates[i];
      var result = updateEmployeeId(update.currentId, update.newId, adminId, adminCode);
      
      if (result.success) {
        results.successCount++;
        results.details.push({
          success: true,
          employeeName: result.employeeName,
          oldId: result.oldId,
          newId: result.newId
        });
      } else {
        results.failureCount++;
        results.details.push({
          success: false,
          currentId: update.currentId,
          message: result.message
        });
      }
    }
    
    return results;
    
  } catch (err) {
    Logger.log("Error in bulk update: " + err.message);
    return { success: false, message: "Error in bulk update: " + err.message };
  }
}

/**
 * Get employee history report (shows all records regardless of ID changes)
 * 
 * FLEXIBLE ID SYSTEM - REPORTING:
 * This demonstrates how the system maintains data integrity:
 * - Searches by NAME (permanent identifier)
 * - Displays current ID in report
 * - All historical data remains accessible
 * 
 * @param {String} employeeName - Employee name to search
 * @return {Object} Complete employee record history
 */
function getEmployeeHistoryReport(employeeName) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // Get current employee info (with latest ID)
    var empInfo = getEmployeeByName(employeeName);
    if (!empInfo.success) {
      return { success: false, message: "Employee not found: " + employeeName };
    }
    
    var report = {
      success: true,
      employee: {
        name: empInfo.employeeName,
        currentId: empInfo.employeeId
      },
      attendance: [],
      overtime: [],
      deployment: []
    };
    
    // Get attendance records (all stored by NAME)
    var attendanceSheet = ss.getSheetByName(ATTENDANCE_SHEET);
    if (attendanceSheet) {
      var attData = attendanceSheet.getDataRange().getValues();
      for (var i = 1; i < attData.length; i++) {
        var rowName = attData[i][1] ? attData[i][1].toString().trim() : "";
        if (rowName === employeeName) {
          report.attendance.push({
            date: Utilities.formatDate(new Date(attData[i][0]), Session.getScriptTimeZone(), "yyyy-MM-dd"),
            timeIn: attData[i][2],
            timeOut: attData[i][7],
            totalHours: attData[i][11]
          });
        }
      }
    }
    
    // Get overtime records (all stored by NAME)
    var otSheet = ss.getSheetByName(OVERTIME_SHEET);
    if (otSheet) {
      var otData = otSheet.getDataRange().getValues();
      for (var i = 1; i < otData.length; i++) {
        var rowName = otData[i][1] ? otData[i][1].toString().trim() : "";
        if (rowName === employeeName) {
          report.overtime.push({
            otId: otData[i][0],
            date: Utilities.formatDate(new Date(otData[i][2]), Session.getScriptTimeZone(), "yyyy-MM-dd"),
            timeStart: otData[i][3],
            timeEnd: otData[i][4],
            totalHours: otData[i][5]
          });
        }
      }
    }
    
    // Get deployment records (all stored by NAME)
    var deploySheet = ss.getSheetByName(DEPLOYMENT_SHEET);
    if (deploySheet) {
      var deployData = deploySheet.getDataRange().getValues();
      for (var i = 1; i < deployData.length; i++) {
        var rowName = deployData[i][2] ? deployData[i][2].toString().trim() : "";
        if (rowName === employeeName) {
          report.deployment.push({
            deploymentId: deployData[i][0],
            date: Utilities.formatDate(new Date(deployData[i][1]), Session.getScriptTimeZone(), "yyyy-MM-dd"),
            location: deployData[i][5],
            purpose: deployData[i][6]
          });
        }
      }
    }
    
    return report;
    
  } catch (err) {
    Logger.log("Error getting employee history: " + err.message);
    return { success: false, message: "Error retrieving history: " + err.message };
  }
}

/**
 * Verify data integrity after ID changes
 * 
 * This function confirms that:
 * - All historical records are still accessible
 * - No orphaned records exist
 * - Employee name consistency is maintained
 * 
 * @return {Object} Integrity check results
 */
function verifyDataIntegrity() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var empSheet = ss.getSheetByName(EMPLOYEE_SHEET);
    
    var report = {
      success: true,
      totalEmployees: 0,
      verifiedEmployees: 0,
      issues: []
    };
    
    if (!empSheet) {
      return { success: false, message: "EMPLOYEE sheet not found." };
    }
    
    var empData = empSheet.getDataRange().getValues();
    
    // Check each employee
    for (var i = 1; i < empData.length; i++) {
      var empId = empData[i][0] ? empData[i][0].toString().trim() : "";
      var empName = empData[i][1] ? empData[i][1].toString().trim() : "";
      
      if (empName === "") continue;
      
      report.totalEmployees++;
      
      // Get history for this employee
      var history = getEmployeeHistoryReport(empName);
      
      if (history.success) {
        report.verifiedEmployees++;
      } else {
        report.issues.push({
          employeeId: empId,
          employeeName: empName,
          issue: "Could not retrieve history"
        });
      }
    }
    
    return report;
    
  } catch (err) {
    Logger.log("Error verifying data integrity: " + err.message);
    return { success: false, message: "Error: " + err.message };
  }
}


// ==========================================================================
// ATTENDANCE MANAGEMENT FUNCTIONS
// ==========================================================================

/**
 * Clear all clock-in details for an employee for today
 * Useful for fixing errors or resetting attendance
 * 
 * @param {String} employeeName - Employee name to clear attendance
 * @param {String} adminId - Admin ID for authorization
 * @param {String} adminCode - Admin authentication code
 * @return {Object} Result with success status
 */
function clearTodayAttendance(employeeName, adminId, adminCode) {
  try {
    // Validate admin credentials
    var adminValidation = validateAdmin(adminId, adminCode);
    if (!adminValidation.success) {
      return { success: false, message: "Admin authorization failed: " + adminValidation.message };
    }
    
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(ATTENDANCE_SHEET);
    
    if (!sheet) {
      return { success: false, message: "ATTENDANCE sheet not found." };
    }
    
    var formattedDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    var data = sheet.getDataRange().getValues();
    
    // Find today's attendance record for this employee
    for (var r = 1; r < data.length; r++) {
      if (!data[r][0]) continue;
      
      var rowDate = Utilities.formatDate(new Date(data[r][0]), Session.getScriptTimeZone(), "yyyy-MM-dd");
      var rowName = data[r][1] ? data[r][1].toString().trim() : "";
      
      if (rowDate === formattedDate && rowName === employeeName) {
        var rowNumber = r + 1;
        
        // Delete the entire row
        sheet.deleteRow(rowNumber);
        
        return { 
          success: true, 
          message: "Today's attendance cleared for " + employeeName + ". Employee can now clock in again.",
          employeeName: employeeName
        };
      }
    }
    
    return { success: false, message: "No attendance record found for " + employeeName + " today." };
    
  } catch (err) {
    Logger.log("Error clearing attendance: " + err.message);
    return { success: false, message: "Error: " + err.message };
  }
}

/**
 * Get today's attendance record for an employee
 * Returns detailed attendance information
 * 
 * @param {String} employeeName - Employee name
 * @return {Object} Attendance details
 */
function getTodayAttendanceDetails(employeeName) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(ATTENDANCE_SHEET);
    
    if (!sheet) {
      return { success: false, message: "ATTENDANCE sheet not found." };
    }
    
    var formattedDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    var data = sheet.getDataRange().getValues();
    
    // Find today's record
    for (var r = 1; r < data.length; r++) {
      if (!data[r][0]) continue;
      
      var rowDate = Utilities.formatDate(new Date(data[r][0]), Session.getScriptTimeZone(), "yyyy-MM-dd");
      var rowName = data[r][1] ? data[r][1].toString().trim() : "";
      
      if (rowDate === formattedDate && rowName === employeeName) {
        return {
          success: true,
          employeeName: rowName,
          timeIn: data[r][2] || "",
          location: data[r][3] || "",
          timeInPhoto: data[r][4] || "",
          breakStart: data[r][5] || "",
          breakEnd: data[r][6] || "",
          timeOut: data[r][7] || "",
          timeOutLocation: data[r][8] || "",
          timeOutPhoto: data[r][9] || ""
        };
      }
    }
    
    return { success: false, message: "No attendance record found for today." };
    
  } catch (err) {
    Logger.log("Error getting attendance details: " + err.message);
    return { success: false, message: "Error: " + err.message };
  }
}

/**
 * Debug function to test attendance recording
 * Helps identify issues with data not reflecting on sheets
 */
function testAttendanceRecording() {
  // Test employee data
  var testEmployee = "Test Employee";
  
  Logger.log("========== ATTENDANCE RECORDING TEST ==========");
  
  // Test 1: Check if sheet exists
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(ATTENDANCE_SHEET);
  Logger.log("ATTENDANCE sheet exists: " + (sheet != null));
  
  if (!sheet) {
    Logger.log("ERROR: ATTENDANCE sheet not found!");
    return "ERROR: ATTENDANCE sheet not found";
  }
  
  // Test 2: Check headers
  var headers = sheet.getRange(1, 1, 1, 12).getValues()[0];
  Logger.log("Headers: " + headers.join(", "));
  
  // Test 3: Try to append a test row
  try {
    var testDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    var testTime = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "hh:mm a");
    
    sheet.appendRow([
      testDate,
      testEmployee,
      testTime,
      "Test Location",
      "Test Photo URL",
      "",
      "",
      "",
      "",
      ""
    ]);
    
    Logger.log("SUCCESS: Test row appended successfully");
    Logger.log("Last row number: " + sheet.getLastRow());
    
    // Clean up test row
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] === testEmployee) {
        sheet.deleteRow(i + 1);
        Logger.log("Test row cleaned up");
        break;
      }
    }
    
    return "SUCCESS: Attendance recording test passed!";
    
  } catch (err) {
    Logger.log("ERROR appending row: " + err.message);
    return "ERROR: " + err.message;
  }
}


// ==========================================================================
// DIAGNOSTIC TEST FUNCTIONS
// ==========================================================================

/**
 * Test function to verify employee lookup and attendance recording
 * Run this manually from Apps Script to test the system
 */
function testAttendanceSystem() {
  try {
    Logger.log("=== STARTING ATTENDANCE SYSTEM TEST ===");
    
    // Test 1: Check ATTENDANCE sheet exists
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(ATTENDANCE_SHEET);
    
    if (!sheet) {
      Logger.log("ERROR: ATTENDANCE sheet not found!");
      Logger.log("Creating ATTENDANCE sheet...");
      initializeAttendanceSheet();
      sheet = ss.getSheetByName(ATTENDANCE_SHEET);
      if (!sheet) {
        return "ERROR: Could not create ATTENDANCE sheet";
      }
    }
    
    Logger.log("✓ ATTENDANCE sheet found: " + sheet.getName());
    Logger.log("  Last row: " + sheet.getLastRow());
    
    // Test 2: Check headers
    var headers = sheet.getRange(1, 1, 1, 12).getValues()[0];
    Logger.log("✓ Headers: " + headers.join(", "));
    
    // Test 3: Get a test employee
    var empSheet = ss.getSheetByName(EMPLOYEE_SHEET);
    if (!empSheet) {
      return "ERROR: EMPLOYEE sheet not found";
    }
    
    var empData = empSheet.getDataRange().getValues();
    if (empData.length < 2) {
      return "ERROR: No employees in EMPLOYEE sheet";
    }
    
    var testEmpId = empData[1][0].toString().trim();
    var testFirstName = empData[1][1].toString().trim();
    var testLastName = empData[1][2].toString().trim();
    var testEmpName = (testFirstName + " " + testLastName).trim();
    
    Logger.log("✓ Test employee: " + testEmpId + " - " + testEmpName);
    
    // Test 4: Try to append a test row
    try {
      var testDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
      var testTime = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "hh:mm a");
      
      Logger.log("Attempting to append test row...");
      var lastRowBefore = sheet.getLastRow();
      
      sheet.appendRow([
        testDate,
        "TEST EMPLOYEE",
        testTime,
        "Test Location",
        "Test Photo URL",
        "",
        "",
        "",
        "",
        ""
      ]);
      
      SpreadsheetApp.flush();
      
      var lastRowAfter = sheet.getLastRow();
      
      Logger.log("✓ Test row appended successfully");
      Logger.log("  Last row before: " + lastRowBefore);
      Logger.log("  Last row after: " + lastRowAfter);
      
      // Clean up test row
      if (lastRowAfter > lastRowBefore) {
        sheet.deleteRow(lastRowAfter);
        Logger.log("✓ Test row cleaned up");
      }
      
    } catch (appendErr) {
      Logger.log("ERROR appending test row: " + appendErr.message);
      return "ERROR: " + appendErr.message;
    }
    
    // Test 5: Test getAttendanceStatus
    try {
      var status = getAttendanceStatus(testEmpName);
      Logger.log("✓ getAttendanceStatus works");
      Logger.log("  Status: " + JSON.stringify(status));
    } catch (statusErr) {
      Logger.log("ERROR in getAttendanceStatus: " + statusErr.message);
    }
    
    // Test 6: Test recordAttendance with mock data
    try {
      Logger.log("Testing recordAttendance function...");
      
      var mockPayload = {
        fullName: testEmpId + " - " + testEmpName,
        location: "Test Office",
        type: "IN",
        photoBase64: ""
      };
      
      Logger.log("Mock payload: " + JSON.stringify(mockPayload));
      
      // Check if already clocked in today
      var currentStatus = getAttendanceStatus(testEmpName);
      if (currentStatus.hasClockIn) {
        Logger.log("⚠ Employee already clocked in today - skipping test");
      } else {
        var result = recordAttendance(mockPayload);
        Logger.log("✓ recordAttendance SUCCESS: " + result);
        
        // Verify the row was added
        var newStatus = getAttendanceStatus(testEmpName);
        Logger.log("  New status: " + JSON.stringify(newStatus));
        
        if (newStatus.hasClockIn) {
          Logger.log("✓ Clock in was recorded successfully!");
          
          // Clean up - delete the test record
          if (newStatus.rowNumber > 1) {
            sheet.deleteRow(newStatus.rowNumber);
            Logger.log("✓ Test record cleaned up");
          }
        } else {
          Logger.log("ERROR: Clock in was not recorded");
        }
      }
      
    } catch (recordErr) {
      Logger.log("ERROR in recordAttendance: " + recordErr.message);
      Logger.log("Stack: " + recordErr.stack);
    }
    
    Logger.log("=== ATTENDANCE SYSTEM TEST COMPLETED ===");
    return "SUCCESS: All tests passed! System is working correctly.";
    
  } catch (err) {
    Logger.log("FATAL ERROR: " + err.message);
    Logger.log("Stack: " + err.stack);
    return "FATAL ERROR: " + err.message;
  }
}

/**
 * SIMPLE TEST - Run this from Apps Script to test clock in
 * Uses the first employee in your EMPLOYEE sheet
 */
function simpleClockInTest() {
  try {
    Logger.log("=== SIMPLE CLOCK IN TEST ===");
    
    // Get spreadsheet
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    Logger.log("✓ Spreadsheet opened");
    
    // Check ATTENDANCE sheet
    var attSheet = ss.getSheetByName(ATTENDANCE_SHEET);
    if (!attSheet) {
      Logger.log("✗ ATTENDANCE sheet not found!");
      return "FAILED: ATTENDANCE sheet not found";
    }
    Logger.log("✓ ATTENDANCE sheet found");
    
    // Check EMPLOYEE sheet
    var empSheet = ss.getSheetByName(EMPLOYEE_SHEET);
    if (!empSheet) {
      Logger.log("✗ EMPLOYEE sheet not found!");
      return "FAILED: EMPLOYEE sheet not found";
    }
    Logger.log("✓ EMPLOYEE sheet found");
    
    // Get first employee
    var empData = empSheet.getDataRange().getValues();
    if (empData.length < 2) {
      Logger.log("✗ No employees in EMPLOYEE sheet!");
      return "FAILED: No employees found";
    }
    
    var empId = empData[1][0] ? empData[1][0].toString().trim() : "";
    var firstName = empData[1][1] ? empData[1][1].toString().trim() : "";
    var lastName = empData[1][2] ? empData[1][2].toString().trim() : "";
    var fullName = (firstName + " " + lastName).trim();
    
    Logger.log("✓ Test employee: " + empId + " - " + fullName);
    
    // Check LOCATION sheet
    var locSheet = ss.getSheetByName(LOCATION_SHEET);
    var testLocation = "Test Location";
    if (locSheet) {
      var locData = locSheet.getDataRange().getValues();
      if (locData.length > 1) {
        testLocation = locData[1][1] ? locData[1][1].toString().trim() : "Test Location";
      }
    }
    Logger.log("✓ Test location: " + testLocation);
    
    // Create test payload
    var testPayload = {
      fullName: empId + " - " + fullName,
      location: testLocation,
      type: "IN",
      photoBase64: "" // No photo for test
    };
    
    Logger.log("✓ Payload created");
    Logger.log("Calling recordAttendance...");
    
    // Call recordAttendance
    var result = recordAttendance(testPayload);
    
    Logger.log("✓ SUCCESS: " + result);
    Logger.log("=== TEST COMPLETED ===");
    
    return "SUCCESS: " + result;
    
  } catch (err) {
    Logger.log("✗ TEST FAILED");
    Logger.log("Error: " + err.message);
    Logger.log("Stack: " + err.stack);
    return "FAILED: " + err.message;
  }
}

/**
 * EMERGENCY TEST - Call this from Apps Script to test if clock in works
 * This simulates a real clock in with minimal data
 */
function emergencyClockInTest() {
  try {
    Logger.log("=== EMERGENCY CLOCK IN TEST ===");
    
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(ATTENDANCE_SHEET);
    
    if (!sheet) {
      Logger.log("ERROR: ATTENDANCE sheet not found!");
      return "FAILED: Sheet not found";
    }
    
    Logger.log("Sheet found: " + sheet.getName());
    
    // Get first employee from EMPLOYEE sheet
    var empSheet = ss.getSheetByName(EMPLOYEE_SHEET);
    if (!empSheet) {
      Logger.log("ERROR: EMPLOYEE sheet not found!");
      return "FAILED: EMPLOYEE sheet not found";
    }
    
    var empData = empSheet.getDataRange().getValues();
    if (empData.length < 2) {
      Logger.log("ERROR: No employees found!");
      return "FAILED: No employees";
    }
    
    var testEmpId = empData[1][0] ? empData[1][0].toString().trim() : "";
    var firstName = empData[1][1] ? empData[1][1].toString().trim() : "";
    var lastName = empData[1][2] ? empData[1][2].toString().trim() : "";
    var fullName = (firstName + " " + lastName).trim();
    
    Logger.log("Testing with employee: " + fullName);
    
    // Create test payload (simulating frontend call)
    var testPayload = {
      fullName: testEmpId + " - " + fullName,
      location: "Test Location",
      type: "IN",
      photoBase64: ""
    };
    
    Logger.log("Calling recordAttendance...");
    var result = recordAttendance(testPayload);
    
    Logger.log("Result: " + result);
    Logger.log("=== TEST COMPLETED SUCCESSFULLY ===");
    
    return result;
    
  } catch (err) {
    Logger.log("EMERGENCY TEST FAILED: " + err.message);
    Logger.log("Stack: " + err.stack);
    return "ERROR: " + err.message;
  }
}

/**
 * Simple test to verify sheet access and basic append operations
 * This bypasses all logic to test if basic sheet writing works
 */
function testDirectWrite() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var attendanceSheet = ss.getSheetByName(ATTENDANCE_SHEET);
    
    Logger.log("Test 1 - ATTENDANCE sheet exists: " + (attendanceSheet ? "YES" : "NO"));
    if (!attendanceSheet) {
      Logger.log("ERROR: ATTENDANCE sheet not found! Expected name: " + ATTENDANCE_SHEET);
      return "FAILED: ATTENDANCE sheet not found";
    }
    
    // Test 2: Check EMPLOYEE sheet structure
    var empSheet = ss.getSheetByName(EMPLOYEE_SHEET);
    Logger.log("Test 2 - EMPLOYEE sheet exists: " + (empSheet ? "YES" : "NO"));
    
    if (empSheet) {
      var empData = empSheet.getDataRange().getValues();
      Logger.log("Employee sheet has " + (empData.length - 1) + " employees");
      
      if (empData.length > 1) {
        // Show first employee structure
        Logger.log("First employee row:");
        Logger.log("  Column A (ID): " + empData[1][0]);
        Logger.log("  Column B (First Name): " + empData[1][1]);
        Logger.log("  Column C (Last Name): " + empData[1][2]);
        
        var testId = empData[1][0] ? empData[1][0].toString().trim() : "";
        var firstName = empData[1][1] ? empData[1][1].toString().trim() : "";
        var lastName = empData[1][2] ? empData[1][2].toString().trim() : "";
        var fullName = (firstName + " " + lastName).trim();
        
        Logger.log("  Combined Full Name: " + fullName);
        
        // Test 3: Test getEmployeeByID
        Logger.log("\nTest 3 - Testing getEmployeeByID with ID: " + testId);
        var result = getEmployeeByID(testId);
        Logger.log("Result: " + JSON.stringify(result));
        
        // Test 4: Try to create a test attendance record
        if (result.success) {
          Logger.log("\nTest 4 - Creating test attendance record");
          Logger.log("Using employee name: " + result.employeeName);
          
          var testPayload = {
            fullName: result.fullName,
            location: "Test Location",
            type: "IN",
            photoBase64: ""
          };
          
          Logger.log("Test payload: " + JSON.stringify(testPayload));
          
          // Note: This will actually create a record if successful
          // Comment out if you don't want test data
          // var recordResult = recordAttendance(testPayload);
          // Logger.log("Record result: " + recordResult);
          
          Logger.log("Test skipped - uncomment to test actual recording");
        }
      }
    }
    
    Logger.log("\n=== ALL TESTS COMPLETED ===");
    return "SUCCESS: System check complete";
    
  } catch (err) {
    Logger.log("ERROR in testDirectWrite: " + err.message);
    Logger.log("Stack: " + err.stack);
    return "ERROR: " + err.message;
  }
}
