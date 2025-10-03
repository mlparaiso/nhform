/**
 * QA NH Form Data Storage Manager - Simplified & Consolidated
 * Automated sheet and folder creation, backup management, and data archiving
 * 
 * @author: Enhanced QA System
 * @version: 2.0 - Simplified
 * @timezone: America/New_York (EST)
 */

// ===================================================================
// === CONFIGURATION ===
// ===================================================================

const DATA_STORAGE_CONFIG = {
  // Main spreadsheet and folder IDs
  MAIN_SPREADSHEET_ID: "1InCv7-d8-KhovEGb8tdI3TxLJx3PTou5ZJKsx1nZhwI",
  BACKUP_FOLDER_ID: "1AtlMVddT2MbMgSD5LeTrRkIPgZle8E2r",
  
  // Timezone settings
  TIMEZONE: "America/New_York",
  
  // Sheet names
  SHEET_NAMES: {
    MAIN_RESPONSES: "QA NH Form Responses",
    QUALITY_LOG: "QA NH DataQuality Log",
    SYSTEM_AUDIT: "QA NH System Audit",
    CONFIG_SETTINGS: "QA NH Config Settings",
    MIGRATION_LOG: "QA NH Migration Log",
    ROSTER: "roster" // Keep existing
  },
  
  // Specific Drive folder IDs
  FOLDER_IDS: {
    ACTIVE_DATA: "1KYxFFz_yXMBsHpgcVdg-_0H8KdfqAcxQ",
    DAILY_BACKUPS: "1wBv1QxCN6K7I0EyeoyLne0k-C4QoEp1E",
    WEEKLY_BACKUPS: "1iHqr4BjuGbuj0yy56OmhCXgrMfA-EvVC",
    MONTHLY_ARCHIVES: "17m2ZMRSuTFQit7_zaAzJLwdvo4-cQqqu",
    QUALITY_REPORTS: "1JgtSJEsxJSgdQCZzSCLkwsmm0SanIQYL",
    SYSTEM_LOGS: "19oCwZmwLN_hwmFsuNKAyAHSOoznbknvs",
    CONFIGURATION: "1JSY6A0n3flSFFNdLfGGtfagvPMWUGeHv"
  },
  
  // Retention policies and schedule
  RETENTION_POLICIES: {
    DAILY_BACKUPS: 30,    // Keep 30 days
    WEEKLY_BACKUPS: 12,   // Keep 12 weeks
    MONTHLY_ARCHIVES: 24, // Keep 24 months
    ARCHIVE_AFTER_MONTHS: 3
  },
  
  SCHEDULE: {
    DAILY_BACKUP_HOUR: 2,     // 2 AM EST
    WEEKLY_MAINTENANCE_HOUR: 3, // 3 AM EST on Sundays
    MONTHLY_ARCHIVE_HOUR: 4    // 4 AM EST on 1st of month
  }
};

// ===================================================================
// === CORE FUNCTIONS ===
// ===================================================================

/**
 * Main function to initialize the complete QA NH Data Storage system
 */
function initializeQANHDataStorage() {
  try {
    console.log("🚀 Starting QA NH Data Storage initialization...");
    const startTime = new Date();
    
    // Step 1: Create spreadsheet structure
    console.log("📊 Setting up spreadsheet structure...");
    const sheetsResult = setupSpreadsheetStructure();
    
    // Step 2: Set up automated schedules
    console.log("⏰ Setting up automated schedules...");
    setupAutomatedSchedules();
    
    // Step 3: Perform initial health check
    console.log("🔍 Performing system health check...");
    const healthReport = checkSystemHealth();
    
    const duration = new Date() - startTime;
    console.log(`✅ QA NH Data Storage system initialized successfully! (${duration}ms)`);
    
    logEvent("INITIALIZATION", "System initialized successfully", "SUCCESS", { duration: `${duration}ms`, sheets: sheetsResult });
    
    return {
      success: true,
      message: "QA NH Data Storage system initialized successfully!",
      duration: `${duration}ms`,
      sheetsCreated: sheetsResult,
      healthReport: healthReport
    };
    
  } catch (error) {
    console.error("❌ Error during initialization:", error.message);
    logEvent("INITIALIZATION_ERROR", error.message, "FAILED");
    
    return {
      success: false,
      message: `Initialization failed: ${error.message}`,
      error: error.toString()
    };
  }
}

/**
 * Unified system health check
 */
function checkSystemHealth(components = ['sheets', 'folders', 'triggers']) {
  try {
    const healthReport = {
      timestamp: getESTTimestamp(),
      overallHealth: "HEALTHY",
      issues: []
    };
    
    // Check sheets
    if (components.includes('sheets')) {
      const sheetsStatus = checkSheetsExist();
      healthReport.sheetsStatus = sheetsStatus;
      if (!sheetsStatus.allExist) {
        healthReport.issues.push("Missing sheets");
      }
    }
    
    // Check folders
    if (components.includes('folders')) {
      const foldersStatus = checkFoldersExist();
      healthReport.foldersStatus = foldersStatus;
      if (!foldersStatus.allExist) {
        healthReport.issues.push("Missing folders");
      }
    }
    
    // Check triggers
    if (components.includes('triggers')) {
      const triggerStatus = checkTriggers();
      healthReport.triggerStatus = triggerStatus;
      if (triggerStatus.missing > 0) {
        healthReport.issues.push("Missing triggers");
      }
    }
    
    // Determine overall health
    if (healthReport.issues.length > 0) {
      healthReport.overallHealth = "ISSUES_DETECTED";
    }
    
    logEvent("HEALTH_CHECK", JSON.stringify(healthReport), healthReport.overallHealth === "HEALTHY" ? "SUCCESS" : "WARNING");
    
    return healthReport;
    
  } catch (error) {
    console.error("❌ Error during health check:", error.message);
    return { overallHealth: "ERROR", error: error.message };
  }
}

// ===================================================================
// === BACKUP & ARCHIVING ===
// ===================================================================

/**
 * Unified backup function for daily, weekly, and monthly operations
 */
function performBackup(backupType = 'daily') {
  try {
    console.log(`💾 Starting ${backupType} backup...`);
    const startTime = new Date();
    
    // Get backup configuration
    const config = getBackupConfig(backupType);
    
    // Get the backup folder
    const backupFolder = getFolderManager(config.folderType);
    if (!backupFolder) {
      throw new Error(`${config.folderType} folder not found`);
    }
    
    // Create backup file
    const backupName = generateBackupName(backupType);
    const originalFile = DriveApp.getFileById(DATA_STORAGE_CONFIG.MAIN_SPREADSHEET_ID);
    const copiedFile = originalFile.makeCopy(backupName);
    
    // Move to backup folder
    backupFolder.addFile(copiedFile);
    DriveApp.getRootFolder().removeFile(copiedFile);
    
    // Clean old backups
    const cleanedCount = cleanupOldFiles(backupFolder, config.retentionDays);
    
    const duration = new Date() - startTime;
    console.log(`✅ ${backupType} backup completed: ${backupName} (${duration}ms)`);
    
    logEvent(`${backupType.toUpperCase()}_BACKUP`, `Backup created: ${backupName}`, "SUCCESS", { 
      duration: `${duration}ms`, 
      cleanedFiles: cleanedCount 
    });
    
    return { 
      success: true, 
      backupName: backupName, 
      duration: `${duration}ms`,
      cleanedFiles: cleanedCount
    };
    
  } catch (error) {
    console.error(`❌ ${backupType} backup failed:`, error.message);
    logEvent(`${backupType.toUpperCase()}_BACKUP_ERROR`, error.message, "FAILED");
    throw error;
  }
}

/**
 * Daily backup function (called by trigger)
 */
function performDailyBackup() {
  return performBackup('daily');
}

/**
 * Weekly maintenance function (called by trigger)
 */
function performWeeklyMaintenance() {
  try {
    const backupResult = performBackup('weekly');
    const healthReport = checkSystemHealth();
    
    return { 
      success: true, 
      backup: backupResult, 
      healthReport: healthReport 
    };
  } catch (error) {
    logEvent("WEEKLY_MAINTENANCE_ERROR", error.message, "FAILED");
    throw error;
  }
}

/**
 * Monthly archiving function (called by trigger)
 */
function performMonthlyArchiving() {
  try {
    const backupResult = performBackup('monthly');
    // Add archiving logic here if needed
    
    return { 
      success: true, 
      backup: backupResult 
    };
  } catch (error) {
    logEvent("MONTHLY_ARCHIVING_ERROR", error.message, "FAILED");
    throw error;
  }
}

/**
 * Get backup configuration based on type
 */
function getBackupConfig(backupType) {
  const configs = {
    daily: {
      folderType: "Daily_Backups",
      retentionDays: DATA_STORAGE_CONFIG.RETENTION_POLICIES.DAILY_BACKUPS,
      prefix: "QA_NH_Backup"
    },
    weekly: {
      folderType: "Weekly_Backups", 
      retentionDays: DATA_STORAGE_CONFIG.RETENTION_POLICIES.WEEKLY_BACKUPS * 7,
      prefix: "QA_NH_Weekly"
    },
    monthly: {
      folderType: "Monthly_Archives",
      retentionDays: DATA_STORAGE_CONFIG.RETENTION_POLICIES.MONTHLY_ARCHIVES * 30,
      prefix: "QA_NH_Archive"
    }
  };
  
  return configs[backupType] || configs.daily;
}

/**
 * Generate backup name based on type
 */
function generateBackupName(backupType) {
  const config = getBackupConfig(backupType);
  const now = getESTTimestamp();
  
  switch (backupType) {
    case 'weekly':
      const weekNumber = Utilities.formatDate(now, DATA_STORAGE_CONFIG.TIMEZONE, 'w');
      const year = Utilities.formatDate(now, DATA_STORAGE_CONFIG.TIMEZONE, 'yyyy');
      return `${config.prefix}_${year}-W${weekNumber.padStart(2, '0')}`;
    
    case 'monthly':
      const monthYear = Utilities.formatDate(now, DATA_STORAGE_CONFIG.TIMEZONE, 'yyyy-MM');
      return `${config.prefix}_${monthYear}`;
    
    default: // daily
      const timestamp = Utilities.formatDate(now, DATA_STORAGE_CONFIG.TIMEZONE, 'yyyy-MM-dd_HH-mm');
      return `${config.prefix}_${timestamp}`;
  }
}

/**
 * Clean up old backup files
 */
function cleanupOldFiles(folder, retentionDays) {
  try {
    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - retentionDays);
    
    const files = folder.getFiles();
    let deletedCount = 0;
    
    while (files.hasNext()) {
      const file = files.next();
      if (file.getDateCreated() < cutoffDate) {
        file.setTrashed(true);
        deletedCount++;
      }
    }
    
    if (deletedCount > 0) {
      console.log(`🗑️ Cleaned ${deletedCount} old backup files`);
    }
    
    return deletedCount;
    
  } catch (error) {
    console.error("❌ Error cleaning old backups:", error.message);
    return 0;
  }
}

// ===================================================================
// === SHEET & FOLDER MANAGEMENT ===
// ===================================================================

/**
 * Unified folder manager
 */
function getFolderManager(folderPath) {
  const folderMapping = {
    "Active_Data": DATA_STORAGE_CONFIG.FOLDER_IDS.ACTIVE_DATA,
    "Daily_Backups": DATA_STORAGE_CONFIG.FOLDER_IDS.DAILY_BACKUPS,
    "Weekly_Backups": DATA_STORAGE_CONFIG.FOLDER_IDS.WEEKLY_BACKUPS,
    "Monthly_Archives": DATA_STORAGE_CONFIG.FOLDER_IDS.MONTHLY_ARCHIVES,
    "Quality_Reports": DATA_STORAGE_CONFIG.FOLDER_IDS.QUALITY_REPORTS,
    "System_Logs": DATA_STORAGE_CONFIG.FOLDER_IDS.SYSTEM_LOGS,
    "Configuration": DATA_STORAGE_CONFIG.FOLDER_IDS.CONFIGURATION
  };
  
  const folderId = folderMapping[folderPath];
  if (!folderId) {
    console.error(`❌ Unknown folder path: ${folderPath}`);
    return null;
  }
  
  try {
    return DriveApp.getFolderById(folderId);
  } catch (error) {
    console.error(`❌ Error accessing folder ${folderPath} (${folderId}):`, error.message);
    return null;
  }
}

/**
 * Check if folders exist
 */
function checkFoldersExist() {
  try {
    // Use the correct folder path names that match getFolderManager mapping
    const requiredFolders = [
      "Active_Data", "Daily_Backups", "Weekly_Backups", "Monthly_Archives",
      "Quality_Reports", "System_Logs", "Configuration"
    ];
    const existingFolders = [];
    const missingFolders = [];
    
    requiredFolders.forEach(folderName => {
      try {
        const folder = getFolderManager(folderName);
        if (folder) {
          existingFolders.push(folderName);
        } else {
          missingFolders.push(folderName);
        }
      } catch (folderError) {
        missingFolders.push(folderName);
      }
    });
    
    return {
      allExist: missingFolders.length === 0,
      required: requiredFolders.length,
      existing: existingFolders.length,
      missing: missingFolders
    };
    
  } catch (error) {
    return { allExist: false, error: error.message };
  }
}

/**
 * Unified spreadsheet structure setup
 */
function setupSpreadsheetStructure() {
  try {
    const ss = SpreadsheetApp.openById(DATA_STORAGE_CONFIG.MAIN_SPREADSHEET_ID);
    console.log(`📊 Working with spreadsheet: ${ss.getName()}`);
    
    const createdSheets = [];
    const sheetConfigs = getSheetConfigurations();
    
    Object.entries(sheetConfigs).forEach(([sheetKey, config]) => {
      const sheetName = DATA_STORAGE_CONFIG.SHEET_NAMES[sheetKey];
      if (createSheet(ss, sheetName, config)) {
        createdSheets.push(sheetName);
      }
    });
    
    console.log(`✅ Created ${createdSheets.length} new sheets`);
    logEvent("SHEET_CREATION", `Created sheets: ${createdSheets.join(', ')}`, "SUCCESS");
    
    return createdSheets;
    
  } catch (error) {
    console.error("❌ Error creating spreadsheet structure:", error.message);
    logEvent("SHEET_CREATION_ERROR", error.message, "FAILED");
    throw error;
  }
}

/**
 * Get sheet configurations
 */
function getSheetConfigurations() {
  return {
    MAIN_RESPONSES: {
      headers: [
        "RecordID", "Timestamp", "FormVersion", "SubmittedBy",
        "Sap ID", "Team Member", "Team Leader", "Operations Manager", "Line of Business",
        "Interaction Date", "Customer's BAN", "FLP Conversation ID", "Listening Type", "Call Duration",
        "Willingness to Help", "Warm and Energetic Greeting that Connects with the Customer", "Listen & Retain",
        "Secure the Customer's Identity Based on Department Guidelines", "Probe Using a Combination of Open-ended and Closed-ended Questions", 
        "Investigate & Check All Applicable Resources", "Offer Solutions / Explanations / Options", 
        "Begin Processing Transaction, Adding Notes and Updating Account Information when Appropriate",
        "Confirm What You Know and Paraphrase Customer", "Check for Understanding", 
        "Lead with a Benefit Prior to Asking Questions and Securing the Account", "Summarize Solution", 
        "End on a High Note with a Personalized Close", "Reinforce the Extras You Did",
        "Early Bridge to Value When Appropriate", "Later Bridge to Value when Appropriate / Value Proposition", 
        "Lead Customer to Accept an Option or Solution", "Check for Satisfaction",
        "Overcome Objections if Necessary", "Acknowledge (Empathy When Needed)",
        "Learn from Your Customer", "Handling Early Objections", "Customer Intimacy", "Value Proposition", 
        "What type of bridging was provided?", "Cancellation Root Cause", "Was there an unethical credit?",
        "Right size", "Use a tiered approach", "Sufficient probing (+3 Questions)", 
        "Speak to value (WHY TELUS, reality check and service superiority)", "Loyalty statements implementation", 
        "Investigation (Review bill, compare competitor offer, search notes on the account, take a look at usage)", 
        "Was the agent able to retain the customer?", "Call Driver Level 1", "Call Driver Level 2", 
        "Call Summary", "Highlights", "Lowlights/Recommendation", "DataChecksum", "FieldCount", "ValidationStatus", "DataQualityScore"
      ],
      headerColor: "#4285f4"
    },
    QUALITY_LOG: {
      headers: ["Timestamp", "RecordID", "QualityScore", "ValidationStatus", "MissingFields", "FieldCount", "DataChecksum", "Notes", "ReviewedBy"],
      headerColor: "#34a853"
    },
    SYSTEM_AUDIT: {
      headers: ["Timestamp", "Action", "User", "RecordID", "Details", "Status", "Duration", "IPAddress"],
      headerColor: "#ea4335"
    },
    CONFIG_SETTINGS: {
      headers: ["Setting", "Value", "Description", "LastModified", "ModifiedBy"],
      headerColor: "#ff9800"
    },
    MIGRATION_LOG: {
      headers: ["Timestamp", "MigrationType", "SourceSheet", "TargetSheet", "RecordsProcessed", "Status", "Details", "Duration"],
      headerColor: "#9c27b0"
    }
  };
}

/**
 * Create a sheet with configuration
 */
function createSheet(spreadsheet, sheetName, config) {
  try {
    let sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
      
      // Set up headers
      sheet.appendRow(config.headers);
      sheet.getRange("1:1").setFontWeight("bold");
      sheet.getRange("1:1").setBackground(config.headerColor);
      sheet.getRange("1:1").setFontColor("white");
      sheet.setFrozenRows(1);
      
      // Set column widths for main responses sheet
      if (sheetName === DATA_STORAGE_CONFIG.SHEET_NAMES.MAIN_RESPONSES) {
        sheet.setColumnWidth(1, 150); // RecordID
        sheet.setColumnWidth(2, 120); // Timestamp
        sheet.setColumnWidth(5, 100); // Sap ID
        sheet.setColumnWidth(6, 150); // Team Member
      }
      
      console.log(`✅ Created and configured sheet: ${sheetName}`);
      return true;
    } else {
      console.log(`📊 Sheet already exists: ${sheetName}`);
      return false;
    }
  } catch (error) {
    console.error(`❌ Error creating sheet ${sheetName}:`, error.message);
    return false;
  }
}

/**
 * Check if sheets exist
 */
function checkSheetsExist() {
  try {
    const ss = SpreadsheetApp.openById(DATA_STORAGE_CONFIG.MAIN_SPREADSHEET_ID);
    const requiredSheets = Object.values(DATA_STORAGE_CONFIG.SHEET_NAMES);
    const existingSheets = ss.getSheets().map(sheet => sheet.getName());
    
    const missingSheets = requiredSheets.filter(sheetName => !existingSheets.includes(sheetName));
    
    return {
      allExist: missingSheets.length === 0,
      required: requiredSheets.length,
      existing: existingSheets.length,
      missing: missingSheets
    };
    
  } catch (error) {
    return { allExist: false, error: error.message };
  }
}

// ===================================================================
// === UTILITIES ===
// ===================================================================

/**
 * Gets EST timestamp
 */
function getESTTimestamp() {
  return new Date(new Date().toLocaleString('en-US', { timeZone: DATA_STORAGE_CONFIG.TIMEZONE }));
}

/**
 * Generates unique Record ID
 */
function generateRecordID() {
  const now = getESTTimestamp();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const day = String(now.getDate()).padStart(2, '0');
  const year = now.getFullYear();
  const dateStr = `${month}${day}${year}`;
  
  const alphabet = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz';
  let nanoid = '';
  for (let i = 0; i < 8; i++) {
    nanoid += alphabet.charAt(Math.floor(Math.random() * alphabet.length));
  }
  
  return `QA-${dateStr}-${nanoid}`;
}

/**
 * Check scheduled triggers
 */
function checkTriggers() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    const requiredTriggers = ['performDailyBackup', 'performWeeklyMaintenance', 'performMonthlyArchiving'];
    const existingTriggers = triggers.map(trigger => trigger.getHandlerFunction());
    
    const missingTriggers = requiredTriggers.filter(triggerName => !existingTriggers.includes(triggerName));
    
    return {
      total: triggers.length,
      required: requiredTriggers.length,
      missing: missingTriggers.length,
      missingTriggers: missingTriggers
    };
    
  } catch (error) {
    return { missing: -1, error: error.message };
  }
}

// ===================================================================
// === LOGGING & MONITORING ===
// ===================================================================

/**
 * Unified logging function
 */
function logEvent(action, details, status, metadata = {}) {
  try {
    const ss = SpreadsheetApp.openById(DATA_STORAGE_CONFIG.MAIN_SPREADSHEET_ID);
    const auditSheet = ss.getSheetByName(DATA_STORAGE_CONFIG.SHEET_NAMES.SYSTEM_AUDIT);
    
    if (auditSheet) {
      let userEmail = "System";
      try {
        userEmail = Session.getActiveUser().getEmail();
      } catch (emailError) {
        userEmail = "System";
      }
      
      const logEntry = [
        getESTTimestamp(),
        action,
        userEmail,
        metadata.recordId || "SYSTEM",
        typeof details === 'object' ? JSON.stringify(details) : details,
        status,
        metadata.duration || "0ms",
        "System"
      ];
      
      auditSheet.appendRow(logEntry);
    }
  } catch (error) {
    console.error("❌ Error logging event:", error.message);
  }
}

// ===================================================================
// === AUTOMATION & TRIGGERS ===
// ===================================================================

/**
 * Sets up all automated triggers
 */
function setupAutomatedSchedules() {
  try {
    console.log("⏰ Setting up automated schedules...");
    
    // Clear existing triggers first
    clearExistingTriggers();
    
    // Daily backup at 2 AM EST
    ScriptApp.newTrigger('performDailyBackup')
      .timeBased()
      .everyDays(1)
      .atHour(DATA_STORAGE_CONFIG.SCHEDULE.DAILY_BACKUP_HOUR)
      .create();
    console.log("✅ Daily backup trigger created (2 AM EST)");
    
    // Weekly maintenance on Sundays at 3 AM EST
    ScriptApp.newTrigger('performWeeklyMaintenance')
      .timeBased()
      .onWeekDay(ScriptApp.WeekDay.SUNDAY)
      .atHour(DATA_STORAGE_CONFIG.SCHEDULE.WEEKLY_MAINTENANCE_HOUR)
      .create();
    console.log("✅ Weekly maintenance trigger created (Sunday 3 AM EST)");
    
    // Monthly archiving on 1st of month at 4 AM EST
    ScriptApp.newTrigger('performMonthlyArchiving')
      .timeBased()
      .onMonthDay(1)
      .atHour(DATA_STORAGE_CONFIG.SCHEDULE.MONTHLY_ARCHIVE_HOUR)
      .create();
    console.log("✅ Monthly archiving trigger created (1st of month 4 AM EST)");
    
    logEvent("TRIGGERS_SETUP", "All automated triggers created successfully", "SUCCESS");
    
  } catch (error) {
    console.error("❌ Error setting up automated schedules:", error.message);
    logEvent("TRIGGERS_SETUP_ERROR", error.message, "FAILED");
    throw error;
  }
}

/**
 * Clears existing triggers to avoid duplicates
 */
function clearExistingTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  const targetFunctions = ['performDailyBackup', 'performWeeklyMaintenance', 'performMonthlyArchiving'];
  
  triggers.forEach(trigger => {
    if (targetFunctions.includes(trigger.getHandlerFunction())) {
      ScriptApp.deleteTrigger(trigger);
      console.log(`🗑️ Removed existing trigger: ${trigger.getHandlerFunction()}`);
    }
  });
}

// ===================================================================
// === TESTING (Optional) ===
// ===================================================================

/**
 * Test the simplified system
 */
function testSimplifiedSystem() {
  console.log("🧪 Testing simplified QA NH Data Storage system...");
  
  try {
    // Test Record ID generation
    const recordId = generateRecordID();
    console.log(`✅ Record ID generated: ${recordId}`);
    
    // Test EST timestamp
    const estTime = getESTTimestamp();
    console.log(`✅ EST timestamp: ${estTime}`);
    
    // Test folder access
    const dailyBackupFolder = getFolderManager("Daily_Backups");
    console.log(`✅ Daily backup folder accessible: ${dailyBackupFolder ? 'Yes' : 'No'}`);
    
    // Test system health
    const healthReport = checkSystemHealth();
    console.log(`✅ System health: ${healthReport.overallHealth}`);
    
    console.log("🎉 Simplified system test completed successfully!");
    
    return {
      success: true,
      recordId: recordId,
      estTime: estTime,
      folderAccess: !!dailyBackupFolder,
      healthReport: healthReport
    };
    
  } catch (error) {
    console.error("❌ Test failed:", error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Test backup functionality
 */
function testBackupFunction() {
  console.log("💾 Testing backup functionality...");
  
  try {
    const result = performBackup('daily');
    console.log("✅ Backup test completed successfully!");
    return result;
    
  } catch (error) {
    console.error("❌ Backup test failed:", error.message);
    return {
      success: false,
      error: error.message
    };
  }
}
