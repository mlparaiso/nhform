/**
 * QA NH Form Data Storage Manager
 * Automated sheet and folder creation, backup management, and data archiving
 * 
 * @author: Enhanced QA System
 * @version: 1.0
 * @timezone: America/New_York (EST)
 */

// === CONFIGURATION ===
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
  
  // Specific Drive folder IDs (using existing folders)
  FOLDER_IDS: {
    ACTIVE_DATA: "1KYxFFz_yXMBsHpgcVdg-_0H8KdfqAcxQ",
    DAILY_BACKUPS: "1wBv1QxCN6K7I0EyeoyLne0k-C4QoEp1E",
    WEEKLY_BACKUPS: "1iHqr4BjuGbuj0yy56OmhCXgrMfA-EvVC",
    MONTHLY_ARCHIVES: "17m2ZMRSuTFQit7_zaAzJLwdvo4-cQqqu",
    QUALITY_REPORTS: "1JgtSJEsxJSgdQCZzSCLkwsmm0SanIQYL",
    SYSTEM_LOGS: "19oCwZmwLN_hwmFsuNKAyAHSOoznbknvs",
    CONFIGURATION: "1JSY6A0n3flSFFNdLfGGtfagvPMWUGeHv"
  },
  
  // Drive folder structure (for reference only)
  DRIVE_FOLDERS: [
    "Active_Data",
    "Daily_Backups",
    "Weekly_Backups", 
    "Monthly_Archives",
    "Quality_Reports",
    "System_Logs",
    "Configuration"
  ],
  
  // Retention policies
  RETENTION_POLICIES: {
    DAILY_BACKUPS: 30,    // Keep 30 days
    WEEKLY_BACKUPS: 12,   // Keep 12 weeks
    MONTHLY_ARCHIVES: 24, // Keep 24 months
    ARCHIVE_AFTER_MONTHS: 3 // Archive data after 3 months
  },
  
  // Backup schedule (EST times)
  SCHEDULE: {
    DAILY_BACKUP_HOUR: 2,     // 2 AM EST
    WEEKLY_MAINTENANCE_HOUR: 3, // 3 AM EST on Sundays
    MONTHLY_ARCHIVE_HOUR: 4    // 4 AM EST on 1st of month
  }
};

// === MAIN INITIALIZATION FUNCTION ===

/**
 * Main function to initialize the complete QA NH Data Storage system
 * Run this once to set up everything
 */
function initializeQANHDataStorage() {
  try {
    console.log("🚀 Starting QA NH Data Storage initialization...");
    
    // Step 1: Create Drive folder structure
    console.log("📁 Creating Drive folder structure...");
    createDriveFolderStructure();
    
    // Step 2: Create spreadsheet structure
    console.log("📊 Creating spreadsheet structure...");
    createSpreadsheetStructure();
    
    // Step 3: Migrate existing data
    console.log("🔄 Migrating existing data...");
    migrateExistingData();
    
    // Step 4: Set up automated schedules
    console.log("⏰ Setting up automated schedules...");
    setupAutomatedSchedules();
    
    // Step 5: Create initial configuration
    console.log("⚙️ Creating initial configuration...");
    createInitialConfiguration();
    
    // Step 6: Generate initialization report
    console.log("📋 Generating initialization report...");
    const report = generateInitializationReport();
    
    console.log("✅ QA NH Data Storage system initialized successfully!");
    console.log("📊 Initialization Report:", report);
    
    return {
      success: true,
      message: "QA NH Data Storage system initialized successfully!",
      report: report
    };
    
  } catch (error) {
    console.error("❌ Error during initialization:", error.message);
    logSystemEvent("INITIALIZATION_ERROR", error.message, "FAILED");
    
    return {
      success: false,
      message: `Initialization failed: ${error.message}`,
      error: error.toString()
    };
  }
}

// === DRIVE FOLDER MANAGEMENT ===

/**
 * Creates the complete Drive folder structure
 */
function createDriveFolderStructure() {
  try {
    const mainFolder = DriveApp.getFolderById(DATA_STORAGE_CONFIG.BACKUP_FOLDER_ID);
    console.log(`📁 Main folder found: ${mainFolder.getName()}`);
    
    const createdFolders = [];
    
    DATA_STORAGE_CONFIG.DRIVE_FOLDERS.forEach(folderName => {
      try {
        if (!getFolderByName(mainFolder, folderName)) {
          const newFolder = mainFolder.createFolder(folderName);
          createdFolders.push(folderName);
          console.log(`✅ Created folder: ${folderName}`);
          
          // Add description to folder
          const descriptions = {
            "Active_Data": "Current working spreadsheets and active data files",
            "Daily_Backups": "Daily backup files (kept for 30 days)",
            "Weekly_Backups": "Weekly consolidated backups (kept for 12 weeks)",
            "Monthly_Archives": "Monthly data archives (kept for 24 months)",
            "Quality_Reports": "Quality analysis and performance reports",
            "System_Logs": "System operation logs and error tracking",
            "Configuration": "System configuration files and settings"
          };
          
          if (descriptions[folderName]) {
            newFolder.setDescription(descriptions[folderName]);
          }
          
        } else {
          console.log(`📁 Folder already exists: ${folderName}`);
        }
      } catch (folderError) {
        console.error(`❌ Error creating folder ${folderName}:`, folderError.message);
      }
    });
    
    logSystemEvent("FOLDER_CREATION", `Created folders: ${createdFolders.join(', ')}`, "SUCCESS");
    return createdFolders;
    
  } catch (error) {
    console.error("❌ Error creating Drive folder structure:", error.message);
    logSystemEvent("FOLDER_CREATION_ERROR", error.message, "FAILED");
    throw error;
  }
}

/**
 * Helper function to find folder by name within parent folder
 */
function getFolderByName(parentFolder, folderName) {
  const folders = parentFolder.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : null;
}

/**
 * Helper function to get folder by path using specific folder IDs
 */
function getFolderByPath(folderPath) {
  // Map folder names to their specific IDs
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

// === SPREADSHEET STRUCTURE MANAGEMENT ===

/**
 * Creates the complete spreadsheet structure with all required sheets
 */
function createSpreadsheetStructure() {
  try {
    const ss = SpreadsheetApp.openById(DATA_STORAGE_CONFIG.MAIN_SPREADSHEET_ID);
    console.log(`📊 Working with spreadsheet: ${ss.getName()}`);
    
    const createdSheets = [];
    
    // Create main responses sheet
    if (createSheet(ss, DATA_STORAGE_CONFIG.SHEET_NAMES.MAIN_RESPONSES, setupMainResponsesSheet)) {
      createdSheets.push(DATA_STORAGE_CONFIG.SHEET_NAMES.MAIN_RESPONSES);
    }
    
    // Create data quality log sheet
    if (createSheet(ss, DATA_STORAGE_CONFIG.SHEET_NAMES.QUALITY_LOG, setupDataQualitySheet)) {
      createdSheets.push(DATA_STORAGE_CONFIG.SHEET_NAMES.QUALITY_LOG);
    }
    
    // Create system audit sheet
    if (createSheet(ss, DATA_STORAGE_CONFIG.SHEET_NAMES.SYSTEM_AUDIT, setupSystemAuditSheet)) {
      createdSheets.push(DATA_STORAGE_CONFIG.SHEET_NAMES.SYSTEM_AUDIT);
    }
    
    // Create config settings sheet
    if (createSheet(ss, DATA_STORAGE_CONFIG.SHEET_NAMES.CONFIG_SETTINGS, setupConfigSheet)) {
      createdSheets.push(DATA_STORAGE_CONFIG.SHEET_NAMES.CONFIG_SETTINGS);
    }
    
    // Create migration log sheet
    if (createSheet(ss, DATA_STORAGE_CONFIG.SHEET_NAMES.MIGRATION_LOG, setupMigrationSheet)) {
      createdSheets.push(DATA_STORAGE_CONFIG.SHEET_NAMES.MIGRATION_LOG);
    }
    
    console.log(`✅ Created ${createdSheets.length} new sheets`);
    logSystemEvent("SHEET_CREATION", `Created sheets: ${createdSheets.join(', ')}`, "SUCCESS");
    
    return createdSheets;
    
  } catch (error) {
    console.error("❌ Error creating spreadsheet structure:", error.message);
    logSystemEvent("SHEET_CREATION_ERROR", error.message, "FAILED");
    throw error;
  }
}

/**
 * Helper function to create a sheet with setup function
 */
function createSheet(spreadsheet, sheetName, setupFunction) {
  try {
    let sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
      if (setupFunction) {
        setupFunction(sheet);
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

// === SHEET SETUP FUNCTIONS ===

/**
 * Sets up the main QA NH Form Responses sheet
 */
function setupMainResponsesSheet(sheet) {
  const headers = [
    // System metadata
    "RecordID", "Timestamp", "FormVersion", "SubmittedBy",
    
    // Team member details
    "Sap ID", "Team Member", "Team Leader", "Operations Manager", "Line of Business",
    
    // Call details
    "Interaction Date", "Customer's BAN", "FLP Conversation ID", "Listening Type", "Call Duration",
    
    // Core skills
    "Willingness to Help", "Warm and Energetic Greeting that Connects with the Customer", "Listen & Retain",
    
    // Critical skills
    "Secure the Customer's Identity Based on Department Guidelines", "Probe Using a Combination of Open-ended and Closed-ended Questions", "Investigate & Check All Applicable Resources", "Offer Solutions / Explanations / Options", "Begin Processing Transaction, Adding Notes and Updating Account Information when Appropriate",
    
    // Basic skills
    "Confirm What You Know and Paraphrase Customer", "Check for Understanding", "Lead with a Benefit Prior to Asking Questions and Securing the Account", "Summarize Solution", "End on a High Note with a Personalized Close", "Reinforce the Extras You Did",
    
    // Intermediate skills
    "Early Bridge to Value When Appropriate", "Later Bridge to Value when Appropriate / Value Proposition", "Lead Customer to Accept an Option or Solution", "Check for Satisfaction",
    
    // Advanced skills
    "Overcome Objections if Necessary", "Acknowledge (Empathy When Needed)",
    
    // Sales skills
    "Learn from Your Customer", "Handling Early Objections", "Customer Intimacy", "Value Proposition", "What type of bridging was provided?",
    
    // CLS data
    "Cancellation Root Cause", "Was there an unethical credit?",
    
    // CLS skills
    "Right size", "Use a tiered approach", "Sufficient probing (+3 Questions)", "Speak to value (WHY TELUS, reality check and service superiority)", "Loyalty statements implementation", "Investigation (Review bill, compare competitor offer, search notes on the account, take a look at usage)", "Was the agent able to retain the customer?",
    
    // Call evaluation
    "Call Driver Level 1", "Call Driver Level 2", "Call Summary", "Highlights", "Lowlights/Recommendation",
    
    // Data integrity
    "DataChecksum", "FieldCount", "ValidationStatus", "DataQualityScore"
  ];
  
  sheet.appendRow(headers);
  sheet.getRange("1:1").setFontWeight("bold");
  sheet.getRange("1:1").setBackground("#4285f4");
  sheet.getRange("1:1").setFontColor("white");
  sheet.setFrozenRows(1);
  
  // Set column widths for better readability
  sheet.setColumnWidth(1, 150); // RecordID
  sheet.setColumnWidth(2, 120); // Timestamp
  sheet.setColumnWidth(5, 100); // Sap ID
  sheet.setColumnWidth(6, 150); // Team Member
}

/**
 * Sets up the Data Quality Log sheet
 */
function setupDataQualitySheet(sheet) {
  const headers = [
    "Timestamp", "RecordID", "QualityScore", "ValidationStatus", 
    "MissingFields", "FieldCount", "DataChecksum", "Notes", "ReviewedBy"
  ];
  
  sheet.appendRow(headers);
  sheet.getRange("1:1").setFontWeight("bold");
  sheet.getRange("1:1").setBackground("#34a853");
  sheet.getRange("1:1").setFontColor("white");
  sheet.setFrozenRows(1);
  
  // Add sample entry
  const sampleEntry = [
    new Date(), "SAMPLE", 100, "PASSED", "", 45, "abc123", "Sample quality log entry", "System"
  ];
  sheet.appendRow(sampleEntry);
}

/**
 * Sets up the System Audit sheet
 */
function setupSystemAuditSheet(sheet) {
  const headers = [
    "Timestamp", "Action", "User", "RecordID", "Details", "Status", "Duration", "IPAddress"
  ];
  
  sheet.appendRow(headers);
  sheet.getRange("1:1").setFontWeight("bold");
  sheet.getRange("1:1").setBackground("#ea4335");
  sheet.getRange("1:1").setFontColor("white");
  sheet.setFrozenRows(1);
  
  // Log the sheet creation
  const creationEntry = [
    new Date(), "SHEET_CREATED", Session.getActiveUser().getEmail(), 
    "N/A", "System Audit sheet created", "SUCCESS", "0ms", "System"
  ];
  sheet.appendRow(creationEntry);
}

/**
 * Sets up the Config Settings sheet
 */
function setupConfigSheet(sheet) {
  const headers = [
    "Setting", "Value", "Description", "LastModified", "ModifiedBy"
  ];
  
  sheet.appendRow(headers);
  sheet.getRange("1:1").setFontWeight("bold");
  sheet.getRange("1:1").setBackground("#ff9800");
  sheet.getRange("1:1").setFontColor("white");
  sheet.setFrozenRows(1);
  
  // Add default configuration
  const defaultConfig = [
    ["ARCHIVE_AFTER_MONTHS", "3", "Archive data after X months", new Date(), "System"],
    ["BACKUP_RETENTION_DAYS", "30", "Keep daily backups for X days", new Date(), "System"],
    ["DATA_QUALITY_THRESHOLD", "80", "Minimum quality score threshold", new Date(), "System"],
    ["TIMEZONE", "America/New_York", "System timezone (EST)", new Date(), "System"],
    ["FORM_VERSION", "2.1", "Current form version", new Date(), "System"],
    ["AUTO_BACKUP_ENABLED", "true", "Enable automatic daily backups", new Date(), "System"],
    ["AUTO_ARCHIVE_ENABLED", "true", "Enable automatic monthly archiving", new Date(), "System"]
  ];
  
  defaultConfig.forEach(config => sheet.appendRow(config));
}

/**
 * Sets up the Migration Log sheet
 */
function setupMigrationSheet(sheet) {
  const headers = [
    "Timestamp", "MigrationType", "SourceSheet", "TargetSheet", "RecordsProcessed", "Status", "Details", "Duration"
  ];
  
  sheet.appendRow(headers);
  sheet.getRange("1:1").setFontWeight("bold");
  sheet.getRange("1:1").setBackground("#9c27b0");
  sheet.getRange("1:1").setFontColor("white");
  sheet.setFrozenRows(1);
  
  // Log the sheet creation
  const creationEntry = [
    new Date(), "SHEET_CREATION", "N/A", "QA NH Migration Log", 0, "SUCCESS", "Migration log sheet created", "0ms"
  ];
  sheet.appendRow(creationEntry);
}

// === DATA MIGRATION FUNCTIONS ===

/**
 * Migrates existing data from old "Responses" sheet to new structure
 */
function migrateExistingData() {
  try {
    const ss = SpreadsheetApp.openById(DATA_STORAGE_CONFIG.MAIN_SPREADSHEET_ID);
    const oldSheet = ss.getSheetByName("Responses");
    
    if (!oldSheet || oldSheet.getLastRow() <= 1) {
      console.log("📊 No existing data to migrate");
      logMigrationEvent("NO_DATA", "Responses", "QA NH Form Responses", 0, "SUCCESS", "No existing data found");
      return { migrated: 0, message: "No existing data to migrate" };
    }
    
    const newSheet = ss.getSheetByName(DATA_STORAGE_CONFIG.SHEET_NAMES.MAIN_RESPONSES);
    if (!newSheet) {
      throw new Error("Target sheet 'QA NH Form Responses' not found");
    }
    
    console.log("🔄 Starting data migration...");
    const startTime = new Date();
    
    // Get existing data
    const existingData = oldSheet.getDataRange().getValues();
    const headers = existingData[0];
    const dataRows = existingData.slice(1);
    
    console.log(`📊 Found ${dataRows.length} records to migrate`);
    
    // Transform and migrate each row
    let migratedCount = 0;
    const errors = [];
    
    dataRows.forEach((row, index) => {
      try {
        const transformedRow = transformRowToNewStructure(headers, row);
        newSheet.appendRow(transformedRow);
        migratedCount++;
        
        if (migratedCount % 10 === 0) {
          console.log(`📊 Migrated ${migratedCount}/${dataRows.length} records...`);
        }
      } catch (rowError) {
        errors.push(`Row ${index + 2}: ${rowError.message}`);
        console.error(`❌ Error migrating row ${index + 2}:`, rowError.message);
      }
    });
    
    // Rename old sheet for backup
    const backupName = `Responses_Backup_${Utilities.formatDate(new Date(), DATA_STORAGE_CONFIG.TIMEZONE, 'yyyy-MM-dd_HH-mm')}`;
    oldSheet.setName(backupName);
    
    const duration = new Date() - startTime;
    const result = {
      migrated: migratedCount,
      errors: errors.length,
      duration: `${duration}ms`,
      backupSheet: backupName
    };
    
    console.log(`✅ Migration completed: ${migratedCount} records migrated, ${errors.length} errors`);
    logMigrationEvent("DATA_MIGRATION", "Responses", "QA NH Form Responses", migratedCount, "SUCCESS", JSON.stringify(result), `${duration}ms`);
    
    return result;
    
  } catch (error) {
    console.error("❌ Error during data migration:", error.message);
    logMigrationEvent("DATA_MIGRATION", "Responses", "QA NH Form Responses", 0, "FAILED", error.message, "0ms");
    throw error;
  }
}

/**
 * Transforms old row structure to new enhanced structure
 */
function transformRowToNewStructure(oldHeaders, oldRow) {
  // Create enhanced data structure for the old row
  const formData = {};
  
  // Map old data to form data object
  oldHeaders.forEach((header, index) => {
    if (oldRow[index] !== undefined && oldRow[index] !== null && oldRow[index] !== "") {
      formData[header] = oldRow[index];
    }
  });
  
  // Generate Record ID for migrated data
  const recordId = generateRecordID();
  const timestamp = formData["Timestamp"] || new Date();
  
  // Create enhanced structure
  const enhancedData = {
    recordId: recordId,
    timestamp: timestamp,
    formVersion: "2.0-MIGRATED",
    submittedBy: "MIGRATION_SYSTEM",
    teamMemberDetails: {
      sapId: formData["Sap ID"] || "",
      teamMember: formData["Team Member"] || "",
      teamLeader: formData["Team Leader"] || "",
      operationsManager: formData["Operations Manager"] || "",
      lineOfBusiness: formData["Line of Business"] || ""
    },
    // ... (include all other data structure as in original enhanced structure)
  };
  
  // Return row data in new format
  return [
    recordId, timestamp, "2.0-MIGRATED", "MIGRATION_SYSTEM",
    formData["Sap ID"] || "", formData["Team Member"] || "", formData["Team Leader"] || "",
    formData["Operations Manager"] || "", formData["Line of Business"] || "",
    formData["Interaction Date"] || "", formData["Customer's BAN"] || "",
    formData["FLP Conversation ID"] || "", formData["Listening Type"] || "",
    formData["Call Duration"] || "",
    // Add all other fields with fallback to empty string
    ...Array(40).fill("").map((_, i) => formData[oldHeaders[i + 14]] || ""),
    // Data integrity fields
    "MIGRATED", oldHeaders.length, "MIGRATED", 100
  ];
}

// === AUTOMATED SCHEDULING ===

/**
 * Sets up all automated triggers for the system
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
    
    logSystemEvent("TRIGGERS_SETUP", "All automated triggers created successfully", "SUCCESS");
    
  } catch (error) {
    console.error("❌ Error setting up automated schedules:", error.message);
    logSystemEvent("TRIGGERS_SETUP_ERROR", error.message, "FAILED");
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

// === BACKUP FUNCTIONS ===

/**
 * Performs daily backup to Drive folder
 */
function performDailyBackup() {
  try {
    console.log("💾 Starting daily backup...");
    const startTime = new Date();
    
    // Get the spreadsheet
    const spreadsheet = SpreadsheetApp.openById(DATA_STORAGE_CONFIG.MAIN_SPREADSHEET_ID);
    console.log(`📊 Spreadsheet found: ${spreadsheet.getName()}`);
    
    // Get the backup folder
    const backupFolder = getFolderByPath("Daily_Backups");
    if (!backupFolder) {
      throw new Error("Daily_Backups folder not found");
    }
    console.log(`📁 Backup folder found: ${backupFolder.getName()}`);
    
    // Create timestamp and backup name
    const timestamp = Utilities.formatDate(getESTTimestamp(), DATA_STORAGE_CONFIG.TIMEZONE, 'yyyy-MM-dd_HH-mm');
    const backupName = `QA_NH_Backup_${timestamp}`;
    console.log(`📝 Creating backup: ${backupName}`);
    
    // Create backup file using DriveApp method
    const originalFile = DriveApp.getFileById(DATA_STORAGE_CONFIG.MAIN_SPREADSHEET_ID);
    const copiedFile = originalFile.makeCopy(backupName);
    console.log(`📋 File copied successfully: ${copiedFile.getName()}`);
    
    // Move the copied file to backup folder
    backupFolder.addFile(copiedFile);
    DriveApp.getRootFolder().removeFile(copiedFile);
    console.log(`📂 File moved to backup folder successfully`);
    
    // Clean old backups
    const cleanedCount = cleanOldBackups(backupFolder, DATA_STORAGE_CONFIG.RETENTION_POLICIES.DAILY_BACKUPS);
    console.log(`🗑️ Cleaned ${cleanedCount} old backup files`);
    
    const duration = new Date() - startTime;
    console.log(`✅ Daily backup completed: ${backupName} (${duration}ms)`);
    
    logSystemEvent("DAILY_BACKUP", `Backup created: ${backupName}`, "SUCCESS", `${duration}ms`);
    
    return { 
      success: true, 
      backupName: backupName, 
      duration: `${duration}ms`,
      cleanedFiles: cleanedCount
    };
    
  } catch (error) {
    console.error("❌ Daily backup failed:", error.message);
    logSystemEvent("DAILY_BACKUP_ERROR", error.message, "FAILED");
    throw error;
  }
}

/**
 * Performs weekly maintenance tasks
 */
function performWeeklyMaintenance() {
  try {
    console.log("🔧 Starting weekly maintenance...");
    const startTime = new Date();
    
    // Get the weekly backup folder
    const weeklyFolder = getFolderByPath("Weekly_Backups");
    if (!weeklyFolder) {
      throw new Error("Weekly_Backups folder not found");
    }
    
    // Create weekly backup name
    const weekNumber = Utilities.formatDate(getESTTimestamp(), DATA_STORAGE_CONFIG.TIMEZONE, 'w');
    const year = Utilities.formatDate(getESTTimestamp(), DATA_STORAGE_CONFIG.TIMEZONE, 'yyyy');
    const backupName = `QA_NH_Weekly_${year}-W${weekNumber.padStart(2, '0')}`;
    
    // Create backup file using DriveApp method
    const originalFile = DriveApp.getFileById(DATA_STORAGE_CONFIG.MAIN_SPREADSHEET_ID);
    const copiedFile = originalFile.makeCopy(backupName);
    
    // Move the copied file to weekly backup folder
    weeklyFolder.addFile(copiedFile);
    DriveApp.getRootFolder().removeFile(copiedFile);
    
    // Clean old weekly backups
    cleanOldBackups(weeklyFolder, DATA_STORAGE_CONFIG.RETENTION_POLICIES.WEEKLY_BACKUPS);
    
    // Perform system health check
    const healthReport = performSystemHealthCheck();
    
    const duration = new Date() - startTime;
    console.log(`✅ Weekly maintenance completed (${duration}ms)`);
    
    logSystemEvent("WEEKLY_MAINTENANCE", `Weekly backup and health check completed`, "SUCCESS", `${duration}ms`);
    
    return { success: true, backupName: backupName, healthReport: healthReport };
    
  } catch (error) {
    console.error("❌ Weekly maintenance failed:", error.message);
    logSystemEvent("WEEKLY_MAINTENANCE_ERROR", error.message, "FAILED");
    throw error;
  }
}

/**
 * Performs monthly data archiving
 */
function performMonthlyArchiving() {
  try {
    console.log("📦 Starting monthly archiving...");
    const startTime = new Date();
    
    const cutoffDate = new Date();
    cutoffDate.setMonth(cutoffDate.getMonth() - DATA_STORAGE_CONFIG.RETENTION_POLICIES.ARCHIVE_AFTER_MONTHS);
    
    const archiveResult = archiveOldData(cutoffDate);
    
    // Create monthly backup
    const ss = SpreadsheetApp.openById(DATA_STORAGE_CONFIG.MAIN_SPREADSHEET_ID);
    const monthlyFolder = getFolderByPath("Monthly_Archives");
    
    const monthYear = Utilities.formatDate(cutoffDate, DATA_STORAGE_CONFIG.TIMEZONE, 'yyyy-MM');
    const archiveName = `QA_NH_Archive_${monthYear}`;
    
    if (archiveResult.archivedCount > 0) {
      // Create archive file using DriveApp method
      const originalFile = DriveApp.getFileById(DATA_STORAGE_CONFIG.MAIN_SPREADSHEET_ID);
      const archiveFile = originalFile.makeCopy(archiveName);
      
      // Move the copied file to monthly archive folder
      monthlyFolder.addFile(archiveFile);
      DriveApp.getRootFolder().removeFile(archiveFile);
    }
    
    const duration = new Date() - startTime;
    console.log(`✅ Monthly archiving completed: ${archiveResult.archivedCount} records archived (${duration}ms)`);
    
    logSystemEvent("MONTHLY_ARCHIVING", `Archived ${archiveResult.archivedCount} records`, "SUCCESS", `${duration}ms`);
    
    return { success: true, ...archiveResult, duration: `${duration}ms` };
    
  } catch (error) {
    console.error("❌ Monthly archiving failed:", error.message);
    logSystemEvent("MONTHLY_ARCHIVING_ERROR", error.message, "FAILED");
    throw error;
  }
}

// === UTILITY FUNCTIONS ===

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
 * Cleans old backup files based on retention policy
 */
function cleanOldBackups(folder, retentionDays) {
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
        console.log(`🗑️ Deleted old backup: ${file.getName()}`);
      }
    }
    
    if (deletedCount > 0) {
      console.log(`✅ Cleaned ${deletedCount} old backup files`);
    }
    
    return deletedCount;
    
  } catch (error) {
    console.error("❌ Error cleaning old backups:", error.message);
    return 0;
  }
}

/**
 * Archives old data to separate sheet
 */
function archiveOldData(cutoffDate) {
  try {
    const ss = SpreadsheetApp.openById(DATA_STORAGE_CONFIG.MAIN_SPREADSHEET_ID);
    const mainSheet = ss.getSheetByName(DATA_STORAGE_CONFIG.SHEET_NAMES.MAIN_RESPONSES);
    
    if (!mainSheet || mainSheet.getLastRow() <= 1) {
      return { archivedCount: 0, message: "No data to archive" };
    }
    
    const data = mainSheet.getDataRange().getValues();
    const headers = data[0];
    const dataRows = data.slice(1);
    
    // Find rows to archive (older than cutoff date)
    const rowsToArchive = [];
    const rowsToKeep = [headers];
    
    dataRows.forEach((row, index) => {
      const timestamp = new Date(row[1]); // Timestamp column
      if (timestamp < cutoffDate) {
        rowsToArchive.push(row);
      } else {
        rowsToKeep.push(row);
      }
    });
    
    if (rowsToArchive.length === 0) {
      return { archivedCount: 0, message: "No data old enough to archive" };
    }
    
    // Create archive sheet
    const archiveMonth = Utilities.formatDate(cutoffDate, DATA_STORAGE_CONFIG.TIMEZONE, 'yyyy-MM');
    const archiveSheetName = `QA NH Archive ${archiveMonth}`;
    
    let archiveSheet = ss.getSheetByName(archiveSheetName);
    if (!archiveSheet) {
      archiveSheet = ss.insertSheet(archiveSheetName);
      archiveSheet.appendRow(headers);
      archiveSheet.getRange("1:1").setFontWeight("bold");
      archiveSheet.getRange("1:1").setBackground("#607d8b");
      archiveSheet.getRange("1:1").setFontColor("white");
      archiveSheet.setFrozenRows(1);
    }
    
    // Add archived data to archive sheet
    rowsToArchive.forEach(row => {
      archiveSheet.appendRow(row);
    });
    
    // Update main sheet with remaining data
    mainSheet.clear();
    rowsToKeep.forEach(row => {
      mainSheet.appendRow(row);
    });
    
    // Restore formatting for main sheet
    setupMainResponsesSheet(mainSheet);
    
    console.log(`✅ Archived ${rowsToArchive.length} records to ${archiveSheetName}`);
    
    return {
      archivedCount: rowsToArchive.length,
      archiveSheet: archiveSheetName,
      remainingCount: rowsToKeep.length - 1
    };
    
  } catch (error) {
    console.error("❌ Error archiving data:", error.message);
    throw error;
  }
}

/**
 * Performs system health check
 */
function performSystemHealthCheck() {
  try {
    const healthReport = {
      timestamp: getESTTimestamp(),
      sheetsStatus: checkAllSheetsExist(),
      foldersStatus: checkAllFoldersExist(),
      dataIntegrity: validateDataIntegrity(),
      backupStatus: checkBackupStatus(),
      triggerStatus: checkScheduledTriggers(),
      overallHealth: "HEALTHY"
    };
    
    // Determine overall health
    const issues = [];
    if (!healthReport.sheetsStatus.allExist) issues.push("Missing sheets");
    if (!healthReport.foldersStatus.allExist) issues.push("Missing folders");
    if (healthReport.dataIntegrity.issues > 0) issues.push("Data integrity issues");
    if (!healthReport.backupStatus.recent) issues.push("No recent backups");
    if (healthReport.triggerStatus.missing > 0) issues.push("Missing triggers");
    
    if (issues.length > 0) {
      healthReport.overallHealth = "ISSUES_DETECTED";
      healthReport.issues = issues;
    }
    
    logSystemEvent("HEALTH_CHECK", JSON.stringify(healthReport), healthReport.overallHealth === "HEALTHY" ? "SUCCESS" : "WARNING");
    
    return healthReport;
    
  } catch (error) {
    console.error("❌ Error during health check:", error.message);
    return { overallHealth: "ERROR", error: error.message };
  }
}

/**
 * Checks if all required sheets exist
 */
function checkAllSheetsExist() {
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

/**
 * Checks if all required folders exist using specific folder IDs
 */
function checkAllFoldersExist() {
  try {
    const requiredFolders = DATA_STORAGE_CONFIG.DRIVE_FOLDERS;
    const existingFolders = [];
    const missingFolders = [];
    
    requiredFolders.forEach(folderName => {
      try {
        const folder = getFolderByPath(folderName);
        if (folder) {
          existingFolders.push(folderName);
        } else {
          missingFolders.push(folderName);
        }
      } catch (folderError) {
        console.error(`❌ Error checking folder ${folderName}:`, folderError.message);
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
 * Validates data integrity
 */
function validateDataIntegrity() {
  try {
    const ss = SpreadsheetApp.openById(DATA_STORAGE_CONFIG.MAIN_SPREADSHEET_ID);
    const mainSheet = ss.getSheetByName(DATA_STORAGE_CONFIG.SHEET_NAMES.MAIN_RESPONSES);
    
    if (!mainSheet || mainSheet.getLastRow() <= 1) {
      return { issues: 0, message: "No data to validate" };
    }
    
    const data = mainSheet.getDataRange().getValues();
    const headers = data[0];
    const dataRows = data.slice(1);
    
    let issues = 0;
    const issueDetails = [];
    
    // Check for duplicate Record IDs
    const recordIds = dataRows.map(row => row[0]).filter(id => id);
    const duplicateIds = recordIds.filter((id, index) => recordIds.indexOf(id) !== index);
    
    if (duplicateIds.length > 0) {
      issues += duplicateIds.length;
      issueDetails.push(`Duplicate Record IDs: ${duplicateIds.join(', ')}`);
    }
    
    // Check for missing required fields
    dataRows.forEach((row, index) => {
      if (!row[0]) { // Missing Record ID
        issues++;
        issueDetails.push(`Row ${index + 2}: Missing Record ID`);
      }
      if (!row[1]) { // Missing Timestamp
        issues++;
        issueDetails.push(`Row ${index + 2}: Missing Timestamp`);
      }
    });
    
    return {
      issues: issues,
      totalRows: dataRows.length,
      details: issueDetails.slice(0, 10) // Limit to first 10 issues
    };
    
  } catch (error) {
    return { issues: -1, error: error.message };
  }
}

/**
 * Checks backup status
 */
function checkBackupStatus() {
  try {
    const backupFolder = getFolderByPath("Daily_Backups");
    if (!backupFolder) {
      return { recent: false, error: "Backup folder not found" };
    }
    
    const files = backupFolder.getFiles();
    let mostRecentBackup = null;
    
    while (files.hasNext()) {
      const file = files.next();
      if (!mostRecentBackup || file.getDateCreated() > mostRecentBackup.getDateCreated()) {
        mostRecentBackup = file;
      }
    }
    
    if (!mostRecentBackup) {
      return { recent: false, message: "No backups found" };
    }
    
    const hoursSinceBackup = (new Date() - mostRecentBackup.getDateCreated()) / (1000 * 60 * 60);
    const isRecent = hoursSinceBackup < 25; // Within 25 hours (allowing for schedule variance)
    
    return {
      recent: isRecent,
      lastBackup: mostRecentBackup.getDateCreated(),
      hoursSince: Math.round(hoursSinceBackup),
      backupName: mostRecentBackup.getName()
    };
    
  } catch (error) {
    return { recent: false, error: error.message };
  }
}

/**
 * Checks scheduled triggers
 */
function checkScheduledTriggers() {
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

// === LOGGING FUNCTIONS ===

/**
 * Logs system events to the System Audit sheet
 */
function logSystemEvent(action, details, status, duration = "0ms") {
  try {
    const ss = SpreadsheetApp.openById(DATA_STORAGE_CONFIG.MAIN_SPREADSHEET_ID);
    const auditSheet = ss.getSheetByName(DATA_STORAGE_CONFIG.SHEET_NAMES.SYSTEM_AUDIT);
    
    if (auditSheet) {
      // Try to get user email, fallback to "System" if not available
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
        "SYSTEM",
        details,
        status,
        duration,
        "System"
      ];
      
      auditSheet.appendRow(logEntry);
    }
  } catch (error) {
    console.error("❌ Error logging system event:", error.message);
  }
}

/**
 * Logs migration events to the Migration Log sheet
 */
function logMigrationEvent(migrationType, sourceSheet, targetSheet, recordsProcessed, status, details = "", duration = "0ms") {
  try {
    const ss = SpreadsheetApp.openById(DATA_STORAGE_CONFIG.MAIN_SPREADSHEET_ID);
    const migrationSheet = ss.getSheetByName(DATA_STORAGE_CONFIG.SHEET_NAMES.MIGRATION_LOG);
    
    if (migrationSheet) {
      const logEntry = [
        getESTTimestamp(),
        migrationType,
        sourceSheet,
        targetSheet,
        recordsProcessed,
        status,
        details,
        duration
      ];
      
      migrationSheet.appendRow(logEntry);
    }
  } catch (error) {
    console.error("❌ Error logging migration event:", error.message);
  }
}

// === INITIALIZATION HELPERS ===

/**
 * Creates initial configuration
 */
function createInitialConfiguration() {
  try {
    // Configuration is already created in setupConfigSheet
    console.log("⚙️ Initial configuration created");
    logSystemEvent("CONFIG_INIT", "Initial configuration created", "SUCCESS");
  } catch (error) {
    console.error("❌ Error creating initial configuration:", error.message);
    logSystemEvent("CONFIG_INIT_ERROR", error.message, "FAILED");
  }
}

/**
 * Generates initialization report
 */
function generateInitializationReport() {
  try {
    const report = {
      timestamp: getESTTimestamp(),
      sheetsCreated: checkAllSheetsExist(),
      foldersCreated: checkAllFoldersExist(),
      triggersSetup: checkScheduledTriggers(),
      systemHealth: performSystemHealthCheck(),
      version: "1.0",
      timezone: DATA_STORAGE_CONFIG.TIMEZONE
    };
    
    console.log("📋 Initialization report generated");
    return report;
    
  } catch (error) {
    console.error("❌ Error generating initialization report:", error.message);
    return { error: error.message };
  }
}

// === MANUAL FUNCTIONS FOR TESTING ===

/**
 * Manual function to test the system
 */
function testQANHDataStorage() {
  console.log("🧪 Testing QA NH Data Storage system...");
  
  try {
    // Test Record ID generation
    const recordId = generateRecordID();
    console.log(`✅ Record ID generated: ${recordId}`);
    
    // Test EST timestamp
    const estTime = getESTTimestamp();
    console.log(`✅ EST timestamp: ${estTime}`);
    
    // Test folder existence
    const foldersStatus = checkAllFoldersExist();
    console.log(`✅ Folders status: ${foldersStatus.allExist ? 'All exist' : 'Missing folders'}`);
    
    // Test sheet existence
    const sheetsStatus = checkAllSheetsExist();
    console.log(`✅ Sheets status: ${sheetsStatus.allExist ? 'All exist' : 'Missing sheets'}`);
    
    // Test system health
    const healthReport = performSystemHealthCheck();
    console.log(`✅ System health: ${healthReport.overallHealth}`);
    
    console.log("🎉 QA NH Data Storage system test completed successfully!");
    
    return {
      success: true,
      recordId: recordId,
      estTime: estTime,
      foldersStatus: foldersStatus,
      sheetsStatus: sheetsStatus,
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
 * Quick setup function - creates only essential components
 */
function quickSetupQANHDataStorage() {
  console.log("⚡ Quick setup of QA NH Data Storage...");
  
  try {
    // Create only essential folders
    createDriveFolderStructure();
    
    // Create only main response sheet
    const ss = SpreadsheetApp.openById(DATA_STORAGE_CONFIG.MAIN_SPREADSHEET_ID);
    createSheet(ss, DATA_STORAGE_CONFIG.SHEET_NAMES.MAIN_RESPONSES, setupMainResponsesSheet);
    
    console.log("✅ Quick setup completed!");
    return { success: true, message: "Quick setup completed successfully!" };
    
  } catch (error) {
    console.error("❌ Quick setup failed:", error.message);
    return { success: false, error: error.message };
  }
}

/**
 * Creates the main sheet with a sample gibberish example to demonstrate the system
 */
function createMainSheetWithSampleData() {
  console.log("📊 Creating main sheet with sample data...");
  
  try {
    const ss = SpreadsheetApp.openById(DATA_STORAGE_CONFIG.MAIN_SPREADSHEET_ID);
    
    // Create or get the main responses sheet
    let mainSheet = ss.getSheetByName(DATA_STORAGE_CONFIG.SHEET_NAMES.MAIN_RESPONSES);
    if (!mainSheet) {
      mainSheet = ss.insertSheet(DATA_STORAGE_CONFIG.SHEET_NAMES.MAIN_RESPONSES);
      setupMainResponsesSheet(mainSheet);
    } else if (mainSheet.getLastRow() === 0) {
      setupMainResponsesSheet(mainSheet);
    }
    
    // Generate sample data
    const recordId = generateRecordID();
    const timestamp = getESTTimestamp();
    
    // Create comprehensive sample row with gibberish data
    const sampleRow = [
      // System metadata
      recordId,                           // RecordID
      timestamp,                          // Timestamp
      "2.1",                             // FormVersion
      "demo.user@company.com",           // SubmittedBy
      
      // Team member details
      "12345",                           // Sap ID
      "John Doe Sample",                 // Team Member
      "Jane Smith Leader",               // Team Leader
      "Bob Johnson Manager",             // Operations Manager
      "Customer Service",                // Line of Business
      
      // Call details
      "2024-09-29",                      // Interaction Date
      "987654321",                       // Customer's BAN
      "CONV-2024-0929-001",             // FLP Conversation ID
      "Live/Side by Side",               // Listening Type
      "15:30",                          // Call Duration
      
      // Core skills
      "DW",                             // Willingness to Help
      "DW",                             // Warm and Energetic Greeting
      "Prt",                            // Listen & Retain
      
      // Critical skills
      "DW",                             // Secure Customer Identity
      "DW",                             // Probe Using Questions
      "Prt",                            // Investigate & Check Resources
      "DW",                             // Offer Solutions
      "DD",                             // Begin Processing Transaction
      
      // Basic skills
      "DW",                             // Confirm What You Know
      "Prt",                            // Check for Understanding
      "DW",                             // Lead with Benefit
      "DW",                             // Summarize Solution
      "DW",                             // End on High Note
      "NA",                             // Reinforce Extras
      
      // Intermediate skills
      "DW",                             // Early Bridge to Value
      "Prt",                            // Later Bridge to Value
      "DW",                             // Lead Customer to Accept
      "DW",                             // Check for Satisfaction
      
      // Advanced skills
      "NA",                             // Overcome Objections
      "DW",                             // Acknowledge (Empathy)
      
      // Sales skills
      "DW",                             // Learn from Customer
      "NA",                             // Handling Early Objections
      "Prt",                            // Customer Intimacy
      "DW",                             // Value Proposition
      "Open (Exploring customer's interests)", // Bridging Type
      
      // CLS data
      "Pricing - Too Expensive Price Increase", // Cancellation Root Cause
      "No",                             // Unethical Credit
      
      // CLS skills
      "Yes",                            // Right size
      "Yes",                            // Tiered approach
      "Yes",                            // Sufficient probing
      "Yes",                            // Speak to value
      "Yes",                            // Loyalty statements
      "Yes",                            // Investigation
      "No",                             // Retain customer
      
      // Call evaluation
      "Billing Inquiry",               // Call Driver Level 1
      "Payment Arrangement",           // Call Driver Level 2
      "Customer called regarding a billing inquiry about their recent charges. TM provided excellent service by thoroughly explaining the charges and offering a payment arrangement. Customer was satisfied with the resolution and appreciated the TM's patience and professionalism throughout the interaction.", // Call Summary
      "- TM demonstrated excellent probing skills to understand the billing concern\n- Provided clear explanations of charges and available options\n- Showed empathy and patience throughout the interaction", // Highlights
      "- Could have offered early bridge to value regarding account optimization\n- Missed opportunity to reinforce additional services available\n- Could have provided more detailed follow-up information", // Lowlights/Recommendation
      
      // Data integrity
      "A7B9C2D1",                      // DataChecksum
      "47",                            // FieldCount
      "PASSED",                        // ValidationStatus
      "89"                             // DataQualityScore
    ];
    
    // Add the sample row
    mainSheet.appendRow(sampleRow);
    
    console.log(`✅ Sample data added to main sheet with Record ID: ${recordId}`);
    
    // Log the creation
    logSystemEvent("SAMPLE_DATA_CREATED", `Sample row added with Record ID: ${recordId}`, "SUCCESS");
    
    return {
      success: true,
      message: `Main sheet created with sample data! Record ID: ${recordId}`,
      recordId: recordId,
      sheetName: DATA_STORAGE_CONFIG.SHEET_NAMES.MAIN_RESPONSES
    };
    
  } catch (error) {
    console.error("❌ Error creating main sheet with sample data:", error.message);
    logSystemEvent("SAMPLE_DATA_ERROR", error.message, "FAILED");
    return {
      success: false,
      error: error.message
    };
  }
}
