/**
 * Test function to verify backup functionality with specific folder IDs
 */
function testBackupFunctionality() {
  console.log("🧪 Testing backup functionality with specific folder IDs...");
  
  try {
    // Test 1: Check folder access
    console.log("📁 Testing folder access...");
    const dailyBackupFolder = getFolderByPath("Daily_Backups");
    const weeklyBackupFolder = getFolderByPath("Weekly_Backups");
    const monthlyArchiveFolder = getFolderByPath("Monthly_Archives");
    
    if (!dailyBackupFolder) {
      throw new Error("Daily_Backups folder not accessible");
    }
    if (!weeklyBackupFolder) {
      throw new Error("Weekly_Backups folder not accessible");
    }
    if (!monthlyArchiveFolder) {
      throw new Error("Monthly_Archives folder not accessible");
    }
    
    console.log("✅ All backup folders accessible");
    
    // Test 2: Test Record ID generation
    console.log("🆔 Testing Record ID generation...");
    const recordId = generateRecordID();
    console.log(`✅ Generated Record ID: ${recordId}`);
    
    // Test 3: Test EST timestamp
    console.log("🕐 Testing EST timestamp...");
    const estTime = getESTTimestamp();
    console.log(`✅ EST timestamp: ${estTime}`);
    
    // Test 4: Test spreadsheet access
    console.log("📊 Testing spreadsheet access...");
    const ss = SpreadsheetApp.openById(DATA_STORAGE_CONFIG.MAIN_SPREADSHEET_ID);
    console.log(`✅ Spreadsheet accessible: ${ss.getName()}`);
    
    // Test 5: Test backup folder file listing
    console.log("📂 Testing backup folder contents...");
    const files = dailyBackupFolder.getFiles();
    let fileCount = 0;
    while (files.hasNext()) {
      files.next();
      fileCount++;
    }
    console.log(`✅ Daily backup folder contains ${fileCount} files`);
    
    console.log("🎉 All backup functionality tests passed!");
    
    return {
      success: true,
      message: "Backup functionality test completed successfully!",
      recordId: recordId,
      estTime: estTime,
      spreadsheetName: ss.getName(),
      dailyBackupFiles: fileCount
    };
    
  } catch (error) {
    console.error("❌ Backup functionality test failed:", error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Test the actual daily backup function
 */
function testDailyBackupFunction() {
  console.log("💾 Testing actual daily backup function...");
  
  try {
    const result = performDailyBackup();
    console.log("✅ Daily backup test completed successfully!");
    console.log("📊 Backup result:", result);
    
    return result;
    
  } catch (error) {
    console.error("❌ Daily backup test failed:", error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Quick test to verify folder ID mapping
 */
function testFolderIDMapping() {
  console.log("🗂️ Testing folder ID mapping...");
  
  const folderTests = [
    { name: "Daily_Backups", id: DATA_STORAGE_CONFIG.FOLDER_IDS.DAILY_BACKUPS },
    { name: "Weekly_Backups", id: DATA_STORAGE_CONFIG.FOLDER_IDS.WEEKLY_BACKUPS },
    { name: "Monthly_Archives", id: DATA_STORAGE_CONFIG.FOLDER_IDS.MONTHLY_ARCHIVES },
    { name: "Active_Data", id: DATA_STORAGE_CONFIG.FOLDER_IDS.ACTIVE_DATA },
    { name: "Quality_Reports", id: DATA_STORAGE_CONFIG.FOLDER_IDS.QUALITY_REPORTS },
    { name: "System_Logs", id: DATA_STORAGE_CONFIG.FOLDER_IDS.SYSTEM_LOGS },
    { name: "Configuration", id: DATA_STORAGE_CONFIG.FOLDER_IDS.CONFIGURATION }
  ];
  
  const results = [];
  
  folderTests.forEach(test => {
    try {
      const folder = DriveApp.getFolderById(test.id);
      console.log(`✅ ${test.name}: ${folder.getName()} (${test.id})`);
      results.push({ name: test.name, status: "SUCCESS", folderName: folder.getName() });
    } catch (error) {
      console.error(`❌ ${test.name}: Error accessing folder ${test.id} - ${error.message}`);
      results.push({ name: test.name, status: "FAILED", error: error.message });
    }
  });
  
  return {
    success: results.every(r => r.status === "SUCCESS"),
    results: results
  };
}
