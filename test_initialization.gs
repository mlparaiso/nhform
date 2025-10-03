/**
 * Test script to initialize the new spreadsheet with sample data
 * This file can be run to test the complete system setup
 */

function testNewSpreadsheetSetup() {
  console.log("🧪 Testing new spreadsheet setup...");
  
  try {
    // Test 1: Create main sheet with sample data
    console.log("📊 Step 1: Creating main sheet with sample data...");
    const sampleResult = createMainSheetWithSampleData();
    console.log("✅ Sample data result:", sampleResult);
    
    // Test 2: Initialize full system
    console.log("🚀 Step 2: Initializing full QA NH Data Storage system...");
    const initResult = initializeQANHDataStorage();
    console.log("✅ Initialization result:", initResult);
    
    // Test 3: Test system health
    console.log("🔍 Step 3: Testing system health...");
    const healthResult = testQANHDataStorage();
    console.log("✅ Health test result:", healthResult);
    
    // Test 4: Test form submission
    console.log("📝 Step 4: Testing enhanced data structure...");
    const dataStructureResult = testEnhancedDataStructure();
    console.log("✅ Data structure test result:", dataStructureResult);
    
    console.log("🎉 All tests completed successfully!");
    
    return {
      success: true,
      message: "New spreadsheet setup completed successfully!",
      results: {
        sampleData: sampleResult,
        initialization: initResult,
        healthCheck: healthResult,
        dataStructure: dataStructureResult
      }
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
 * Quick test to verify the new spreadsheet ID is working
 */
function testSpreadsheetAccess() {
  try {
    const ss = SpreadsheetApp.openById("1InCv7-d8-KhovEGb8tdI3TxLJx3PTou5ZJKsx1nZhwI");
    console.log(`✅ Successfully accessed spreadsheet: ${ss.getName()}`);
    console.log(`📊 Current sheets: ${ss.getSheets().map(sheet => sheet.getName()).join(', ')}`);
    
    return {
      success: true,
      spreadsheetName: ss.getName(),
      existingSheets: ss.getSheets().map(sheet => sheet.getName())
    };
    
  } catch (error) {
    console.error("❌ Cannot access new spreadsheet:", error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Test Drive folder access
 */
function testDriveFolderAccess() {
  try {
    const folder = DriveApp.getFolderById("1AtlMVddT2MbMgSD5LeTrRkIPgZle8E2r");
    console.log(`✅ Successfully accessed Drive folder: ${folder.getName()}`);
    
    const subfolders = [];
    const folders = folder.getFolders();
    while (folders.hasNext()) {
      subfolders.push(folders.next().getName());
    }
    
    console.log(`📁 Existing subfolders: ${subfolders.join(', ')}`);
    
    return {
      success: true,
      folderName: folder.getName(),
      existingSubfolders: subfolders
    };
    
  } catch (error) {
    console.error("❌ Cannot access Drive folder:", error.message);
    return {
      success: false,
      error: error.message
    };
  }
}
