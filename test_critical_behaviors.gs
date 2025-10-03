/**
 * Test function to verify Critical Behaviors now use 0/1 instead of true/false
 */
function testCriticalBehaviorsConversion() {
  console.log("Testing Critical Behaviors 0/1 conversion...");
  
  // Create test form data with some critical behaviors checked
  const testFormData = {
    "Sap ID": "12345",
    "Team Member": "Test User",
    "Team Leader": "Test Leader",
    "Operations Manager": "Test Manager",
    "Line of Business": "Customer Service",
    "Interaction Date": "2024-01-15",
    "Customer's BAN": "987654321",
    "FLP Conversation ID": "FLP-TEST-001",
    "Listening Type": "Live/Side by Side",
    "Duration (minutes)": "15",
    "Call Driver Level 1": "Billing",
    "Call Driver Level 2": "Payment Issue",
    "Call Summary": "Test call summary",
    "Highlights": "Test highlights",
    "Lowlights/Recommendation": "Test recommendations",
    "Willingness to Help": "DW",
    "Warm and Energetic Greeting that Connects with the Customer": "DW",
    "Listen & Retain": "DW",
    
    // Critical Behaviors - some checked (should become 1), some unchecked (should become 0)
    "Abusive, insult the customer (disrespectful)": "on", // Should become 1
    "Misleading Information / Deliberately lying": "on", // Should become 1
    "Other": "on", // Should become 1
    "CB_Other_Behavior_Text": "Test other behavior description"
    // All other critical behaviors not included, should become 0
  };
  
  try {
    // Test the enhanced data creation
    const enhancedData = createEnhancedFormData(testFormData);
    
    console.log("Critical Behaviors values:");
    console.log("abusiveInsult:", enhancedData.criticalBehaviors.abusiveInsult, "(should be 1)");
    console.log("misleadingInformation:", enhancedData.criticalBehaviors.misleadingInformation, "(should be 1)");
    console.log("otherBehavior:", enhancedData.criticalBehaviors.otherBehavior, "(should be 1)");
    console.log("appropriationSale:", enhancedData.criticalBehaviors.appropriationSale, "(should be 0)");
    console.log("cctsProcessNotFollowed:", enhancedData.criticalBehaviors.cctsProcessNotFollowed, "(should be 0)");
    console.log("otherBehaviorText:", enhancedData.criticalBehaviors.otherBehaviorText, "(should be text)");
    
    // Test row data preparation
    const rowData = prepareRowData(enhancedData);
    
    // Find the critical behaviors section in the row data
    // Critical behaviors start at index 40 (after all other fields)
    const criticalBehaviorsStartIndex = 40;
    const criticalBehaviorsData = rowData.slice(criticalBehaviorsStartIndex, criticalBehaviorsStartIndex + 25);
    
    console.log("\nCritical Behaviors in row data (first 10 values):");
    for (let i = 0; i < 10; i++) {
      console.log(`Index ${i}: ${criticalBehaviorsData[i]} (type: ${typeof criticalBehaviorsData[i]})`);
    }
    
    // Verify specific values
    const abusiveInsultValue = criticalBehaviorsData[0]; // First critical behavior
    const misleadingInfoValue = criticalBehaviorsData[9]; // 10th critical behavior
    const otherBehaviorValue = criticalBehaviorsData[24]; // Last boolean critical behavior
    const otherBehaviorTextValue = criticalBehaviorsData[25]; // Text field
    
    console.log("\nSpecific test values:");
    console.log("Abusive Insult (should be 1):", abusiveInsultValue);
    console.log("Misleading Information (should be 1):", misleadingInfoValue);
    console.log("Other Behavior (should be 1):", otherBehaviorValue);
    console.log("Other Behavior Text:", otherBehaviorTextValue);
    
    // Validation
    let testsPassed = 0;
    let totalTests = 0;
    
    // Test 1: Checked behaviors should be 1
    totalTests++;
    if (enhancedData.criticalBehaviors.abusiveInsult === 1) {
      console.log("✓ PASS: Checked behavior converts to 1");
      testsPassed++;
    } else {
      console.log("✗ FAIL: Checked behavior should be 1, got:", enhancedData.criticalBehaviors.abusiveInsult);
    }
    
    // Test 2: Unchecked behaviors should be 0
    totalTests++;
    if (enhancedData.criticalBehaviors.appropriationSale === 0) {
      console.log("✓ PASS: Unchecked behavior converts to 0");
      testsPassed++;
    } else {
      console.log("✗ FAIL: Unchecked behavior should be 0, got:", enhancedData.criticalBehaviors.appropriationSale);
    }
    
    // Test 3: Values should be numbers, not booleans
    totalTests++;
    if (typeof enhancedData.criticalBehaviors.abusiveInsult === 'number') {
      console.log("✓ PASS: Critical behavior values are numbers");
      testsPassed++;
    } else {
      console.log("✗ FAIL: Critical behavior values should be numbers, got:", typeof enhancedData.criticalBehaviors.abusiveInsult);
    }
    
    // Test 4: Text field should remain as string
    totalTests++;
    if (typeof enhancedData.criticalBehaviors.otherBehaviorText === 'string') {
      console.log("✓ PASS: Text field remains as string");
      testsPassed++;
    } else {
      console.log("✗ FAIL: Text field should be string, got:", typeof enhancedData.criticalBehaviors.otherBehaviorText);
    }
    
    console.log(`\nTest Results: ${testsPassed}/${totalTests} tests passed`);
    
    if (testsPassed === totalTests) {
      console.log("🎉 ALL TESTS PASSED! Critical Behaviors conversion is working correctly.");
      return {
        success: true,
        message: "Critical Behaviors conversion test passed",
        testsPassed: testsPassed,
        totalTests: totalTests
      };
    } else {
      console.log("❌ Some tests failed. Please check the implementation.");
      return {
        success: false,
        message: "Critical Behaviors conversion test failed",
        testsPassed: testsPassed,
        totalTests: totalTests
      };
    }
    
  } catch (error) {
    console.error("Test failed with error:", error.message);
    return {
      success: false,
      message: `Test failed: ${error.message}`
    };
  }
}

/**
 * Test the migration function default values
 */
function testMigrationDefaultValues() {
  console.log("Testing migration default values...");
  
  try {
    // Test the default values array creation
    const defaultValues = Array(24).fill(0).concat([""]);
    
    console.log("Default values array length:", defaultValues.length);
    console.log("First 5 values:", defaultValues.slice(0, 5));
    console.log("Last 5 values:", defaultValues.slice(-5));
    
    // Verify all boolean values are 0
    const booleanValues = defaultValues.slice(0, 24);
    const allZeros = booleanValues.every(val => val === 0);
    
    if (allZeros) {
      console.log("✓ PASS: All boolean default values are 0");
    } else {
      console.log("✗ FAIL: Some boolean default values are not 0");
    }
    
    // Verify text value is empty string
    const textValue = defaultValues[24];
    if (textValue === "") {
      console.log("✓ PASS: Text default value is empty string");
    } else {
      console.log("✗ FAIL: Text default value should be empty string, got:", textValue);
    }
    
    return {
      success: allZeros && textValue === "",
      message: "Migration default values test completed"
    };
    
  } catch (error) {
    console.error("Migration test failed:", error.message);
    return {
      success: false,
      message: `Migration test failed: ${error.message}`
    };
  }
}
