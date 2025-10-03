# Critical Behaviors Data Migration Guide

## Overview
This guide explains how to migrate existing data to include the new Critical Behaviors columns that were added to the QA evaluation form.

## What the Migration Does

### 1. **Backup Creation**
- Creates a timestamped backup sheet before any changes
- Backup format: `QA NH Form Responses_Backup_YYYY-MM-DDTHH-MM-SS`
- Includes metadata note with migration timestamp

### 2. **Column Addition**
Adds 25 new columns for Critical Behaviors:

**Boolean Columns (24 total):**
- CB_Abusive_Insult_Customer
- CB_Appropriation_Sale_Another_Agent
- CB_CCTS_Process_Not_Followed
- CB_Delayed_Response_Inbound
- CB_Didnt_Complete_Transaction_Promised
- CB_Excessive_Dead_Air_Hold_Idle
- CB_Failure_Follow_Processes_Onesource
- CB_Incorrect_Sales_Order_Processing
- CB_Incorrect_Use_Credits_Discounts
- CB_Misleading_Information_Lying
- CB_Misleading_Product_Offer_Info
- CB_Missed_Call_Back
- CB_Modify_Account_Without_Consent
- CB_No_Authentication_Verification
- CB_No_Response_Inbound_Contact
- CB_No_Retention_Value_Generation
- CB_Nonwork_Related_Distractions
- CB_Overtalking_Interrupt_Customer
- CB_Proactively_Transferring_Calls
- CB_Staying_Line_Unnecessarily
- CB_System_Tools_Manipulation
- CB_Unjustified_Disconnection
- CB_Unnecessary_Invalid_Consult_Transfer
- CB_Unwillingness_To_Help
- CB_Other_Behavior

**Text Column (1 total):**
- CB_Other_Behavior_Text

### 3. **Data Population**
- Sets all boolean columns to `FALSE` for existing rows
- Sets text column to empty string `""` for existing rows
- Processes data in batches of 100 rows to avoid timeouts

## How to Run Migration

### Option 1: Using the UI Button (Recommended)
1. Open the QA Evaluation Form
2. Look for the purple "Migrate Data" button in the Team Member Details section
3. Click the button
4. Confirm the migration when prompted
5. Wait for completion message

### Option 2: Using Google Apps Script Editor
1. Open Google Apps Script editor
2. Navigate to Code.gs
3. Run the `migrateCriticalBehaviorsData()` function
4. Check the execution log for results

### Option 3: Using Test Function
1. Run `testMigration()` function for a test run
2. This provides detailed logging and validation

## Migration Safety Features

### **Idempotent Operation**
- Can be run multiple times safely
- Automatically detects if migration has already been completed
- Will not duplicate columns or data

### **Automatic Backup**
- Creates backup before any changes
- Backup includes all original data
- Timestamped for easy identification

### **Batch Processing**
- Processes large datasets in batches
- Prevents timeout errors
- Shows progress for datasets > 100 rows

### **Validation**
- Validates spreadsheet structure before proceeding
- Confirms successful migration after completion
- Reports any issues found

## Expected Results

### **Before Migration:**
- Original columns: ~40 columns
- No Critical Behaviors data

### **After Migration:**
- Total columns: ~65 columns (original + 25 new)
- All existing rows have Critical Behaviors columns with default values
- New submissions will populate Critical Behaviors data

## Rollback Procedure

### **If Migration Needs to be Reversed:**
1. Use the `rollbackMigration(backupSheetName)` function
2. Provide the exact backup sheet name from migration results
3. This will restore the original data structure

### **Example Rollback:**
```javascript
rollbackMigration("QA NH Form Responses_Backup_2024-01-15T10-30-45");
```

## Troubleshooting

### **Common Issues:**

**1. "Migration already completed" message**
- This is normal if migration was run before
- No action needed - data is already migrated

**2. Timeout errors**
- Migration uses batch processing to prevent this
- If it occurs, re-run the migration (it will resume safely)

**3. Permission errors**
- Ensure you have edit access to the spreadsheet
- Contact administrator if needed

**4. Column count mismatch**
- Check if manual columns were added to the sheet
- Verify the expected column count in validation results

### **Validation Checks:**
- Minimum expected columns: 65
- Required headers: CB_Abusive_Insult_Customer, CB_Other_Behavior_Text
- All existing data preserved

## Migration Results

### **Success Response Example:**
```json
{
  "success": true,
  "message": "Critical Behaviors data migration completed successfully!",
  "details": {
    "backupSheet": "QA NH Form Responses_Backup_2024-01-15T10-30-45",
    "originalRows": 150,
    "originalColumns": 40,
    "newColumns": 25,
    "rowsUpdated": 149,
    "totalColumns": 65,
    "migrationTime": "2024-01-15T10:30:45.123Z"
  }
}
```

## Post-Migration Verification

### **Manual Checks:**
1. Verify backup sheet was created
2. Check that new columns appear at the end of the sheet
3. Confirm existing data is intact
4. Test new form submissions include Critical Behaviors data

### **Automated Checks:**
- Migration includes built-in validation
- Reports any issues found
- Confirms successful completion

## Support

### **For Issues:**
- Check the Google Apps Script execution log
- Review the migration results message
- Contact the development team with specific error messages

### **Available Functions:**
- `migrateCriticalBehaviorsData()` - Main migration function
- `testMigration()` - Test migration with detailed logging
- `listMigrationBackups()` - List all available backups
- `rollbackMigration(backupName)` - Rollback to specific backup

## Data Impact

### **Existing Evaluations:**
- All existing evaluations remain unchanged
- Critical Behaviors columns added with default values (FALSE/empty)
- Can be manually updated if needed for historical analysis

### **New Evaluations:**
- Will include Critical Behaviors data
- Form validates and stores checkbox selections
- Supports "Other" behavior with custom text

### **Reporting:**
- New columns available for filtering and analysis
- Boolean format allows for easy counting and aggregation
- Compatible with existing reporting tools

---

**Migration completed successfully!** The Critical Behaviors functionality is now fully integrated with backward compatibility maintained.
