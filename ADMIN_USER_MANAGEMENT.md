# Admin User Management System

## Overview

The QA New Hire Evaluation Form includes an admin user management system that restricts certain sensitive features to authorized administrators only. This system provides both client-side UI controls and server-side security validation.

## Admin-Only Features

The following features are restricted to admin users only:

### 1. Fill Test Data Button
- **Location**: Team Member Details section
- **Purpose**: Populates the form with comprehensive test data for evaluation testing
- **Security**: Hidden from non-admin users, no server-side restrictions needed (client-side only feature)

### 2. Debug Data Button
- **Location**: Team Member Details section  
- **Purpose**: Provides diagnostic information about the spreadsheet data structure
- **Security**: Hidden from non-admin users + server-side validation in `debugSpreadsheetData()` function

## Current Admin Users

The admin user list is maintained in the `isAdminUser()` function in `Code.gs`:

```javascript
function isAdminUser() {
  try {
    const currentUserEmail = Session.getActiveUser().getEmail();
    const adminEmails = ['michael.paraiso@telus.com'];
    
    Logger.log(`Checking admin status for: ${currentUserEmail}`);
    const isAdmin = adminEmails.includes(currentUserEmail);
    Logger.log(`Admin status: ${isAdmin}`);
    
    return isAdmin;
  } catch (error) {
    Logger.log(`Error checking admin status: ${error.message}`);
    return false; // Default to non-admin if error occurs
  }
}
```

## How It Works

### Client-Side Implementation

1. **Page Load Check**: When the form loads, `checkAdminStatus()` is called
2. **Admin Status Request**: Frontend calls `isAdminUser()` backend function
3. **UI Adjustment**: Based on the response, admin-only buttons are shown/hidden
4. **Button Visibility**: 
   - Admin users: All buttons visible
   - Regular users: "Fill Test Data" and "Debug Data" buttons hidden

### Server-Side Security

1. **Debug Function Protection**: The `debugSpreadsheetData()` function includes admin validation:
   ```javascript
   if (!isAdminUser()) {
     Logger.log(`Non-admin user attempted to access debug data: ${Session.getActiveUser().getEmail()}`);
     throw new Error("Access denied. Debug functionality is restricted to admin users only.");
   }
   ```

2. **Logging**: All admin access attempts are logged for security auditing

## Available Features for All Users

The following features remain accessible to all users:

- **Process Script**: AI script analysis functionality
- **Clear Form**: Form clearing utility
- **Rewrite All with FueliX**: AI text improvement for all text areas
- **Individual Rewrite with FueliX**: AI text improvement for individual fields
- **All form submission and editing capabilities**

## Adding New Admin Users

To add a new admin user:

1. Open `Code.gs` in the Google Apps Script editor
2. Locate the `isAdminUser()` function
3. Add the new email address to the `adminEmails` array:
   ```javascript
   const adminEmails = [
     'michael.paraiso@telus.com',
     'new.admin@telus.com'  // Add new admin email here
   ];
   ```
4. Save the script
5. The changes take effect immediately

## Removing Admin Users

To remove an admin user:

1. Open `Code.gs` in the Google Apps Script editor
2. Locate the `isAdminUser()` function
3. Remove the email address from the `adminEmails` array
4. Save the script

## Security Considerations

### Current Security Level
- **Client-side**: UI controls hide admin buttons from non-admin users
- **Server-side**: Debug function validates admin status before execution
- **Logging**: Admin access attempts are logged for audit trails

### Recommendations for Enhanced Security

If additional admin-only features are added in the future, consider:

1. **Server-side validation** for all admin functions
2. **Role-based permissions** for different admin levels
3. **Session management** for admin authentication
4. **Audit logging** for all admin actions

## Testing Admin Functionality

### Testing as Admin User
1. Ensure you're logged in as `michael.paraiso@telus.com`
2. Load the form - you should see all buttons including "Fill Test Data" and "Debug Data"
3. Test the "Fill Test Data" button - should populate form with test data
4. Test the "Debug Data" button - should show spreadsheet diagnostic information

### Testing as Regular User
1. Log in with any other email address
2. Load the form - "Fill Test Data" and "Debug Data" buttons should be hidden
3. All other functionality should work normally

## Troubleshooting

### Admin Buttons Not Showing
1. Check that you're logged in with the correct admin email
2. Check browser console for any JavaScript errors
3. Verify the email address is correctly added to the `adminEmails` array

### Debug Function Access Denied
1. Verify the user email is in the admin list
2. Check the Google Apps Script logs for admin status check results
3. Ensure the `isAdminUser()` function is working correctly

### Console Logging
Admin status checks are logged to help with troubleshooting:
- `"Checking admin status for: [email]"`
- `"Admin status: [true/false]"`
- `"Admin user [email] accessed debug data"` (for successful debug access)
- `"Non-admin user attempted to access debug data: [email]"` (for blocked access)

## Future Enhancements

Potential improvements to the admin system:

1. **Configuration Sheet**: Move admin emails to a configuration sheet for easier management
2. **Role Levels**: Implement different admin permission levels (e.g., Super Admin, QA Admin, etc.)
3. **Admin Dashboard**: Create an admin-only dashboard for system management
4. **Bulk Operations**: Add admin-only bulk data operations
5. **User Management UI**: Create a UI for managing admin users without code changes

## Change Log

- **2024-09-30**: Initial implementation with basic admin user management
  - Added `isAdminUser()` function
  - Implemented client-side admin button visibility
  - Added server-side validation for debug function
  - Restricted "Fill Test Data" and "Debug Data" to admin users only
