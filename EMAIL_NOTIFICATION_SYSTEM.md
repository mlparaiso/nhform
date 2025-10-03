# Email Notification System Documentation

## Overview
The QA Evaluation Form now includes an automated email notification system that sends confirmation emails to users after form submission or updates.

## Features
- **Automatic Confirmation Emails**: Sent after successful form submission
- **Update Notifications**: Sent when existing evaluations are updated
- **Gmail-Compatible HTML Format**: Professional, responsive email design
- **Key Evaluation Details**: Includes Record ID, Team Member, SAP ID, BAN, and FLP ID
- **Error Handling**: Email failures don't prevent form submission

## Email Content Structure

### Subject Line
- **New Submission**: `✅ QA Evaluation Submitted Successfully - Record ID: [RecordID]`
- **Update**: `✅ QA Evaluation Updated Successfully - Record ID: [RecordID]`

### Email Body Includes
- Personalized greeting using user's name (extracted from email)
- Confirmation message
- **Evaluation Details Box** with:
  - Record ID (unique identifier)
  - Team Member name
  - SAP ID
  - BAN (Customer's BAN)
  - FLP ID (FLP Conversation ID)
- Professional footer with system branding

## Technical Implementation

### Key Functions

#### `sendSubmissionConfirmation(formData, recordId, isUpdate)`
Main function that sends confirmation emails
- **Parameters**:
  - `formData`: The submitted form data
  - `recordId`: Generated record ID
  - `isUpdate`: Boolean indicating if this is an update (default: false)

#### `extractNameFromEmail(email)`
Converts email addresses to display names
- Example: `john.doe@company.com` → `John Doe`

#### `generateEmailHTML(userName, recordId, evaluationDetails, isUpdate)`
Creates Gmail-compatible HTML email content
- Responsive design that works on desktop and mobile
- Professional styling with TELUS brand colors

### Integration Points

1. **New Submissions**: Called in `saveFormData()` function after successful save
2. **Updates**: Called in `updateExistingEvaluation()` function after successful update

### Error Handling
- Email failures are logged but don't prevent form submission
- Graceful fallbacks for name extraction and user identification
- Try-catch blocks prevent email errors from breaking the main workflow

## Testing

### Test Function: `testEmailNotification()`
Sends a sample email with test data to verify the system works correctly.

```javascript
function testEmailNotification() {
  // Creates sample data and sends test email
  // Returns success/failure status
}
```

## Configuration

### Email Template Customization
The email template can be customized by modifying the `generateEmailHTML()` function:
- Colors and styling
- Content structure
- Additional evaluation details
- Branding elements

### User Identification
The system uses `Session.getActiveUser().getEmail()` to identify the current user and send emails to their address.

## Security & Privacy
- Only sends emails to the authenticated user who submitted the form
- No external email addresses are used
- Email content only includes evaluation metadata, not sensitive details

## Maintenance Notes
- Email system is independent of form submission - failures won't affect data saving
- All email operations are logged for debugging
- HTML email template is inline-styled for maximum compatibility
- System automatically handles both new submissions and updates

## Future Enhancements
Potential improvements could include:
- Email preferences/opt-out functionality
- Additional recipients (supervisors, team leads)
- Email templates for different evaluation types
- Attachment of evaluation summaries
- Integration with external email systems

## Troubleshooting
- Check Google Apps Script execution logs for email errors
- Verify user has valid email address in Google account
- Test with `testEmailNotification()` function
- Ensure MailApp service has proper permissions
