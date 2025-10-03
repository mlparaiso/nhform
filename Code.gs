/**
 * @OnlyCurrentDoc
 */

// --- CONFIGURATION ---
const SPREADSHEET_ID = "1InCv7-d8-KhovEGb8tdI3TxLJx3PTou5ZJKsx1nZhwI";
const RESPONSE_SHEET_NAME = "QA NH Form Responses";
const ROSTER_SPREADSHEET_ID = "1-UIO2kTsIdx5oru_uvg0q9rtNM-1Hb5e_R_YgISoSh4";
const ROSTER_SHEET_NAME = "roster";

// External dropdown configuration
const DROPDOWN_SPREADSHEET_ID = "1jZPduJvJfGjgvqt7K869osKzLb9zAKeszP7FmJCA8H0";
const DROPDOWN_SHEET_NAME = "Dropdowns";

// Enhanced data structure configuration
const FORM_VERSION = "2.1";

// --- ENHANCED DATA STRUCTURE FUNCTIONS ---

/**
 * Generates a unique Record ID in format: NH-mmddyyyy-8char
 * @returns {string} Unique record ID
 */
function generateRecordID() {
  const now = new Date();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const day = String(now.getDate()).padStart(2, '0');
  const year = now.getFullYear();
  const dateStr = `${month}${day}${year}`;
  
  // Generate 8-character nanoid
  const alphabet = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz';
  let nanoid = '';
  for (let i = 0; i < 8; i++) {
    nanoid += alphabet.charAt(Math.floor(Math.random() * alphabet.length));
  }
  
  return `NH-${dateStr}-${nanoid}`;
}

/**
 * Creates enhanced data structure with metadata and organization
 * @param {Object} formData - Raw form data from submission
 * @returns {Object} Enhanced data structure with metadata
 */
function createEnhancedFormData(formData) {
  const recordId = generateRecordID();
  const timestamp = new Date();
  
  return {
    // System metadata
    recordId: recordId,
    timestamp: timestamp,
    formVersion: FORM_VERSION,
    submittedBy: Session.getActiveUser().getEmail(),
    
    // Team member details
    teamMemberDetails: {
      sapId: formData["Sap ID"] || "",
      teamMember: formData["Team Member"] || "",
      teamLeader: formData["Team Leader"] || "",
      operationsManager: formData["Operations Manager"] || "",
      lineOfBusiness: formData["Line of Business"] || ""
    },
    
    // Call details
    callDetails: {
      interactionDate: formData["Interaction Date"] || "",
      customerBAN: formData["Customer's BAN"] || "",
      flpConversationId: formData["FLP Conversation ID"] || "",
      listeningType: formData["Listening Type"] || "",
      callDuration: formData["Duration (minutes)"] || "",
      callDriverLevel1: formData["Call Driver Level 1"] || "",
      callDriverLevel2: formData["Call Driver Level 2"] || ""
    },
    
    // Skill evaluations organized by category
    skillEvaluations: {
      coreSkills: {
        willingness: formData["Willingness to Help"] || "",
        greeting: formData["Warm and Energetic Greeting that Connects with the Customer"] || "",
        listenRetain: formData["Listen & Retain"] || ""
      },
      criticalSkills: {
        secureIdentity: formData["Secure the Customer's Identity Based on Department Guidelines"] || "",
        probing: formData["Probe Using a Combination of Open-ended and Closed-ended Questions"] || "",
        investigate: formData["Investigate & Check All Applicable Resources"] || "",
        offerSolutions: formData["Offer Solutions / Explanations / Options"] || "",
        processTransaction: formData["Begin Processing Transaction, Adding Notes and Updating Account Information when Appropriate"] || ""
      },
      basicSkills: {
        confirmParaphrase: formData["Confirm What You Know and Paraphrase Customer"] || "",
        checkUnderstanding: formData["Check for Understanding"] || "",
        leadWithBenefit: formData["Lead with a Benefit Prior to Asking Questions and Securing the Account"] || "",
        summarizeSolution: formData["Summarize Solution"] || "",
        personalizedClose: formData["End on a High Note with a Personalized Close"] || "",
        reinforceExtras: formData["Reinforce the Extras You Did"] || ""
      },
      intermediateSkills: {
        earlyBridge: formData["Early Bridge to Value When Appropriate"] || "",
        laterBridge: formData["Later Bridge to Value when Appropriate / Value Proposition"] || "",
        leadToAccept: formData["Lead Customer to Accept an Option or Solution"] || "",
        checkSatisfaction: formData["Check for Satisfaction"] || ""
      },
      advancedSkills: {
        overcomeobjections: formData["Overcome Objections if Necessary"] || "",
        empathy: formData["Acknowledge (Empathy When Needed)"] || ""
      },
      salesSkills: {
        learnFromCustomer: formData["Learn from Your Customer"] || "",
        handleEarlyObjections: formData["Handling Early Objections"] || "",
        customerIntimacy: formData["Customer Intimacy"] || "",
        valueProposition: formData["Value Proposition"] || "",
        bridgingType: formData["What type of bridging was provided?"] || ""
      },
      clsSkills: {
        rightSize: formData["Right size"] || "",
        tieredApproach: formData["Use a tiered approach"] || "",
        sufficientProbing: formData["Sufficient probing (+3 Questions)"] || "",
        speakToValue: formData["Speak to value (WHY TELUS, reality check and service superiority)"] || "",
        loyaltyStatements: formData["Loyalty statements implementation"] || "",
        investigation: formData["Investigation (Review bill, compare competitor offer, search notes on the account, take a look at usage)"] || "",
        retainCustomer: formData["Was the agent able to retain the customer?"] || ""
      }
    },
    
    // CLS specific data
    clsData: {
      cancellationRootCause: formData["Cancellation Root Cause"] || "",
      unethicalCredit: formData["Was there an unethical credit?"] || ""
    },
    
    // Evaluation content
    evaluationContent: {
      callSummary: formData["Call Summary"] || "",
      highlights: formData["Highlights"] || "",
      lowlights: formData["Lowlights/Recommendation"] || ""
    },
    
    // Critical behaviors data
    criticalBehaviors: {
      abusiveInsult: formData["Abusive, insult the customer (disrespectful)"] === "on" ? 1 : 0,
      appropriationSale: formData["Appropriation of the sale from another agent"] === "on" ? 1 : 0,
      cctsProcessNotFollowed: formData["CCTS Process not Followed"] === "on" ? 1 : 0,
      delayedResponse: formData["Delayed Response to Inbound Contact"] === "on" ? 1 : 0,
      didntCompleteTransaction: formData["Didn't complete transaction as promised"] === "on" ? 1 : 0,
      excessiveDeadAir: formData["Excessive Dead Air, Hold or Idle Time (with the clear intention of making the customer disconnect the call)"] === "on" ? 1 : 0,
      failureFollowProcesses: formData["Failure to follow processes as per Onesource or use tools properly"] === "on" ? 1 : 0,
      incorrectSalesOrder: formData["Incorrect sales order processing"] === "on" ? 1 : 0,
      incorrectCredits: formData["Incorrect use of credits/discounts"] === "on" ? 1 : 0,
      misleadingInformation: formData["Misleading Information / Deliberately lying"] === "on" ? 1 : 0,
      misleadingProduct: formData["Misleading product/offer information"] === "on" ? 1 : 0,
      missedCallBack: formData["Missed Call Back"] === "on" ? 1 : 0,
      modifyAccountWithoutConsent: formData["Modify customer's account without consent"] === "on" ? 1 : 0,
      noAuthentication: formData["No Authentication/Verification"] === "on" ? 1 : 0,
      noResponseInbound: formData["No Response to Inbound Contact"] === "on" ? 1 : 0,
      noRetentionAttempt: formData["No Retention Attempt or No Value Generation effort"] === "on" ? 1 : 0,
      nonworkDistractions: formData["Nonwork related distractions"] === "on" ? 1 : 0,
      overtalking: formData["Overtalking to the customer, interrupt the customer"] === "on" ? 1 : 0,
      proactivelyTransferring: formData["Proactively transferring calls"] === "on" ? 1 : 0,
      stayingLineUnnecessarily: formData["Staying on the line unnecessarily"] === "on" ? 1 : 0,
      systemToolsManipulation: formData["System/Tools Manipulation"] === "on" ? 1 : 0,
      unjustifiedDisconnection: formData["Unjustified disconnection of customer interaction"] === "on" ? 1 : 0,
      unnecessaryConsult: formData["Unnecessary/Invalid Consult, Dispatch or Transfer"] === "on" ? 1 : 0,
      unwillingnessToHelp: formData["Unwillingness to Help"] === "on" ? 1 : 0,
      otherBehavior: formData["Other"] === "on" ? 1 : 0,
      otherBehaviorText: formData["CB_Other_Behavior_Text"] || ""
    },
    
    // Data integrity metadata
    dataIntegrity: {
      checksum: calculateChecksum(formData),
      fieldCount: Object.keys(formData).length,
      validationStatus: "pending",
      missingFields: [],
      dataQualityScore: 0
    }
  };
}

/**
 * Calculates a simple checksum for data integrity
 * @param {Object} data - Data to calculate checksum for
 * @returns {string} Hexadecimal checksum
 */
function calculateChecksum(data) {
  const dataString = JSON.stringify(data);
  let hash = 0;
  for (let i = 0; i < dataString.length; i++) {
    const char = dataString.charCodeAt(i);
    hash = ((hash << 5) - hash) + char;
    hash = hash & hash; // Convert to 32-bit integer
  }
  return Math.abs(hash).toString(16);
}

/**
 * Validates form data and calculates quality score
 * @param {Object} enhancedData - Enhanced data structure to validate
 * @returns {Object} Validation results with errors and quality score
 */
function validateEnhancedData(enhancedData) {
  const errors = [];
  const missingFields = [];
  
  // Required field validation
  if (!enhancedData.teamMemberDetails.sapId) {
    errors.push("SAP ID is required");
    missingFields.push("Sap ID");
  }
  
  if (!enhancedData.teamMemberDetails.teamMember) {
    errors.push("Team Member is required");
    missingFields.push("Team Member");
  }
  
  if (!enhancedData.callDetails.interactionDate) {
    errors.push("Interaction Date is required");
    missingFields.push("Interaction Date");
  }
  
  // Calculate data quality score (0-100)
  const totalFields = enhancedData.dataIntegrity.fieldCount;
  const filledFields = totalFields - missingFields.length;
  const qualityScore = Math.round((filledFields / totalFields) * 100);
  
  // Update data integrity
  enhancedData.dataIntegrity.validationStatus = errors.length === 0 ? "passed" : "failed";
  enhancedData.dataIntegrity.missingFields = missingFields;
  enhancedData.dataIntegrity.dataQualityScore = qualityScore;
  
  return {
    isValid: errors.length === 0,
    errors: errors,
    qualityScore: qualityScore
  };
}

/**
 * Prepares row data maintaining backward compatibility
 * @param {Object} enhancedData - Enhanced data structure
 * @returns {Array} Array of values for spreadsheet row
 */
function prepareRowData(enhancedData) {
  return [
    // System metadata
    enhancedData.recordId,
    enhancedData.timestamp,
    enhancedData.formVersion,
    enhancedData.submittedBy,
    
    // Team member details
    enhancedData.teamMemberDetails.sapId,
    enhancedData.teamMemberDetails.teamMember,
    enhancedData.teamMemberDetails.teamLeader,
    enhancedData.teamMemberDetails.operationsManager,
    enhancedData.teamMemberDetails.lineOfBusiness,
    
    // Call details
    enhancedData.callDetails.interactionDate,
    enhancedData.callDetails.customerBAN,
    enhancedData.callDetails.flpConversationId,
    enhancedData.callDetails.listeningType,
    enhancedData.callDetails.callDuration,
    
    // Skills (maintaining original order for backward compatibility)
    enhancedData.skillEvaluations.coreSkills.willingness,
    enhancedData.skillEvaluations.coreSkills.greeting,
    enhancedData.skillEvaluations.coreSkills.listenRetain,
    enhancedData.skillEvaluations.criticalSkills.secureIdentity,
    enhancedData.skillEvaluations.criticalSkills.probing,
    enhancedData.skillEvaluations.criticalSkills.investigate,
    enhancedData.skillEvaluations.criticalSkills.offerSolutions,
    enhancedData.skillEvaluations.criticalSkills.processTransaction,
    enhancedData.skillEvaluations.basicSkills.confirmParaphrase,
    enhancedData.skillEvaluations.basicSkills.checkUnderstanding,
    enhancedData.skillEvaluations.basicSkills.leadWithBenefit,
    enhancedData.skillEvaluations.basicSkills.summarizeSolution,
    enhancedData.skillEvaluations.basicSkills.personalizedClose,
    enhancedData.skillEvaluations.basicSkills.reinforceExtras,
    enhancedData.skillEvaluations.intermediateSkills.earlyBridge,
    enhancedData.skillEvaluations.intermediateSkills.laterBridge,
    enhancedData.skillEvaluations.intermediateSkills.leadToAccept,
    enhancedData.skillEvaluations.intermediateSkills.checkSatisfaction,
    enhancedData.skillEvaluations.advancedSkills.overcomeobjections,
    enhancedData.skillEvaluations.advancedSkills.empathy,
    enhancedData.skillEvaluations.salesSkills.learnFromCustomer,
    enhancedData.skillEvaluations.salesSkills.handleEarlyObjections,
    enhancedData.skillEvaluations.salesSkills.customerIntimacy,
    enhancedData.skillEvaluations.salesSkills.valueProposition,
    enhancedData.skillEvaluations.salesSkills.bridgingType,
    enhancedData.clsData.cancellationRootCause,
    enhancedData.clsData.unethicalCredit,
    enhancedData.skillEvaluations.clsSkills.rightSize,
    enhancedData.skillEvaluations.clsSkills.tieredApproach,
    enhancedData.skillEvaluations.clsSkills.sufficientProbing,
    enhancedData.skillEvaluations.clsSkills.speakToValue,
    enhancedData.skillEvaluations.clsSkills.loyaltyStatements,
    enhancedData.skillEvaluations.clsSkills.investigation,
    enhancedData.skillEvaluations.clsSkills.retainCustomer,
    enhancedData.callDetails.callDriverLevel1,
    enhancedData.callDetails.callDriverLevel2,
    enhancedData.evaluationContent.callSummary,
    enhancedData.evaluationContent.highlights,
    enhancedData.evaluationContent.lowlights,
    
    // Critical Behaviors boolean values
    enhancedData.criticalBehaviors.abusiveInsult,
    enhancedData.criticalBehaviors.appropriationSale,
    enhancedData.criticalBehaviors.cctsProcessNotFollowed,
    enhancedData.criticalBehaviors.delayedResponse,
    enhancedData.criticalBehaviors.didntCompleteTransaction,
    enhancedData.criticalBehaviors.excessiveDeadAir,
    enhancedData.criticalBehaviors.failureFollowProcesses,
    enhancedData.criticalBehaviors.incorrectSalesOrder,
    enhancedData.criticalBehaviors.incorrectCredits,
    enhancedData.criticalBehaviors.misleadingInformation,
    enhancedData.criticalBehaviors.misleadingProduct,
    enhancedData.criticalBehaviors.missedCallBack,
    enhancedData.criticalBehaviors.modifyAccountWithoutConsent,
    enhancedData.criticalBehaviors.noAuthentication,
    enhancedData.criticalBehaviors.noResponseInbound,
    enhancedData.criticalBehaviors.noRetentionAttempt,
    enhancedData.criticalBehaviors.nonworkDistractions,
    enhancedData.criticalBehaviors.overtalking,
    enhancedData.criticalBehaviors.proactivelyTransferring,
    enhancedData.criticalBehaviors.stayingLineUnnecessarily,
    enhancedData.criticalBehaviors.systemToolsManipulation,
    enhancedData.criticalBehaviors.unjustifiedDisconnection,
    enhancedData.criticalBehaviors.unnecessaryConsult,
    enhancedData.criticalBehaviors.unwillingnessToHelp,
    enhancedData.criticalBehaviors.otherBehavior,
    enhancedData.criticalBehaviors.otherBehaviorText,
    
    // Data integrity fields
    enhancedData.dataIntegrity.checksum,
    enhancedData.dataIntegrity.fieldCount,
    enhancedData.dataIntegrity.validationStatus,
    enhancedData.dataIntegrity.dataQualityScore
  ];
}

/**
 * Sets up enhanced headers for the response sheet
 * @param {Sheet} sheet - Google Sheets sheet object
 */
function setupEnhancedHeaders(sheet) {
  const headers = [
    // System metadata
    "RecordID", "Timestamp", "FormVersion", "SubmittedBy",
    
    // Original headers for backward compatibility
    "Sap ID", "Team Member", "Team Leader", "Operations Manager", "Line of Business",
    "Interaction Date", "Customer's BAN", "FLP Conversation ID", "Listening Type", "Call Duration",
    
    // All skill fields in original order
    "Willingness to Help", "Warm and Energetic Greeting that Connects with the Customer", "Listen & Retain",
    "Secure the Customer's Identity Based on Department Guidelines", "Probe Using a Combination of Open-ended and Closed-ended Questions", "Investigate & Check All Applicable Resources", "Offer Solutions / Explanations / Options", "Begin Processing Transaction, Adding Notes and Updating Account Information when Appropriate",
    "Confirm What You Know and Paraphrase Customer", "Check for Understanding", "Lead with a Benefit Prior to Asking Questions and Securing the Account", "Summarize Solution", "End on a High Note with a Personalized Close", "Reinforce the Extras You Did",
    "Early Bridge to Value When Appropriate", "Later Bridge to Value when Appropriate / Value Proposition", "Lead Customer to Accept an Option or Solution", "Check for Satisfaction",
    "Overcome Objections if Necessary", "Acknowledge (Empathy When Needed)",
    "Learn from Your Customer", "Handling Early Objections", "Customer Intimacy", "Value Proposition", "What type of bridging was provided?",
    "Cancellation Root Cause", "Was there an unethical credit?",
    "Right size", "Use a tiered approach", "Sufficient probing (+3 Questions)", "Speak to value (WHY TELUS, reality check and service superiority)", "Loyalty statements implementation", "Investigation (Review bill, compare competitor offer, search notes on the account, take a look at usage)", "Was the agent able to retain the customer?",
    "Call Driver Level 1", "Call Driver Level 2",
    "Call Summary", "Highlights", "Lowlights/Recommendation",
    
    // Critical Behaviors boolean columns
    "CB_Abusive_Insult_Customer", "CB_Appropriation_Sale_Another_Agent", "CB_CCTS_Process_Not_Followed", 
    "CB_Delayed_Response_Inbound", "CB_Didnt_Complete_Transaction_Promised", "CB_Excessive_Dead_Air_Hold_Idle",
    "CB_Failure_Follow_Processes_Onesource", "CB_Incorrect_Sales_Order_Processing", "CB_Incorrect_Use_Credits_Discounts",
    "CB_Misleading_Information_Lying", "CB_Misleading_Product_Offer_Info", "CB_Missed_Call_Back",
    "CB_Modify_Account_Without_Consent", "CB_No_Authentication_Verification", "CB_No_Response_Inbound_Contact",
    "CB_No_Retention_Value_Generation", "CB_Nonwork_Related_Distractions", "CB_Overtalking_Interrupt_Customer",
    "CB_Proactively_Transferring_Calls", "CB_Staying_Line_Unnecessarily", "CB_System_Tools_Manipulation",
    "CB_Unjustified_Disconnection", "CB_Unnecessary_Invalid_Consult_Transfer", "CB_Unwillingness_To_Help",
    "CB_Other_Behavior", "CB_Other_Behavior_Text",
    
    // Data integrity fields
    "DataChecksum", "FieldCount", "ValidationStatus", "DataQualityScore"
  ];
  
  sheet.appendRow(headers);
  sheet.getRange("1:1").setFontWeight("bold");
  sheet.setFrozenRows(1);
}

// Prompts for the rewriteText function
const systemPrompts = {
  'callSummary': "You are a professional quality assurance analyst. Create a detailed call summary in paragraph format with important details such as offers, pricing, reasons, actions from agent and customer, resolution, etc. Address the customer as 'Customer' at first mention, then 'Cx' in subsequent references. Address the agent as 'TM' throughout. Write the summary as flowing paragraphs without headers, bullet points, or structured formatting. Include specific details about what was discussed, any offers made, pricing mentioned, customer concerns, agent actions, and final resolution in a narrative paragraph style. Crucially, you must only use the information provided in the input text. Do not add any new details, outcomes, or actions that are not explicitly mentioned. If the input is empty or meaningless, return it unchanged.",
  'highlights': "You are a call monitoring specialist focused on identifying positive performance in 3 core skills: 1) Reducing Repeat Calls (probing, resolution skills), 2) Generating Revenue (bridging, sales skills), and 3) Customer Retention skills. Polish the user's input notes into professional, encouraging sentences. Each distinct point or line from the user should be converted into its own bullet point starting with a hyphen (-). Provide maximum 3 highlights, each on a separate line with a '-' prefix. Focus specifically on these three core skill areas and how the TM demonstrated them effectively. Stick closely to the concepts provided by the user; do not add new, unrelated points or expand one idea into multiple bullets. If the input is empty or meaningless, return it unchanged.",
  'lowlights': "You are a QA analyst providing constructive feedback focused on 3 core skills: 1) Reducing Repeat Calls (probing, resolution skills), 2) Generating Revenue (bridging, sales skills), and 3) Customer Retention skills. Polish the user's input notes into professional, objective recommendations. Each distinct point or line from the user should be converted into its own bullet point starting with a hyphen (-). Refer to the agent as 'TM' or 'the TM'. Provide maximum 3 recommendations, each on a separate line with a '-' prefix. Generate suggested statements that are connected to the customer concern but still embody the Customer Experience Blueprint skills concepts. Focus on actionable improvements in these three core areas. Stick closely to the concepts provided by the user; do not add new, unrelated points or expand one idea into multiple bullets. If the input is empty or meaningless, return it unchanged.",
};

/**
 * Gets the current user's email address
 * @returns {string} Current user's email
 */
function getCurrentUserEmail() {
  try {
    return Session.getActiveUser().getEmail();
  } catch (error) {
    Logger.log(`Error getting user email: ${error.message}`);
    return 'user@company.com'; // Fallback for testing
  }
}

/**
 * Checks if the current user is an admin user
 * @returns {boolean} True if user is admin, false otherwise
 */
function isAdminUser() {
  try {
    const currentUserEmail = Session.getActiveUser().getEmail();
    
    Logger.log(`Checking admin status for: ${currentUserEmail}`);
    
    // Only michael.paraiso@telus.com has admin access
    const isAdmin = currentUserEmail === 'michael.paraiso@telus.com';
    
    Logger.log(`Admin status: ${isAdmin}`);
    
    return isAdmin;
  } catch (error) {
    Logger.log(`Error checking admin status: ${error.message}`);
    return false; // Default to non-admin if error occurs
  }
}

/**
 * Serves the main HTML page of the web app.
 */
function doGet(e) {
  const html = HtmlService.createTemplateFromFile('index').evaluate();
  html.setTitle('New Hire Evaluation Form');
  html.addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  return html;
}

/**
 * Fetches dropdown options from the external configuration spreadsheet.
 * @returns {Object} An object containing arrays of options for each dropdown field.
 */
function getDropdownOptions() {
  try {
    const ss = SpreadsheetApp.openById(DROPDOWN_SPREADSHEET_ID);
    const sheet = ss.getSheetByName(DROPDOWN_SHEET_NAME);
    if (!sheet) {
      throw new Error(`Dropdown sheet "${DROPDOWN_SHEET_NAME}" not found.`);
    }
    
    const data = sheet.getDataRange().getValues();
    
    // Skip the first row (headers) and extract options from each column, filtering out empty cells
    const dataRows = data.slice(1); // Skip header row
    const listeningTypeOptions = dataRows.map(row => row[0]).filter(cell => cell && cell.toString().trim() !== '');
    const bridgingTypeOptions = dataRows.map(row => row[1]).filter(cell => cell && cell.toString().trim() !== '');
    const cancellationRootCauseOptions = dataRows.map(row => row[2]).filter(cell => cell && cell.toString().trim() !== '');
    const unethicalCreditOptions = dataRows.map(row => row[3]).filter(cell => cell && cell.toString().trim() !== '');
    
    // Process Call Driver data for cascading dropdowns
    const callDriverData = dataRows
      .filter(row => row[4] && row[4].toString().trim() !== '' && row[5] && row[5].toString().trim() !== '')
      .map(row => ({
        level1: row[4].toString().trim(),
        level2: row[5].toString().trim()
      }));
    
    // Create unique Level 1 options
    const callDriverLevel1Options = [...new Set(callDriverData.map(item => item.level1))];
    
    // Create mapping for cascading functionality
    const callDriverMapping = {};
    callDriverData.forEach(item => {
      if (!callDriverMapping[item.level1]) {
        callDriverMapping[item.level1] = [];
      }
      if (!callDriverMapping[item.level1].includes(item.level2)) {
        callDriverMapping[item.level1].push(item.level2);
      }
    });
    
    return {
      "Listening Type": listeningTypeOptions,
      "What type of bridging was provided?": bridgingTypeOptions,
      "Cancellation Root Cause": cancellationRootCauseOptions,
      "Was there an unethical credit?": unethicalCreditOptions,
      "Call Driver Level 1": callDriverLevel1Options,
      "Call Driver Mapping": callDriverMapping
    };
  } catch (error) {
    Logger.log(`Error fetching dropdown options: ${error.message}`);
    // Return fallback hardcoded options if external sheet fails
    return {
      "Listening Type": ["Live/Side by Side", "Recorded"],
      "What type of bridging was provided?": [
        "Open (Exploring customer's interests)",
        "Closed (Specific offer towards a product)",
        "No bridging",
        "Not applicable – Solutioning not within agent skill / role, agent not trained",
        "Not applicable – interaction not adequate for solutioning"
      ],
      "Cancellation Root Cause": [
        "Usage - No Longer Needs Home Security Service",
        "Usage - Deceased Customer",
        "Move - Landlord Policies",
        "Move - No TELUS Coverage",
        "Move - Moving Outside Canada",
        "Move - Moving to Secured Building",
        "Move - Existing Services on New Property",
        "Move - Moving to a Care Facility",
        "Product/Equipment - Not Working",
        "Product/Equipment - Lack of Features",
        "Product/Equipment - Not Compatible",
        "Pricing - Too Expensive Price Increase",
        "Pricing - Too Expensive Migration",
        "Pricing - Monthly Charge",
        "TELUS Experience - Technician Did Not Arrived",
        "TELUS Experience - Multiple Transfers",
        "TELUS Experience - Issue Not Solved",
        "Installation - Wiring",
        "Installation - Incomplete Installation",
        "Installation - DIY Installation",
        "Installation - No Technicians Available",
        "Unhappy with procedures - ADT migration to SHS"
      ],
      "Was there an unethical credit?": [
        "No",
        "Yes - Ineligible for offer",
        "Yes - Incorrect discount/offer",
        "Yes - Excessive discount",
        "Yes - Proactively offered credit"
      ]
    };
  }
}

/**
 * Diagnostic function to check roster sheet structure
 * @returns {Object} Roster sheet diagnostic information
 */
function debugRosterSheet() {
  try {
    const ss = SpreadsheetApp.openById(ROSTER_SPREADSHEET_ID);
    const sheet = ss.getSheetByName(ROSTER_SHEET_NAME);
    
    if (!sheet) {
      return {
        error: `Roster sheet "${ROSTER_SHEET_NAME}" not found`,
        availableSheets: ss.getSheets().map(sheet => sheet.getName())
      };
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const totalRows = data.length;
    
    // Get sample data from first few rows
    const sampleData = [];
    for (let i = 1; i < Math.min(4, totalRows); i++) {
      if (data[i]) {
        const rowData = {};
        headers.forEach((header, index) => {
          rowData[header] = data[i][index];
        });
        sampleData.push(rowData);
      }
    }
    
    return {
      sheetExists: true,
      totalRows: totalRows - 1, // Exclude header
      totalColumns: headers.length,
      headers: headers,
      sampleData: sampleData,
      hasData: totalRows > 1
    };
    
  } catch (error) {
    return {
      error: error.message,
      details: error.toString()
    };
  }
}

/**
 * Gets team member details from the roster sheet based on SAP ID.
 */
function getTeamMemberDetails(sapId) {
  if (!sapId) { 
    Logger.log("getTeamMemberDetails: No SAP ID provided");
    return null; 
  }
  
  try {
    Logger.log(`=== ROSTER LOOKUP DEBUG START ===`);
    Logger.log(`Looking up SAP ID: "${sapId}"`);
    
    const ss = SpreadsheetApp.openById(ROSTER_SPREADSHEET_ID);
    const sheet = ss.getSheetByName(ROSTER_SHEET_NAME);
    
    if (!sheet) { 
      Logger.log(`Roster sheet "${ROSTER_SHEET_NAME}" not found`);
      throw new Error(`Roster sheet "${ROSTER_SHEET_NAME}" not found.`); 
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0]; // Get headers from first row
    const dataRows = data.slice(1); // Remove header row
    
    Logger.log(`Roster sheet has ${dataRows.length} data rows and ${headers.length} columns`);
    Logger.log(`Headers: ${headers.join(', ')}`);
    
    // Use fixed column positions based on roster sheet structure
    // This is more reliable than dynamic header detection
    const sapIdIndex = 0;        // Column A: Employee ID
    const teamMemberIndex = 2;   // Column C: Full Name
    const teamLeaderIndex = 8;   // Column I: Direct Supervisor
    const opsManagerIndex = 10;  // Column K: Ops Mgr
    const lobIndex = 3;          // Column D: LOB
    
    Logger.log(`Column indices found:`);
    Logger.log(`SAP ID: ${sapIdIndex} (${headers[sapIdIndex]})`);
    Logger.log(`Team Member: ${teamMemberIndex} (${headers[teamMemberIndex]})`);
    Logger.log(`Team Leader: ${teamLeaderIndex} (${headers[teamLeaderIndex]})`);
    Logger.log(`Operations Manager: ${opsManagerIndex} (${headers[opsManagerIndex]})`);
    Logger.log(`Line of Business: ${lobIndex} (${headers[lobIndex]})`);
    
    if (sapIdIndex === -1) {
      Logger.log("❌ SAP ID column not found in roster sheet");
      throw new Error("SAP ID column not found in roster sheet. Please check the roster sheet structure.");
    }
    
    // Search for the SAP ID in the data
    for (let i = 0; i < dataRows.length; i++) {
      const row = dataRows[i];
      const currentSapId = row[sapIdIndex];
      
      Logger.log(`Row ${i + 2}: SAP ID = "${currentSapId}"`);
      
      if (currentSapId && currentSapId.toString().trim() === sapId.toString().trim()) {
        Logger.log(`✅ MATCH FOUND at row ${i + 2}!`);
        
        const result = {
          "Sap ID": currentSapId,
          "Team Member": teamMemberIndex !== -1 ? (row[teamMemberIndex] || '') : '',
          "Team Leader": teamLeaderIndex !== -1 ? (row[teamLeaderIndex] || '') : '',
          "Operations Manager": opsManagerIndex !== -1 ? (row[opsManagerIndex] || '') : '',
          "Line of Business": lobIndex !== -1 ? (row[lobIndex] || '') : ''
        };
        
        Logger.log(`Returning result:`, JSON.stringify(result, null, 2));
        Logger.log(`=== ROSTER LOOKUP DEBUG END - SUCCESS ===`);
        
        return result;
      }
    }
    
    Logger.log(`❌ No matching SAP ID found for: "${sapId}"`);
    Logger.log(`=== ROSTER LOOKUP DEBUG END - NOT FOUND ===`);
    return null;
    
  } catch (error) {
    Logger.log(`❌ Error in getTeamMemberDetails: ${error.message}`);
    Logger.log(`Error stack: ${error.stack}`);
    Logger.log(`=== ROSTER LOOKUP DEBUG END - ERROR ===`);
    throw new Error(`Could not retrieve details. Reason: ${error.message}`);
  }
}

/**
 * Rewrites text using the AI model.
 */
function rewriteText(text, type) {
  // Function is unchanged
  if (!text || text.trim().length === 0) { return text; }
  if (!type || !systemPrompts[type]) { throw new Error("Invalid rewrite type specified."); }
  const systemPrompt = systemPrompts[type];
  const payload = { model: "gpt-4o", messages: [{ role: "system", content: systemPrompt }, { role: "user", content: text }], temperature: 0.5 };
  const options = { method: "post", contentType: "application/json", headers: { "Authorization": "Bearer ak-fnRymjzGJpyatnpMnnaAW4ShBfQI" }, payload: JSON.stringify(payload), muteHttpExceptions: true };
  try {
    const response = UrlFetchApp.fetch("https://api-beta.fuelix.ai/v1/chat/completions", options);
    const responseCode = response.getResponseCode();
    if (responseCode !== 200) { throw new Error(`Rewrite API request failed with status ${responseCode}.`); }
    const jsonResponse = JSON.parse(response.getContentText());
    return jsonResponse.choices[0].message.content.trim();
  } catch (error) {
    Logger.log(`Rewrite failed: ${error.toString()}`);
    throw new Error(`Rewrite failed. Details: ${error.message}`);
  }
}

/**
 * Builds a detailed rubric string for the AI prompt.
 * @returns {string} A formatted string of the entire rubric.
 */
function buildAiRubricPrompt() {
  const formStructure = getFormStructure();
  const rubricDetails = getRubricDetails();
  const dropdownOptions = getDropdownOptions();
  let rubricString = "";

  formStructure.forEach(section => {
    if (section.subsections) {
      for (const subTitle in section.subsections) {
        rubricString += `\n### Subsection: ${subTitle}\n`;
        const skills = section.subsections[subTitle];
        skills.forEach(skill => {
          const details = rubricDetails[skill.name];
          if (details) {
            rubricString += `\n**Skill:** ${skill.name}\n`;
            rubricString += `*Description:* ${details.description}\n`;
            rubricString += `*What to Observe:*\n${details.observe.map(item => `- ${item}`).join('\n')}\n`;
            rubricString += `*Rating Criteria:*\n`;
            for(const rating in details.ratings) {
              rubricString += `- **${rating}**: ${details.ratings[rating]}\n`;
            }
          } else if (skill.type === 'dropdown') {
            // Handle special cases like the dropdown
            rubricString += `\n**Skill:** ${skill.name}\n`;
            rubricString += `*Instruction:* For this skill ONLY, the "rating" MUST be one of the following exact strings: "${skill.options.join('", "')}".\n`;
          }
        });
      }
    }
  });
  
  // Add Call Driver dropdown options to the rubric
  if (dropdownOptions["Call Driver Level 1"] && dropdownOptions["Call Driver Level 1"].length > 0) {
    rubricString += `\n### Call Driver Analysis\n`;
    rubricString += `\n**Call Driver Level 1 Options:**\n`;
    rubricString += dropdownOptions["Call Driver Level 1"].map(option => `- ${option}`).join('\n') + '\n';
    rubricString += `\n**Instructions for Call Driver Analysis:**\n`;
    rubricString += `- Analyze the call script to determine the primary reason/driver for the customer's call\n`;
    rubricString += `- Select the most appropriate Call Driver Level 1 from the options above\n`;
    rubricString += `- The Call Driver Level 2 will be determined based on the Level 1 selection\n`;
  }
  
  return rubricString;
}

/**
 * Analyzes a call script using the AI model and returns a structured evaluation.
 */
function analyzeCallScript(scriptText) {
  if (!scriptText || scriptText.trim().length === 0) {
    throw new Error("Script text cannot be empty.");
  }

  const detailedRubric = buildAiRubricPrompt();

  const systemPrompt = `
    You are an expert QA Analyst. Your task is to analyze the following call script between a Customer and a Team Member (TM). Based on the script, you must evaluate the TM's performance against the detailed rubric provided below. In all of your written output, you must refer to the agent as "TM" and the customer as "Customer".

    You MUST return your analysis as a single, valid JSON object. Do not include any text outside the JSON object. The JSON object must contain the following top-level keys: "skillEvaluations", "callSummary", "highlights", "lowlights", "callDriverLevel1", and "callDriverLevel2".

    The "skillEvaluations" key must be an object where each key is a subsection title from the rubric (e.g., "Core", "Critical", "CLS Skills Evaluation") and the value is an array of objects. Each object in the array represents a skill and must contain three keys: "skillName" (string), "rating" (string), and "remarks" (string, a brief justification for the rating based on the script). For skills with dropdowns or Yes/No/NA ratings, ensure the "rating" value is one of the provided options.

    IMPORTANT SKILL RATING RULES:
    1. DO NOT provide ratings for these skills as they require manual verification: "Secure the Customer's Identity Based on Department Guidelines", "Investigate & Check All Applicable Resources", "Offer Solutions / Explanations / Options", "Begin Processing Transaction, Adding Notes and Updating Account Information when Appropriate". For these skills, set rating to "MANUAL" and remarks to "Requires manual verification".

    2. For Sales Skills bridging logic: If the bridging dropdown would be "No bridging", "Not applicable – Solutioning not within agent skill / role, agent not trained", or "Not applicable – interaction not adequate for solutioning", then set ALL Sales Skills to "NA" EXCEPT "Learn from Your Customer" which can still be evaluated since it can be assessed even without bridging attempts.

    For the "callSummary" field: Create a detailed call summary in paragraph format with important details such as offers, pricing, reasons, actions from agent and customer, resolution, etc. Address the customer as 'Customer' at first mention, then 'Cx' in subsequent references. Address the agent as 'TM' throughout. Write the summary as flowing paragraphs without headers, bullet points, or structured formatting. Include specific details about what was discussed, any offers made, pricing mentioned, customer concerns, agent actions, and final resolution in a narrative paragraph style. Crucially, you must only use the information provided in the call script. Do not add any new details, outcomes, or actions that are not explicitly mentioned.

    For the "highlights" field: Identify positive performance in 3 core skills: 1) Reducing Repeat Calls (probing, resolution skills), 2) Generating Revenue (bridging, sales skills), and 3) Customer Retention skills. Polish positive observations from the call into professional, encouraging sentences. Each distinct positive point should be converted into its own bullet point starting with a hyphen (-). Provide maximum 3 highlights, each on a separate line with a '-' prefix. Focus specifically on these three core skill areas and how the TM demonstrated them effectively. Stick closely to what was observed in the call; do not add new, unrelated points or expand one idea into multiple bullets.

    For the "lowlights" field: Provide constructive feedback focused on 3 core skills: 1) Reducing Repeat Calls (probing, resolution skills), 2) Generating Revenue (bridging, sales skills), and 3) Customer Retention skills. Polish constructive feedback into professional, objective recommendations. Each distinct point should be converted into its own bullet point starting with a hyphen (-). Refer to the agent as 'TM' or 'the TM'. Provide maximum 3 recommendations, each on a separate line with a '-' prefix. Generate suggested statements that are connected to the customer concern but still embody the Customer Experience Blueprint skills concepts. Focus on actionable improvements in these three core areas. Stick closely to the observations from the call; do not add new, unrelated points or expand one idea into multiple bullets.

    For the "callDriverLevel1" field: Analyze the call script to determine the primary reason/driver for the customer's call. Select the most appropriate Call Driver Level 1 from the options provided in the rubric. This should be a string value that exactly matches one of the available options.

    For the "callDriverLevel2" field: This will be determined based on the Level 1 selection and should be left as an empty string ("") since the cascading logic will be handled by the frontend.

    Here is an example of the required structure:
    {
      "skillEvaluations": {
        "Core": [
          { "skillName": "Willingness to Help", "rating": "DW", "remarks": "The TM showed immediate willingness by stating 'I can definitely help you with that'." },
          { "skillName": "Listen & Retain", "rating": "Prt", "remarks": "The TM understood the main issue but had to ask for the account number again after the customer had already provided it." }
        ],
        "Critical": [
          { "skillName": "Secure the Customer's Identity Based on Department Guidelines", "rating": "DD", "remarks": "The TM failed to verify the customer's identity according to the guidelines." }
        ],
        "CLS Skills Evaluation": [
          { "skillName": "Right size", "rating": "Yes", "remarks": "The TM attempted to right-size the plan." }
        ]
      },
      "callSummary": "...",
      "highlights": "...",
      "lowlights": "..."
    }

    --- START OF DETAILED SKILL RUBRIC ---
    ${detailedRubric}
    --- END OF DETAILED SKILL RUBRIC ---

    Based on your analysis of the script, provide a complete JSON object with all four top-level keys populated according to the specified structure.
  `;

  const payload = {
    model: "gpt-4o",
    messages: [{ role: "system", content: systemPrompt }, { role: "user", content: scriptText }],
    temperature: 0.2,
    response_format: { "type": "json_object" }
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: { "Authorization": "Bearer ak-fnRymjzGJpyatnpMnnaAW4ShBfQI" },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch("https://api-beta.fuelix.ai/v1/chat/completions", options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    if (responseCode !== 200) {
      Logger.log(`API Error Response: ${responseText}`);
      throw new Error(`API request failed with status ${responseCode}.`);
    }

    const jsonResponse = JSON.parse(responseText);
    const aiContent = jsonResponse.choices[0].message.content;
    return JSON.parse(aiContent);
  } catch (error) {
    Logger.log(`AI analysis failed: ${error.toString()}`);
    throw new Error(`AI analysis failed. Details: ${error.message}`);
  }
}

/**
 * Returns the full structured rubric details for all skills.
 * @returns {Object} An object containing descriptions, observation points, and rating rules for each skill.
 */
function getRubricDetails() {
  return {
    "Warm and Energetic Greeting that Connects with the Customer": {
      "description": "Opening the call with energy to build rapport and start the call strong.",
      "observe": ["The team member's tone is clearly not boring or too slow paced.", "Welcomed the customer.", "Introduced him/herself by their name and branded TELUS or Koodo.", "Used a friendly statement to connect with the customer."],
      "ratings": { "DW": "All 'What should be observed' has been 100% executed.", "Prt": "Missed 50% of 'what should be observed'.", "DD": "Missed 75% of 'what should be observed' or customer complained about the greeting." }
    },
    "Listen & Retain": {
      "description": "Allowing the customer to explain their reason for calling.",
      "observe": ["Avoided talking over the customer.", "If TM has internet issues and talked over, they should explain and apologize.", "Had active listening and understood the information explained.", "Did not ask for specific information the customer had already provided."],
      "ratings": { "DW": "All 'What should be observed' has been 100% executed.", "Prt": "Missed 50% of 'what should be observed'.", "DD": "Missed 75% of 'what should be observed'." }
    },
    "Willingness to Help": {
      "description": "Showing customer that they reached the right person and we are here to help them with their concern.",
      "observe": ["Demonstrated willingness to help during the entire interaction (consider assurance or willingness statements, or willingness attitude through tone, or specific actions like being polite, prompt, or assertive).", "Rate this from a customer's perspective."],
      "ratings": { "DW": "All 'What should be observed' has been 100% executed.", "DD": "Missed 100% of 'what should be observed'." }
    },
    "Secure the Customer's Identity Based on Department Guidelines": {
      "description": "Ensuring the departmental policy for verifying the customer is followed.",
      "observe": ["Acknowledged customer entering PIN on IVR.", "Accurately followed authentication guidelines (PIN or 3 pieces of info).", "Politely asked for pieces of information."],
      "ratings": { "DW": "All 'What should be observed' has been 100% executed.", "DD": "Missed 100% of 'what should be observed'." }
    },
    "Probe Using a Combination of Open-ended and Closed-ended Questions": {
      "description": "Asking strategic questions to resolve the customers concern.",
      "observe": ["Preferably asked permission to ask questions.", "Questions should be accurate to understand or resolve the inquiry.", "Start with open-ended questions to diagnose.", "Did not assume the root cause.", "Transition to closed-ended questions for specifics.", "Captured any missing details to identify the root-cause."],
      "ratings": { "DW": "All 'What should be observed' has been 100% executed.", "DD": "Missed 50% of 'what should be observed' or no probing at all." }
    },
    "Investigate & Check All Applicable Resources": {
        "description": "Leverage existing tools to ensure they correctly serve the customer.",
        "observe": ["Used all applicable tools to resolve accurately.", "Provided correct and required info from OneSource.", "Set customer's expectation before/during investigation.", "Checked-in with customer after 2 mins of hold time."],
        "ratings": { "DW": "Criteria met.", "Prt": "Criteria partially met.", "DD": "Criteria not met." } // Assuming generic ratings as none were provided
    },
    "Offer Solutions / Explanations / Options": {
      "description": "Effective positioning of the recommended resolution to the customer's concern.",
      "observe": ["Offered the best possible solution to the customer's main inquiry.", "The resolution should be based from Business Guidelines/One Source to reinforce completeness and correctness."],
      "ratings": { "DW": "Observed Clarity, Proactivity, Effectiveness, Thoroughness, Audience Awareness, and Creativity.", "DD": "One or more of the DW criteria was not observed." }
    },
    "Begin Processing Transaction, Adding Notes and Updating Account Information when Appropriate": {
        "description": "Accurately performing the necessary steps to complete the customers resolution.",
        "observe": ["Accurately processed changes or requests.", "Accurately documented (or diary notes) on all applicable tools (DT1 or Smart Desktop, Outcome, Lynx)."],
        "ratings": { "DW": "Criteria met.", "Prt": "Criteria partially met.", "DD": "Criteria not met."} // Assuming generic ratings
    },
    "Confirm What You Know and Paraphrase Customer": {
      "description": "Using active listening to summarize the customer's request.",
      "observe": ["The agent paraphrased what the customer said when they explained the issue.", "Shared with customer what they knew and saw that is relevant to the customer’s original statement."],
      "ratings": { "DW": "All 'What should be observed' has been 100% executed.", "Prt": "Missed 50% of 'what should be observed'.", "DD": "Missed 75% of 'what should be observed'." }
    },
    "Check for Understanding": {
      "description": "Confirming that the customer fully understands and can verbalize the solution or next steps.",
      "observe": ["Politely asked if what was tackled was clear / paused to check if the customer was following throughout the interaction (Question + Pause)."],
      "ratings": { "DW": "All 'What should be observed' has been 100% executed.", "Prt": "Missed 50% of 'what should be observed'.", "DD": "Missed 75% of 'what should be observed'." }
    },
    "Lead with a Benefit Prior to Asking Questions and Securing the Account": {
      "description": "Giving the customer a reason to do something before asking them to do it.",
      "observe": ["Gave the customer a reason to do something before asking them to do it."],
      "ratings": { "DW": "All 'What should be observed' has been 100% executed.", "Prt": "Missed 50% of 'what should be observed'.", "DD": "Missed 75% of 'what should be observed' or no lead with benefit statement." }
    },
    "Summarize Solution": {
      "description": "Recapping all steps taken in the interaction.",
      "observe": ["Clearly recapped the key steps they took for the customer.", "Clearly recapped the key pieces of information that the customer needed to remember."],
      "ratings": { "DW": "All 'What should be observed' has been 100% executed.", "Prt": "Missed 50% of 'what should be observed'.", "DD": "Missed 75% of 'what should be observed'." }
    },
    "End on a High Note with a Personalized Close": {
      "description": "Ending the call with warmth while offering an authentic personal goodbye.",
      "observe": ["Thanked the customer for contacting.", "Wished the customer good-bye.", "Used what you learned from the customer to end the call with a personal touch."],
      "ratings": { "DW": "All 'What should be observed' has been 100% executed.", "DD": "Missed 100% of 'what should be observed'." }
    },
    "Reinforce the Extras You Did": {
      "description": "Sharing the additional value you were able to provide for the customer.",
      "observe": ["Highlighted 'value' provided (outside of solutions to main reason for contacting)."],
      "ratings": { "DW": "All 'What should be observed' has been 100% executed.", "DD": "Missed 100% of 'what should be observed'." }
    },
    "Early Bridge to Value When Appropriate": {
      "description": "Planting a seed for an extra discussion later in the call.",
      "observe": ["Set Expectations by Leading with Benefit (Early Bridge Step 2).", "Ask for time/permission or give a Call to Action (Early Bridge Step 3)."],
      "ratings": { "DW": "All 'What should be observed' has been 100% executed.", "DD": "Missed 100% of 'what should be observed'." }
    },
    "Later Bridge to Value when Appropriate / Value Proposition": {
      "description": "Opening up a discussion about something outside of the reason for the customer calling in that will add value to their services.",
      "observe": ["Acknowledged the resolution of their initial inquiry.", "Set Expectations by Leading with Benefit.", "Asked for time / confirmed to proceed.", "Offered to add value that applies to the customer's needs."],
      "ratings": { "DW": "All 'What should be observed' has been 100% executed.", "DD": "Missed 100% of 'what should be observed'." }
    },
    "Lead Customer to Accept an Option or Solution": {
      "description": "Guiding the customer to agree to the resolution.",
      "observe": ["REINFORCE: Clearly and concisely summarized the discussion to reinforce the proposed solution.", "CONFIRM: Used a clear, direct question to confirm the customer's acceptance of the solution."],
      "ratings": { "DW": "Agent executed both Reinforcement and Confirmation components effectively.", "Prt": "Agent attempted but didn't fully or effectively execute both components.", "DD": "The criteria for Reinforcement and Confirmation were not observed." }
    },
    "Check for Satisfaction": {
      "description": "Ensuring the customer is satisfied with the interaction.",
      "observe": ["Clearly asked the customer about how comfortable they feel with the solutions or next steps provided.", "If a callback is required, this is 'Did Well' as long as TM was clear and provided assurance."],
      "ratings": { "DW": "All 'What should be observed' has been 100% executed.", "DD": "Missed 100% of 'what should be observed'." }
    },
    "Overcome Objections if Necessary": {
      "description": "Addressing any questions or concerns the customer may have about the resolution.",
      "observe": ["Acknowledged the customer’s objection or used empathy if necessary.", "Probed to understand the customer’s objection.", "Answered the customer’s questions / reason for objecting."],
      "ratings": { "DW": "All 'What should be observed' has been 100% executed.", "Prt": "Missed 50% of 'what should be observed'.", "DD": "Missed 75% of 'what should be observed'." }
    },
    "Acknowledge (Empathy When Needed)": {
      "description": "Show the customer we are listening and that we care about what they are saying.",
      "observe": ["Acknowledged or showed empathy when needed throughout the entire call, including positive, neutral and negative comments or situations by the customer."],
      "ratings": { "DW": "All 'What should be observed' has been 100% executed.", "Prt": "Missed 50% of 'what should be observed'.", "DD": "Missed 75% of 'what should be observed' or no acknowledgement or empathy statement." }
    },
    "Learn from Your Customer": {
      "description": "Asking about current and future needs.",
      "observe": ["Get the full picture by asking open & closed-ended questions with relevant follow up questions.", "Listen for clues about the customer's frequency and interest."],
      "ratings": { "DW": "All 'What should be observed' has been 100% executed.", "Prt": "Missed 50% of 'what should be observed'.", "DD": "Missed 75% of 'what should be observed'." }
    },
    "Handling Early Objections": {
      "description": "Validating concerns and transitioning to benefits.",
      "observe": ["Acknowledge the customer's comments and emotion.", "Transition back to WIIFM (What's In It For Me)."],
      "ratings": { "DW": "All 'What should be observed' has been 100% executed.", "Prt": "Missed 50% of 'what should be observed'.", "DD": "Missed 75% of 'what should be observed'." }
    },
    "Customer Intimacy": {
      "description": "Knowing what to offer and leading with benefits.",
      "observe": ["Learn the customer's frequency of use and level of interest.", "Learn what the customer would like, doesn't like, & doesn't have.", "Prioritize the WIIFM and lead with the strongest benefit."],
      "ratings": { "DW": "All 'What should be observed' has been 100% executed.", "Prt": "Missed 50% of 'what should be observed'.", "DD": "Missed 75% of 'what should be observed'." }
    },
    "Value Proposition": {
      "description": "Same as 'Later Bridge to Value'.", // Re-using for consistency
      "observe": ["Acknowledged the resolution of their initial inquiry.", "Set Expectations by Leading with Benefit.", "Asked for time / confirmed to proceed.", "Offered to add value that applies to the customer's needs."],
      "ratings": { "DW": "All 'What should be observed' has been 100% executed.", "DD": "Missed 100% of 'what should be observed'." }
    },
    "Right size": {
      "description": "Agents assess the customer's actual usage, requirements, and budget to recommend plans or services that are an optimal fit.",
      "observe": ["Assesses the customer's actual usage.", "Considers customer's requirements and budget.", "Recommends plans/services that are neither over- nor under-provisioned."],
      "ratings": { "Yes": "At least an attempt was made.", "No": "No attempt was made." }
    },
    "Use a tiered approach": {
      "description": "Agents start with initial offers and progressively present more attractive options if the customer remains unsatisfied.",
      "observe": ["Starts with a basic/initial offer first.", "Presents more attractive options only if the customer is not satisfied."],
      "ratings": { "Yes": "The agent used a tiered approach.", "No": "The agent immediately jumped to the 'best' or 'final' offer." }
    },
    "Sufficient probing (+3 Questions)": {
      "description": "This skill requires agents to ask at least three meaningful questions to fully understand the customer's situation, concerns, and motivations for wanting to leave.",
      "observe": ["Asks at least three meaningful questions.", "Questions help uncover the root cause of dissatisfaction.", "Questions gather information needed for targeted retention offers."],
      "ratings": { "Yes": "Agent demonstrated a thorough understanding by asking enough meaningful questions.", "No": "Agent failed to gather necessary information and/or jumped to a solution without sufficient probing." }
    },
    "Speak to value (WHY TELUS, reality check and service superiority)": {
      "description": "Agents articulate TELUS's unique value proposition by highlighting competitive advantages, service quality, network reliability, and overall value.",
      "observe": ["Highlights competitive advantages (e.g., network reliability, service quality).", "Helps customer understand what they might lose by switching ('reality check')."],
      "ratings": { "Yes": "The agent explained the value of choosing TELUS.", "No": "The agent did not explain the value of choosing TELUS." }
    },
    "Loyalty statements implementation": {
      "description": "Agents use statements that acknowledge the customer's loyalty, express appreciation for their business, and reinforce their value to TELUS.",
      "observe": ["Uses statements that acknowledge the customer's tenure/loyalty.", "Expresses appreciation for their business.", "Reinforces their value to the company."],
      "ratings": { "Yes": "The agent used appropriate loyalty statements.", "No": "No loyalty statement was used." }
    },
    "Investigation (Review bill, compare competitor offer, search notes on the account, take a look at usage)": {
        "description": "A thorough examination of the customer's account and history to inform the retention strategy.",
        "observe": ["Reviews the customer's bill.", "Compares competitor's offer if mentioned.", "Searches notes on the account.", "Looks at service usage."],
        "ratings": { "Yes": "Criteria met.", "No": "Criteria not met." }
    },
    "Was the agent able to retain the customer?": {
      "description": "This measures whether the agent successfully convinced the customer to stay with TELUS.",
      "observe": ["Observe the final outcome of the call regarding the customer's services."],
      "ratings": { "Yes": "No service was cancelled.", "No": "The agent processed a cancellation." }
    }
  };
}

/**
 * Fetches the form structure with dynamic dropdown options.
 */
function getFormStructure() {
  try {
    // Get dynamic dropdown options
    const dropdownOptions = getDropdownOptions();
    
    const radioOptions = ["DW", "Prt", "DD", "NA"];
    const radioOptionsNoPrt = ["DW", "Prt-disabled", "DD", "NA"]; // For skills that don't have partial option - Prt will be disabled
    const clsRadioOptions = ["Yes", "No", "NA"];
    const formStructure = [
      {
        title: "Team Member Details",
        fields: [
          { name: "Sap ID", type: "text" }, { name: "Team Member", type: "text" },
          { name: "Team Leader", type: "text" }, { name: "Operations Manager", type: "text" },
          { name: "Line of Business", type: "text" },
        ], subsections: {}
      },
      {
        title: "Call Details",
        fields: [
          { name: "Interaction Date", type: "date" }, { name: "Customer's BAN", type: "text" },
          { name: "FLP Conversation ID", type: "text" },
          { name: "Listening Type", type: "dropdown", options: dropdownOptions["Listening Type"] },
          { name: "Duration (minutes)", type: "number" },
        ], subsections: {}
      },
      {
        title: "Customer Experience Blueprint", collapsible: true, mandatory: true, fields: [],
        subsections: {
          "Core": [
            { name: "Willingness to Help", type: "radio", options: radioOptions },
            { name: "Warm and Energetic Greeting that Connects with the Customer", type: "radio", options: radioOptions },
            { name: "Listen & Retain", type: "radio", options: radioOptions }
          ], "Critical": [
            { name: "Secure the Customer's Identity Based on Department Guidelines", type: "radio", options: radioOptionsNoPrt },
            { name: "Probe Using a Combination of Open-ended and Closed-ended Questions", type: "radio", options: radioOptions },
            { name: "Investigate & Check All Applicable Resources", type: "radio", options: radioOptionsNoPrt },
            { name: "Offer Solutions / Explanations / Options", type: "radio", options: radioOptionsNoPrt },
            { name: "Begin Processing Transaction, Adding Notes and Updating Account Information when Appropriate", type: "radio", options: radioOptionsNoPrt }
          ], "Basic": [
            { name: "Confirm What You Know and Paraphrase Customer", type: "radio", options: radioOptions },
            { name: "Check for Understanding", type: "radio", options: radioOptions },
            { name: "Lead with a Benefit Prior to Asking Questions and Securing the Account", type: "radio", options: radioOptions },
            { name: "Summarize Solution", type: "radio", options: radioOptions },
            { name: "End on a High Note with a Personalized Close", type: "radio", options: radioOptions },
            { name: "Reinforce the Extras You Did", type: "radio", options: radioOptions }
          ], "Intermediate": [
            { name: "Early Bridge to Value When Appropriate", type: "radio", options: radioOptions },
            { name: "Later Bridge to Value when Appropriate / Value Proposition", type: "radio", options: radioOptions },
            { name: "Lead Customer to Accept an Option or Solution", type: "radio", options: radioOptions },
            { name: "Check for Satisfaction", type: "radio", options: radioOptions }
          ], "Advanced": [
            { name: "Overcome Objections if Necessary", type: "radio", options: radioOptions },
            { name: "Acknowledge (Empathy When Needed)", type: "radio", options: radioOptions }
          ], "Sales Skills": [
            { name: "Learn from Your Customer", type: "radio", options: radioOptions },
            { name: "Handling Early Objections", type: "radio", options: radioOptions },
            { name: "Customer Intimacy", type: "radio", options: radioOptions },
            { name: "Value Proposition", type: "radio", options: radioOptions },
            { name: "What type of bridging was provided?", type: "dropdown", options: dropdownOptions["What type of bridging was provided?"] }
          ]
        }
      },
      {
        title: "Critical Behaviors Check-up", collapsible: true,
        fields: [],
        subsections: {
          "Critical Behaviors": [
            { name: "Abusive, insult the customer (disrespectful)", type: "checkbox" },
            { name: "Appropriation of the sale from another agent", type: "checkbox" },
            { name: "CCTS Process not Followed", type: "checkbox" },
            { name: "Delayed Response to Inbound Contact", type: "checkbox" },
            { name: "Didn't complete transaction as promised", type: "checkbox" },
            { name: "Excessive Dead Air, Hold or Idle Time (with the clear intention of making the customer disconnect the call)", type: "checkbox" },
            { name: "Failure to follow processes as per Onesource or use tools properly", type: "checkbox" },
            { name: "Incorrect sales order processing", type: "checkbox" },
            { name: "Incorrect use of credits/discounts", type: "checkbox" },
            { name: "Misleading Information / Deliberately lying", type: "checkbox" },
            { name: "Misleading product/offer information", type: "checkbox" },
            { name: "Missed Call Back", type: "checkbox" },
            { name: "Modify customer's account without consent", type: "checkbox" },
            { name: "No Authentication/Verification", type: "checkbox" },
            { name: "No Response to Inbound Contact", type: "checkbox" },
            { name: "No Retention Attempt or No Value Generation effort", type: "checkbox" },
            { name: "Nonwork related distractions", type: "checkbox" },
            { name: "Overtalking to the customer, interrupt the customer", type: "checkbox" },
            { name: "Proactively transferring calls", type: "checkbox" },
            { name: "Staying on the line unnecessarily", type: "checkbox" },
            { name: "System/Tools Manipulation", type: "checkbox" },
            { name: "Unjustified disconnection of customer interaction", type: "checkbox" },
            { name: "Unnecessary/Invalid Consult, Dispatch or Transfer", type: "checkbox" },
            { name: "Unwillingness to Help", type: "checkbox" },
            { name: "Other", type: "checkbox_with_text", textField: "CB_Other_Behavior_Text" }
          ]
        }
      },
      {
        title: "Customer Loyalty & Support (CLS)", collapsible: true,
        fields: [
          { name: "Cancellation Root Cause", type: "dropdown", options: dropdownOptions["Cancellation Root Cause"] },
          { name: "Was there an unethical credit?", type: "dropdown", options: dropdownOptions["Was there an unethical credit?"] }
        ],
        subsections: {
          "CLS Skills Evaluation": [
            { name: "Right size", type: "radio", options: clsRadioOptions }, { name: "Use a tiered approach", type: "radio", options: clsRadioOptions },
            { name: "Sufficient probing (+3 Questions)", type: "radio", options: clsRadioOptions }, { name: "Speak to value (WHY TELUS, reality check and service superiority)", type: "radio", options: clsRadioOptions },
            { name: "Loyalty statements implementation", type: "radio", options: clsRadioOptions }, { name: "Investigation (Review bill, compare competitor offer, search notes on the account, take a look at usage)", type: "radio", options: clsRadioOptions },
            { name: "Was the agent able to retain the customer?", type: "radio", options: clsRadioOptions }
          ]
        }
      },
      {
        title: "Call Evaluation Details", mandatory: true,
        fields: [
          { name: "Call Driver Level 1", type: "dropdown", options: dropdownOptions["Call Driver Level 1"] },
          { name: "Call Driver Level 2", type: "dropdown", options: [], cascading: true, dependsOn: "Call Driver Level 1", mapping: dropdownOptions["Call Driver Mapping"] },
          { name: "Call Summary", type: "textarea", rows: 13 },
          { name: "Highlights", type: "textarea", rows: 5 },
          { name: "Lowlights/Recommendation", type: "textarea", rows: 5 }
        ],
        subsections: {}
      }
    ];
    return formStructure;
  } catch (error) {
    Logger.log(`Error in getFormStructure: ${error.message}`);
    return { error: "Could not build form structure.", details: error.message };
  }
}

/**
 * Test function to verify the enhanced data structure implementation
 */
function testEnhancedDataStructure() {
  // Sample form data for testing
  const testFormData = {
    "Sap ID": "12345",
    "Team Member": "John Doe",
    "Team Leader": "Jane Smith",
    "Interaction Date": "2024-01-15",
    "Willingness to Help": "DW",
    "Call Summary": "Test call summary",
    "Highlights": "Test highlights"
  };
  
  try {
    // Test Record ID generation
    const recordId = generateRecordID();
    console.log(`Generated Record ID: ${recordId}`);
    
    // Test enhanced data structure creation
    const enhancedData = createEnhancedFormData(testFormData);
    console.log(`Enhanced data created with Record ID: ${enhancedData.recordId}`);
    
    // Test validation
    const validation = validateEnhancedData(enhancedData);
    console.log(`Validation result: ${validation.isValid}, Quality Score: ${validation.qualityScore}%`);
    
    // Test row data preparation
    const rowData = prepareRowData(enhancedData);
    console.log(`Row data prepared with ${rowData.length} columns`);
    
    return {
      success: true,
      message: "Enhanced data structure test completed successfully",
      recordId: enhancedData.recordId,
      qualityScore: validation.qualityScore
    };
    
  } catch (error) {
    console.error(`Test failed: ${error.message}`);
    return {
      success: false,
      message: `Test failed: ${error.message}`
    };
  }
}

// --- EDIT FUNCTIONALITY FUNCTIONS ---

/**
 * Creates the Edit History sheet if it doesn't exist
 */
function createEditHistorySheet() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let historySheet = ss.getSheetByName("Edit_History");
    
    if (!historySheet) {
      historySheet = ss.insertSheet("Edit_History");
      
      // Set up headers for edit history
      const headers = [
        "RecordID",
        "Version", 
        "EditType",
        "EditedBy",
        "EditTimestamp",
        "OriginalData",
        "ChangeSummary",
        "EditReason"
      ];
      
      historySheet.appendRow(headers);
      historySheet.getRange("1:1").setFontWeight("bold");
      historySheet.setFrozenRows(1);
      
      Logger.log("Edit History sheet created successfully");
    }
    
    return historySheet;
  } catch (error) {
    Logger.log(`Error creating Edit History sheet: ${error.message}`);
    throw error;
  }
}

/**
 * Debug function to check what data exists in the spreadsheet
 * @returns {Object} Debug information about the spreadsheet
 */
function debugSpreadsheetData() {
  try {
    // Check if user is admin before allowing debug access
    if (!isAdminUser()) {
      Logger.log(`Non-admin user attempted to access debug data: ${Session.getActiveUser().getEmail()}`);
      throw new Error("Access denied. Debug functionality is restricted to admin users only.");
    }
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const responseSheet = ss.getSheetByName(RESPONSE_SHEET_NAME);
    
    if (!responseSheet) {
      return {
        error: `Response sheet "${RESPONSE_SHEET_NAME}" not found`,
        availableSheets: ss.getSheets().map(sheet => sheet.getName())
      };
    }
    
    const data = responseSheet.getDataRange().getValues();
    const headers = data[0];
    const totalRows = data.length;
    
    // Get sample record IDs
    const recordIdColumnIndex = headers.indexOf("RecordID");
    const sampleRecordIds = [];
    
    if (recordIdColumnIndex !== -1) {
      for (let i = 1; i < Math.min(6, totalRows); i++) {
        if (data[i][recordIdColumnIndex]) {
          sampleRecordIds.push(data[i][recordIdColumnIndex]);
        }
      }
    }
    
    Logger.log(`Admin user ${Session.getActiveUser().getEmail()} accessed debug data`);
    
    return {
      sheetExists: true,
      totalRows: totalRows - 1, // Exclude header
      totalColumns: headers.length,
      headers: headers,
      recordIdColumnExists: recordIdColumnIndex !== -1,
      recordIdColumnIndex: recordIdColumnIndex,
      sampleRecordIds: sampleRecordIds,
      hasData: totalRows > 1
    };
    
  } catch (error) {
    return {
      error: error.message,
      details: error.toString()
    };
  }
}

/**
 * Sanitizes data for JSON serialization by converting problematic types
 * @param {Object} data - The data object to sanitize
 * @returns {Object} Sanitized data object
 */
function sanitizeDataForSerialization(data) {
  const sanitized = {};
  
  for (const [key, value] of Object.entries(data)) {
    if (value instanceof Date) {
      // Convert Date objects to ISO strings
      sanitized[key] = value.toISOString();
    } else if (value === null) {
      // Convert null to empty string
      sanitized[key] = "";
    } else if (value === undefined) {
      // Convert undefined to empty string
      sanitized[key] = "";
    } else if (typeof value === 'object' && value !== null) {
      // For nested objects, recursively sanitize
      sanitized[key] = sanitizeDataForSerialization(value);
    } else {
      // Keep primitive values as-is
      sanitized[key] = value;
    }
  }
  
  return sanitized;
}

/**
 * Tests if an object can be JSON serialized
 * @param {Object} obj - Object to test
 * @returns {boolean} True if serializable, false otherwise
 */
function testJsonSerialization(obj) {
  try {
    JSON.stringify(obj);
    return true;
  } catch (error) {
    Logger.log(`JSON serialization test failed: ${error.message}`);
    return false;
  }
}

/**
 * Retrieves an evaluation by Record ID
 * @param {string} recordId - The Record ID to search for
 * @returns {Object|null} Evaluation data or null if not found
 */
function getEvaluationByRecordId(recordId) {
  try {
    Logger.log(`=== SEARCH DEBUG START ===`);
    Logger.log(`Searching for Record ID: "${recordId}"`);
    Logger.log(`Record ID length: ${recordId.length}`);
    Logger.log(`Record ID type: ${typeof recordId}`);
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const responseSheet = ss.getSheetByName(RESPONSE_SHEET_NAME);
    
    if (!responseSheet) {
      throw new Error(`Response sheet "${RESPONSE_SHEET_NAME}" not found.`);
    }
    
    const data = responseSheet.getDataRange().getValues();
    const headers = data[0];
    
    Logger.log(`Sheet has ${data.length - 1} data rows and ${headers.length} columns`);
    Logger.log(`Headers: ${headers.slice(0, 5).join(', ')}...`);
    
    // Find the RecordID column (should be first column)
    const recordIdColumnIndex = headers.indexOf("RecordID");
    if (recordIdColumnIndex === -1) {
      Logger.log(`Available headers: ${headers.join(', ')}`);
      throw new Error("RecordID column not found in response sheet");
    }
    
    Logger.log(`RecordID column found at index: ${recordIdColumnIndex}`);
    
    // Search for the record
    for (let i = 1; i < data.length; i++) {
      const currentRecordId = data[i][recordIdColumnIndex];
      Logger.log(`Row ${i}: RecordID = "${currentRecordId}" (length: ${currentRecordId ? currentRecordId.length : 'null'}, type: ${typeof currentRecordId})`);
      Logger.log(`Comparison: "${currentRecordId}" === "${recordId}" = ${currentRecordId === recordId}`);
      Logger.log(`Trimmed comparison: "${currentRecordId ? currentRecordId.toString().trim() : ''}" === "${recordId.trim()}" = ${currentRecordId ? currentRecordId.toString().trim() === recordId.trim() : false}`);
      
      // Try multiple comparison methods
      if (currentRecordId === recordId || 
          (currentRecordId && currentRecordId.toString().trim() === recordId.trim()) ||
          (currentRecordId && currentRecordId.toString() === recordId)) {
        Logger.log(`✅ MATCH FOUND at row ${i + 1}!`);
        
        // Convert row data to object using headers
        const rawEvaluationData = {};
        headers.forEach((header, index) => {
          rawEvaluationData[header] = data[i][index];
        });
        
        // Add row number for future updates
        rawEvaluationData._rowNumber = i + 1;
        
        Logger.log(`Raw evaluation data created with ${Object.keys(rawEvaluationData).length} fields`);
        Logger.log(`Sample fields: RecordID=${rawEvaluationData.RecordID}, Team Member=${rawEvaluationData['Team Member']}`);
        
        // Sanitize data for JSON serialization
        const sanitizedData = sanitizeDataForSerialization(rawEvaluationData);
        Logger.log(`Data sanitized for JSON serialization`);
        
        // Test JSON serialization
        const isSerializable = testJsonSerialization(sanitizedData);
        Logger.log(`JSON serialization test: ${isSerializable ? 'PASSED' : 'FAILED'}`);
        
        if (!isSerializable) {
          Logger.log(`❌ Data cannot be serialized, returning null`);
          return null;
        }
        
        // Log the sanitized data
        try {
          const jsonString = JSON.stringify(sanitizedData, null, 2);
          Logger.log(`Sanitized evaluation data JSON (first 500 chars): ${jsonString.substring(0, 500)}...`);
        } catch (jsonError) {
          Logger.log(`❌ Failed to log JSON: ${jsonError.message}`);
        }
        
        Logger.log(`About to return sanitized evaluation data to frontend...`);
        Logger.log(`=== SEARCH DEBUG END - SUCCESS ===`);
        
        // Return the sanitized data
        return sanitizedData;
      }
    }
    
    Logger.log(`❌ No matching record found for RecordID: "${recordId}"`);
    Logger.log(`=== SEARCH DEBUG END - NOT FOUND ===`);
    return null; // Record not found
    
  } catch (error) {
    Logger.log(`❌ Error retrieving evaluation by Record ID: ${error.message}`);
    Logger.log(`Error stack: ${error.stack}`);
    Logger.log(`=== SEARCH DEBUG END - ERROR ===`);
    throw error;
  }
}

/**
 * Searches for evaluations by SAP ID and date
 * @param {string} sapId - The SAP ID to search for
 * @param {string} interactionDate - The interaction date (YYYY-MM-DD format)
 * @returns {Array} Array of matching evaluations
 */
function searchEvaluationsBySapIdAndDate(sapId, interactionDate) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const responseSheet = ss.getSheetByName(RESPONSE_SHEET_NAME);
    
    if (!responseSheet) {
      throw new Error(`Response sheet "${RESPONSE_SHEET_NAME}" not found.`);
    }
    
    const data = responseSheet.getDataRange().getValues();
    const headers = data[0];
    
    // Find relevant column indices
    const recordIdIndex = headers.indexOf("RecordID");
    const sapIdIndex = headers.indexOf("Sap ID");
    const interactionDateIndex = headers.indexOf("Interaction Date");
    const teamMemberIndex = headers.indexOf("Team Member");
    const timestampIndex = headers.indexOf("Timestamp");
    
    if (recordIdIndex === -1 || sapIdIndex === -1 || interactionDateIndex === -1) {
      throw new Error("Required columns not found in response sheet");
    }
    
    const matches = [];
    
    // Search for matching records
    for (let i = 1; i < data.length; i++) {
      const rowSapId = data[i][sapIdIndex];
      const rowDate = data[i][interactionDateIndex];
      
      // Convert date to string format for comparison
      let dateToCompare = rowDate;
      if (rowDate instanceof Date) {
        dateToCompare = rowDate.toISOString().split('T')[0]; // YYYY-MM-DD format
      }
      
      if (rowSapId === sapId && dateToCompare === interactionDate) {
        matches.push({
          recordId: data[i][recordIdIndex],
          sapId: rowSapId,
          teamMember: data[i][teamMemberIndex] || 'Unknown',
          interactionDate: dateToCompare,
          timestamp: data[i][timestampIndex] || '',
          _rowNumber: i + 1
        });
      }
    }
    
    return matches;
    
  } catch (error) {
    Logger.log(`Error searching evaluations: ${error.message}`);
    throw error;
  }
}

/**
 * Updates an existing evaluation in place
 * @param {string} recordId - The Record ID to update
 * @param {Object} newFormData - The new form data
 * @param {string} editReason - Reason for the edit (optional)
 * @returns {Object} Update result
 */
function updateExistingEvaluation(recordId, newFormData, editReason = '') {
  try {
    Logger.log(`=== UPDATE EVALUATION DEBUG START ===`);
    Logger.log(`Updating Record ID: ${recordId}`);
    Logger.log(`Edit reason: ${editReason}`);
    
    // First, get the current evaluation data for backup
    const currentEvaluation = getEvaluationByRecordId(recordId);
    if (!currentEvaluation) {
      throw new Error(`Evaluation with Record ID "${recordId}" not found.`);
    }
    
    Logger.log(`Current evaluation found, row: ${currentEvaluation._rowNumber}`);
    
    // Create enhanced data structure from new form data
    const enhancedData = createEnhancedFormData(newFormData);
    
    // Override the Record ID to maintain the same ID
    enhancedData.recordId = recordId;
    
    // Fix version handling - handle both string and undefined cases
    let currentVersion = currentEvaluation.FormVersion || currentEvaluation.formVersion || "1.0";
    Logger.log(`Current version: ${currentVersion} (type: ${typeof currentVersion})`);
    
    // Ensure currentVersion is a string
    if (typeof currentVersion !== 'string') {
      currentVersion = String(currentVersion);
    }
    
    // Handle version increment safely
    let newVersion;
    try {
      const versionParts = currentVersion.split('.');
      const majorVersion = parseInt(versionParts[0]) || 1;
      const minorVersion = parseInt(versionParts[1]) || 0;
      newVersion = `${majorVersion}.${minorVersion + 1}`;
    } catch (versionError) {
      Logger.log(`Version parsing error: ${versionError.message}, using default`);
      newVersion = "1.1"; // Default fallback
    }
    
    Logger.log(`New version: ${newVersion}`);
    enhancedData.formVersion = newVersion;
    
    // Update metadata
    enhancedData.submittedBy = Session.getActiveUser().getEmail();
    enhancedData.timestamp = new Date();
    
    // Validate the new data
    const validation = validateEnhancedData(enhancedData);
    if (!validation.isValid) {
      Logger.log(`Validation warnings for update: ${validation.errors.join(', ')}`);
    }
    
    // Log the edit to history before updating
    logEditToHistory(recordId, currentEvaluation, enhancedData, editReason);
    
    // Update the main sheet row
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const responseSheet = ss.getSheetByName(RESPONSE_SHEET_NAME);
    const rowData = prepareRowData(enhancedData);
    
    // Update the specific row
    const range = responseSheet.getRange(currentEvaluation._rowNumber, 1, 1, rowData.length);
    range.setValues([rowData]);
    
    // Send update confirmation email to user
    try {
      sendSubmissionConfirmation(newFormData, recordId, true); // true indicates this is an update
      Logger.log(`Update confirmation email sent for Record ID: ${recordId}`);
    } catch (emailError) {
      Logger.log(`Update email sending failed: ${emailError.message}`);
      // Don't fail the entire update if email fails
    }
    
    Logger.log(`Evaluation ${recordId} updated successfully to version ${newVersion}`);
    
    return {
      success: true,
      message: `Evaluation updated successfully! New version: ${newVersion}`,
      recordId: recordId,
      version: newVersion,
      qualityScore: validation.qualityScore
    };
    
  } catch (error) {
    Logger.log(`Error updating evaluation: ${error.message}`);
    return {
      success: false,
      message: `Failed to update evaluation: ${error.message}`
    };
  }
}

/**
 * Logs an edit to the history sheet
 * @param {string} recordId - The Record ID being edited
 * @param {Object} originalData - The original evaluation data
 * @param {Object} newData - The new evaluation data
 * @param {string} editReason - Reason for the edit
 */
function logEditToHistory(recordId, originalData, newData, editReason = '') {
  try {
    const historySheet = createEditHistorySheet();
    
    // Create a summary of changes
    const changeSummary = generateChangeSummary(originalData, newData);
    
    // Prepare history row
    const historyRow = [
      recordId,
      newData.formVersion,
      "Edit",
      Session.getActiveUser().getEmail(),
      new Date(),
      JSON.stringify(originalData), // Store original data as JSON
      changeSummary,
      editReason
    ];
    
    historySheet.appendRow(historyRow);
    
    Logger.log(`Edit logged to history for Record ID: ${recordId}`);
    
  } catch (error) {
    Logger.log(`Error logging edit to history: ${error.message}`);
    // Don't throw error here as it shouldn't prevent the main update
  }
}

/**
 * Generates a summary of changes between original and new data
 * @param {Object} originalData - Original evaluation data
 * @param {Object} newData - New evaluation data
 * @returns {string} Summary of changes
 */
function generateChangeSummary(originalData, newData) {
  const changes = [];
  
  // Compare key fields that are likely to be edited
  const fieldsToCheck = [
    'Willingness to Help',
    'Call Summary',
    'Highlights',
    'Lowlights/Recommendation',
    'Call Driver Level 1',
    'Call Driver Level 2'
  ];
  
  fieldsToCheck.forEach(field => {
    const originalValue = originalData[field];
    const newValue = newData[field] || getNestedValue(newData, field);
    
    if (originalValue !== newValue && newValue !== undefined) {
      changes.push(`${field}: "${originalValue}" → "${newValue}"`);
    }
  });
  
  if (changes.length === 0) {
    return "Minor updates made";
  }
  
  return changes.slice(0, 3).join('; ') + (changes.length > 3 ? '; ...' : '');
}

/**
 * Helper function to get nested values from enhanced data structure
 * @param {Object} data - The enhanced data object
 * @param {string} fieldName - The field name to search for
 * @returns {any} The field value or undefined
 */
function getNestedValue(data, fieldName) {
  // Search through the nested structure
  const searchIn = [
    data.skillEvaluations?.coreSkills,
    data.skillEvaluations?.criticalSkills,
    data.skillEvaluations?.basicSkills,
    data.skillEvaluations?.intermediateSkills,
    data.skillEvaluations?.advancedSkills,
    data.skillEvaluations?.salesSkills,
    data.skillEvaluations?.clsSkills,
    data.callDetails,
    data.evaluationContent
  ];
  
  for (const obj of searchIn) {
    if (obj && typeof obj === 'object') {
      for (const [key, value] of Object.entries(obj)) {
        if (key === fieldName || key.includes(fieldName.replace(/\s+/g, ''))) {
          return value;
        }
      }
    }
  }
  
  return undefined;
}

// --- DATA MIGRATION FUNCTIONS ---

/**
 * Main function to migrate existing data to include Critical Behaviors columns
 * This function safely adds the new columns and populates them with default values
 */
function migrateCriticalBehaviorsData() {
  try {
    Logger.log("Starting Critical Behaviors data migration...");
    
    // Step 1: Validate environment and get spreadsheet
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const responseSheet = ss.getSheetByName(RESPONSE_SHEET_NAME);
    
    if (!responseSheet) {
      throw new Error(`Response sheet "${RESPONSE_SHEET_NAME}" not found.`);
    }
    
    // Step 2: Create backup
    const backupResult = createMigrationBackup(ss, responseSheet);
    Logger.log(`Backup created: ${backupResult.backupSheetName}`);
    
    // Step 3: Analyze current structure
    const analysis = analyzeCurrentStructure(responseSheet);
    Logger.log(`Current structure: ${analysis.totalRows} rows, ${analysis.totalColumns} columns`);
    
    if (analysis.hasCriticalBehaviorsColumns) {
      Logger.log("Critical Behaviors columns already exist. Migration may have been run before.");
      return {
        success: true,
        message: "Critical Behaviors columns already exist. No migration needed.",
        details: analysis
      };
    }
    
    // Step 4: Add new headers if needed
    const headerResult = addCriticalBehaviorsHeaders(responseSheet, analysis);
    Logger.log(`Headers updated: ${headerResult.headersAdded} new columns added`);
    
    // Step 5: Migrate existing data
    const migrationResult = migrateExistingRows(responseSheet, analysis);
    Logger.log(`Data migration completed: ${migrationResult.rowsUpdated} rows updated`);
    
    // Step 6: Validate migration
    const validationResult = validateMigration(responseSheet);
    Logger.log(`Migration validation: ${validationResult.isValid ? 'PASSED' : 'FAILED'}`);
    
    const finalResult = {
      success: true,
      message: "Critical Behaviors data migration completed successfully!",
      details: {
        backupSheet: backupResult.backupSheetName,
        originalRows: analysis.totalRows,
        originalColumns: analysis.totalColumns,
        newColumns: headerResult.headersAdded,
        rowsUpdated: migrationResult.rowsUpdated,
        totalColumns: validationResult.totalColumns,
        migrationTime: new Date().toISOString()
      }
    };
    
    Logger.log(`Migration completed successfully: ${JSON.stringify(finalResult)}`);
    return finalResult;
    
  } catch (error) {
    Logger.log(`Migration failed: ${error.message}`);
    return {
      success: false,
      message: `Migration failed: ${error.message}`,
      error: error.toString()
    };
  }
}

/**
 * Creates a timestamped backup of the response sheet before migration
 */
function createMigrationBackup(spreadsheet, sourceSheet) {
  const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
  const backupSheetName = `${RESPONSE_SHEET_NAME}_Backup_${timestamp}`;
  
  // Copy the entire sheet
  const backupSheet = sourceSheet.copyTo(spreadsheet);
  backupSheet.setName(backupSheetName);
  
  // Add backup metadata
  const metadataRange = backupSheet.getRange(1, 1, 1, 1);
  const note = `Backup created before Critical Behaviors migration on ${new Date().toISOString()}`;
  metadataRange.setNote(note);
  
  return {
    backupSheetName: backupSheetName,
    timestamp: timestamp,
    note: note
  };
}

/**
 * Analyzes the current structure of the response sheet
 */
function analyzeCurrentStructure(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  
  let headers = [];
  let hasCriticalBehaviorsColumns = false;
  
  if (lastRow > 0) {
    headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
    
    // Check if Critical Behaviors columns already exist
    const criticalBehaviorsHeaders = [
      "CB_Abusive_Insult_Customer",
      "CB_Other_Behavior_Text"
    ];
    
    hasCriticalBehaviorsColumns = criticalBehaviorsHeaders.some(header => 
      headers.includes(header)
    );
  }
  
  return {
    totalRows: lastRow,
    totalColumns: lastColumn,
    headers: headers,
    hasCriticalBehaviorsColumns: hasCriticalBehaviorsColumns,
    dataRows: Math.max(0, lastRow - 1) // Exclude header row
  };
}

/**
 * Adds Critical Behaviors headers to the response sheet
 */
function addCriticalBehaviorsHeaders(sheet, analysis) {
  if (analysis.hasCriticalBehaviorsColumns) {
    return { headersAdded: 0, message: "Headers already exist" };
  }
  
  // Get the new headers from setupEnhancedHeaders function logic
  const newHeaders = [
    // Critical Behaviors boolean columns
    "CB_Abusive_Insult_Customer", "CB_Appropriation_Sale_Another_Agent", "CB_CCTS_Process_Not_Followed", 
    "CB_Delayed_Response_Inbound", "CB_Didnt_Complete_Transaction_Promised", "CB_Excessive_Dead_Air_Hold_Idle",
    "CB_Failure_Follow_Processes_Onesource", "CB_Incorrect_Sales_Order_Processing", "CB_Incorrect_Use_Credits_Discounts",
    "CB_Misleading_Information_Lying", "CB_Misleading_Product_Offer_Info", "CB_Missed_Call_Back",
    "CB_Modify_Account_Without_Consent", "CB_No_Authentication_Verification", "CB_No_Response_Inbound_Contact",
    "CB_No_Retention_Value_Generation", "CB_Nonwork_Related_Distractions", "CB_Overtalking_Interrupt_Customer",
    "CB_Proactively_Transferring_Calls", "CB_Staying_Line_Unnecessarily", "CB_System_Tools_Manipulation",
    "CB_Unjustified_Disconnection", "CB_Unnecessary_Invalid_Consult_Transfer", "CB_Unwillingness_To_Help",
    "CB_Other_Behavior", "CB_Other_Behavior_Text"
  ];
  
  // Add headers to the first row
  const startColumn = analysis.totalColumns + 1;
  const headerRange = sheet.getRange(1, startColumn, 1, newHeaders.length);
  headerRange.setValues([newHeaders]);
  
  // Format the new headers
  headerRange.setFontWeight("bold");
  headerRange.setBackground("#f0f0f0");
  
  return {
    headersAdded: newHeaders.length,
    startColumn: startColumn,
    newHeaders: newHeaders
  };
}

/**
 * Migrates existing rows to include default values for Critical Behaviors columns
 */
function migrateExistingRows(sheet, analysis) {
  if (analysis.dataRows === 0) {
    return { rowsUpdated: 0, message: "No data rows to migrate" };
  }
  
  // Prepare default values for Critical Behaviors columns
  // 24 numeric columns (0) + 1 text column (empty string)
  const defaultValues = Array(24).fill(0).concat([""]);
  
  // Get the current total columns after adding headers
  const currentColumns = sheet.getLastColumn();
  const startColumn = currentColumns - 24; // Start of Critical Behaviors columns
  
  // Process rows in batches to avoid timeout
  const batchSize = 100;
  let totalRowsUpdated = 0;
  
  for (let startRow = 2; startRow <= analysis.totalRows; startRow += batchSize) {
    const endRow = Math.min(startRow + batchSize - 1, analysis.totalRows);
    const rowCount = endRow - startRow + 1;
    
    // Create a 2D array with default values for this batch
    const batchData = Array(rowCount).fill().map(() => [...defaultValues]);
    
    // Update the batch
    const range = sheet.getRange(startRow, startColumn, rowCount, defaultValues.length);
    range.setValues(batchData);
    
    totalRowsUpdated += rowCount;
    
    // Log progress for large datasets
    if (analysis.totalRows > 100) {
      Logger.log(`Migration progress: ${totalRowsUpdated}/${analysis.dataRows} rows updated`);
    }
  }
  
  return {
    rowsUpdated: totalRowsUpdated,
    batchesProcessed: Math.ceil(analysis.dataRows / batchSize),
    defaultValuesApplied: defaultValues
  };
}

/**
 * Validates that the migration was successful
 */
function validateMigration(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  
  let isValid = true;
  const issues = [];
  
  // Check if headers exist
  if (lastRow > 0) {
    const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
    
    const requiredHeaders = [
      "CB_Abusive_Insult_Customer",
      "CB_Other_Behavior_Text"
    ];
    
    requiredHeaders.forEach(header => {
      if (!headers.includes(header)) {
        isValid = false;
        issues.push(`Missing header: ${header}`);
      }
    });
    
    // Check if we have the expected number of columns
    const expectedMinColumns = 65; // Original columns + 25 new Critical Behaviors columns
    if (lastColumn < expectedMinColumns) {
      isValid = false;
      issues.push(`Expected at least ${expectedMinColumns} columns, found ${lastColumn}`);
    }
  }
  
  return {
    isValid: isValid,
    totalRows: lastRow,
    totalColumns: lastColumn,
    issues: issues
  };
}

/**
 * Rollback function to restore from backup if needed
 */
function rollbackMigration(backupSheetName) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const backupSheet = ss.getSheetByName(backupSheetName);
    const currentSheet = ss.getSheetByName(RESPONSE_SHEET_NAME);
    
    if (!backupSheet) {
      throw new Error(`Backup sheet "${backupSheetName}" not found.`);
    }
    
    if (!currentSheet) {
      throw new Error(`Current response sheet "${RESPONSE_SHEET_NAME}" not found.`);
    }
    
    // Delete current sheet and rename backup
    ss.deleteSheet(currentSheet);
    backupSheet.setName(RESPONSE_SHEET_NAME);
    
    Logger.log(`Successfully rolled back to backup: ${backupSheetName}`);
    return {
      success: true,
      message: `Successfully rolled back to backup: ${backupSheetName}`
    };
    
  } catch (error) {
    Logger.log(`Rollback failed: ${error.message}`);
    return {
      success: false,
      message: `Rollback failed: ${error.message}`
    };
  }
}

/**
 * Utility function to list available backups
 */
function listMigrationBackups() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheets = ss.getSheets();
    
    const backups = sheets
      .filter(sheet => sheet.getName().startsWith(`${RESPONSE_SHEET_NAME}_Backup_`))
      .map(sheet => ({
        name: sheet.getName(),
        created: sheet.getRange(1, 1).getNote() || "No timestamp available"
      }))
      .sort((a, b) => b.name.localeCompare(a.name)); // Most recent first
    
    return {
      success: true,
      backups: backups,
      count: backups.length
    };
    
  } catch (error) {
    Logger.log(`Error listing backups: ${error.message}`);
    return {
      success: false,
      message: `Error listing backups: ${error.message}`,
      backups: [],
      count: 0
    };
  }
}

/**
 * Test function to run migration on a small test dataset
 */
function testMigration() {
  try {
    Logger.log("Starting migration test...");
    
    // Run the migration
    const result = migrateCriticalBehaviorsData();
    
    if (result.success) {
      Logger.log("Migration test completed successfully!");
      Logger.log(`Details: ${JSON.stringify(result.details, null, 2)}`);
    } else {
      Logger.log(`Migration test failed: ${result.message}`);
    }
    
    return result;
    
  } catch (error) {
    Logger.log(`Migration test error: ${error.message}`);
    return {
      success: false,
      message: `Migration test error: ${error.message}`
    };
  }
}

/**
 * Converts flat evaluation data back to form data structure for editing
 * @param {Object} evaluationData - Flat evaluation data from spreadsheet
 * @returns {Object} Form data structure suitable for form population
 */
function convertEvaluationToFormData(evaluationData) {
  try {
    Logger.log(`=== CONVERT EVALUATION DEBUG START ===`);
    Logger.log(`Input evaluation data:`, JSON.stringify(evaluationData, null, 2));
    Logger.log(`Input data keys: ${Object.keys(evaluationData).join(', ')}`);
    Logger.log(`Input data type: ${typeof evaluationData}`);
    
    const formData = {};
    
    // Direct field mappings (these map directly from spreadsheet columns to form fields)
    const directMappings = {
      "Sap ID": "Sap ID",
      "Team Member": "Team Member", 
      "Team Leader": "Team Leader",
      "Operations Manager": "Operations Manager",
      "Line of Business": "Line of Business",
      "Interaction Date": "Interaction Date",
      "Customer's BAN": "Customer's BAN",
      "FLP Conversation ID": "FLP Conversation ID",
      "Listening Type": "Listening Type",
      "Call Duration": "Duration (minutes)",
      "Call Driver Level 1": "Call Driver Level 1",
      "Call Driver Level 2": "Call Driver Level 2",
      "Call Summary": "Call Summary",
      "Highlights": "Highlights",
      "Lowlights/Recommendation": "Lowlights/Recommendation",
      "Cancellation Root Cause": "Cancellation Root Cause",
      "Was there an unethical credit?": "Was there an unethical credit?"
    };
    
    // Map direct fields with special handling for dates
    Object.entries(directMappings).forEach(([spreadsheetColumn, formField]) => {
      if (evaluationData[spreadsheetColumn] !== undefined) {
        let value = evaluationData[spreadsheetColumn];
        
        // Special handling for date fields - ADD 1 DAY TO COMPENSATE FOR TIMEZONE OFFSET
        if (formField === "Interaction Date" && value) {
          Logger.log(`Processing Interaction Date: ${value} (type: ${typeof value})`);
          
          if (value instanceof Date) {
            // Add 1 day to compensate for timezone offset issue
            const adjustedDate = new Date(value);
            adjustedDate.setDate(adjustedDate.getDate() + 1);
            Logger.log(`Original date: ${value}, Adjusted date (added 1 day): ${adjustedDate}`);
            
            // Convert adjusted date to YYYY-MM-DD format
            try {
              value = adjustedDate.toLocaleDateString('en-CA');
              Logger.log(`Converted adjusted Date object to: ${value}`);
            } catch (error) {
              // Fallback to manual extraction if locale fails
              const year = adjustedDate.getFullYear();
              const month = String(adjustedDate.getMonth() + 1).padStart(2, '0');
              const day = String(adjustedDate.getDate()).padStart(2, '0');
              value = `${year}-${month}-${day}`;
              Logger.log(`Fallback conversion with adjustment: ${value}`);
            }
          } else if (typeof value === 'string' && value.trim()) {
            // For strings, check if already in correct format, otherwise try to parse and adjust
            const cleanDate = value.trim();
            if (cleanDate.match(/^\d{4}-\d{2}-\d{2}$/)) {
              // Parse the date string and add 1 day
              try {
                const dateObj = new Date(cleanDate + 'T00:00:00'); // Add time to avoid timezone issues
                dateObj.setDate(dateObj.getDate() + 1);
                value = dateObj.toLocaleDateString('en-CA');
                Logger.log(`String date adjusted by +1 day: ${cleanDate} → ${value}`);
              } catch (error) {
                Logger.log(`Failed to adjust string date, keeping original: ${cleanDate}`);
                value = cleanDate;
              }
            } else {
              // Try to parse, adjust, and convert
              try {
                const dateObj = new Date(cleanDate);
                if (!isNaN(dateObj.getTime())) {
                  dateObj.setDate(dateObj.getDate() + 1);
                  value = dateObj.toLocaleDateString('en-CA');
                  Logger.log(`Converted and adjusted string date to: ${value}`);
                } else {
                  Logger.log(`Invalid date string, keeping as-is: ${value}`);
                }
              } catch (error) {
                Logger.log(`Date parsing failed, keeping original: ${value}`);
              }
            }
          }
        }
        
        formData[formField] = value;
      }
    });
    
    // Map skill evaluation fields (radio buttons)
    const skillMappings = {
      "Willingness to Help": "Willingness to Help",
      "Warm and Energetic Greeting that Connects with the Customer": "Warm and Energetic Greeting that Connects with the Customer",
      "Listen & Retain": "Listen & Retain",
      "Secure the Customer's Identity Based on Department Guidelines": "Secure the Customer's Identity Based on Department Guidelines",
      "Probe Using a Combination of Open-ended and Closed-ended Questions": "Probe Using a Combination of Open-ended and Closed-ended Questions",
      "Investigate & Check All Applicable Resources": "Investigate & Check All Applicable Resources",
      "Offer Solutions / Explanations / Options": "Offer Solutions / Explanations / Options",
      "Begin Processing Transaction, Adding Notes and Updating Account Information when Appropriate": "Begin Processing Transaction, Adding Notes and Updating Account Information when Appropriate",
      "Confirm What You Know and Paraphrase Customer": "Confirm What You Know and Paraphrase Customer",
      "Check for Understanding": "Check for Understanding",
      "Lead with a Benefit Prior to Asking Questions and Securing the Account": "Lead with a Benefit Prior to Asking Questions and Securing the Account",
      "Summarize Solution": "Summarize Solution",
      "End on a High Note with a Personalized Close": "End on a High Note with a Personalized Close",
      "Reinforce the Extras You Did": "Reinforce the Extras You Did",
      "Early Bridge to Value When Appropriate": "Early Bridge to Value When Appropriate",
      "Later Bridge to Value when Appropriate / Value Proposition": "Later Bridge to Value when Appropriate / Value Proposition",
      "Lead Customer to Accept an Option or Solution": "Lead Customer to Accept an Option or Solution",
      "Check for Satisfaction": "Check for Satisfaction",
      "Overcome Objections if Necessary": "Overcome Objections if Necessary",
      "Acknowledge (Empathy When Needed)": "Acknowledge (Empathy When Needed)",
      "Learn from Your Customer": "Learn from Your Customer",
      "Handling Early Objections": "Handling Early Objections",
      "Customer Intimacy": "Customer Intimacy",
      "Value Proposition": "Value Proposition",
      "What type of bridging was provided?": "What type of bridging was provided?",
      "Right size": "Right size",
      "Use a tiered approach": "Use a tiered approach",
      "Sufficient probing (+3 Questions)": "Sufficient probing (+3 Questions)",
      "Speak to value (WHY TELUS, reality check and service superiority)": "Speak to value (WHY TELUS, reality check and service superiority)",
      "Loyalty statements implementation": "Loyalty statements implementation",
      "Investigation (Review bill, compare competitor offer, search notes on the account, take a look at usage)": "Investigation (Review bill, compare competitor offer, search notes on the account, take a look at usage)",
      "Was the agent able to retain the customer?": "Was the agent able to retain the customer?"
    };
    
    // Map skill fields
    Object.entries(skillMappings).forEach(([spreadsheetColumn, formField]) => {
      if (evaluationData[spreadsheetColumn] !== undefined) {
        formData[formField] = evaluationData[spreadsheetColumn];
      }
    });
    
    // Map Critical Behaviors (convert from 0/1 back to checkbox format)
    const criticalBehaviorMappings = {
      "CB_Abusive_Insult_Customer": "Abusive, insult the customer (disrespectful)",
      "CB_Appropriation_Sale_Another_Agent": "Appropriation of the sale from another agent",
      "CB_CCTS_Process_Not_Followed": "CCTS Process not Followed",
      "CB_Delayed_Response_Inbound": "Delayed Response to Inbound Contact",
      "CB_Didnt_Complete_Transaction_Promised": "Didn't complete transaction as promised",
      "CB_Excessive_Dead_Air_Hold_Idle": "Excessive Dead Air, Hold or Idle Time (with the clear intention of making the customer disconnect the call)",
      "CB_Failure_Follow_Processes_Onesource": "Failure to follow processes as per Onesource or use tools properly",
      "CB_Incorrect_Sales_Order_Processing": "Incorrect sales order processing",
      "CB_Incorrect_Use_Credits_Discounts": "Incorrect use of credits/discounts",
      "CB_Misleading_Information_Lying": "Misleading Information / Deliberately lying",
      "CB_Misleading_Product_Offer_Info": "Misleading product/offer information",
      "CB_Missed_Call_Back": "Missed Call Back",
      "CB_Modify_Account_Without_Consent": "Modify customer's account without consent",
      "CB_No_Authentication_Verification": "No Authentication/Verification",
      "CB_No_Response_Inbound_Contact": "No Response to Inbound Contact",
      "CB_No_Retention_Value_Generation": "No Retention Attempt or No Value Generation effort",
      "CB_Nonwork_Related_Distractions": "Nonwork related distractions",
      "CB_Overtalking_Interrupt_Customer": "Overtalking to the customer, interrupt the customer",
      "CB_Proactively_Transferring_Calls": "Proactively transferring calls",
      "CB_Staying_Line_Unnecessarily": "Staying on the line unnecessarily",
      "CB_System_Tools_Manipulation": "System/Tools Manipulation",
      "CB_Unjustified_Disconnection": "Unjustified disconnection of customer interaction",
      "CB_Unnecessary_Invalid_Consult_Transfer": "Unnecessary/Invalid Consult, Dispatch or Transfer",
      "CB_Unwillingness_To_Help": "Unwillingness to Help",
      "CB_Other_Behavior": "Other",
      "CB_Other_Behavior_Text": "CB_Other_Behavior_Text"
    };
    
    // Map Critical Behaviors (convert 1 to "on", 0 to unchecked)
    Object.entries(criticalBehaviorMappings).forEach(([spreadsheetColumn, formField]) => {
      if (evaluationData[spreadsheetColumn] !== undefined) {
        if (formField === "CB_Other_Behavior_Text") {
          // Text field - map directly
          formData[formField] = evaluationData[spreadsheetColumn] || "";
        } else {
          // Checkbox - convert 1 to "on", anything else is unchecked
          if (evaluationData[spreadsheetColumn] === 1 || evaluationData[spreadsheetColumn] === "1") {
            formData[formField] = "on";
          }
          // If 0 or falsy, don't set the field (unchecked state)
        }
      }
    });
    
    Logger.log(`=== CONVERSION SUMMARY ===`);
    Logger.log(`Form data created with ${Object.keys(formData).length} fields`);
    Logger.log(`Sample form data fields: ${Object.keys(formData).slice(0, 10).join(', ')}`);
    Logger.log(`Form data sample values:`, JSON.stringify(Object.fromEntries(Object.entries(formData).slice(0, 5)), null, 2));
    Logger.log(`=== CONVERT EVALUATION DEBUG END - SUCCESS ===`);
    
    return formData;
    
  } catch (error) {
    Logger.log(`❌ Error converting evaluation to form data: ${error.message}`);
    Logger.log(`Error stack: ${error.stack}`);
    Logger.log(`=== CONVERT EVALUATION DEBUG END - ERROR ===`);
    throw error;
  }
}

/**
 * Enhanced saveFormData function with metadata and validation
 */
function saveFormData(formData) {
  try {
    // Create enhanced data structure
    const enhancedData = createEnhancedFormData(formData);
    
    // Validate data
    const validation = validateEnhancedData(enhancedData);
    
    if (!validation.isValid) {
      Logger.log(`Validation warnings: ${validation.errors.join(', ')}`);
      // Continue saving even with warnings, but log them
    }
    
    // Get or create response sheet
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let responseSheet = ss.getSheetByName(RESPONSE_SHEET_NAME);
    
    if (!responseSheet) {
      responseSheet = ss.insertSheet(RESPONSE_SHEET_NAME);
      setupEnhancedHeaders(responseSheet);
    } else if (responseSheet.getLastRow() === 0) {
      setupEnhancedHeaders(responseSheet);
    }
    
    // Prepare row data (maintaining backward compatibility)
    const rowData = prepareRowData(enhancedData);
    
    // Append to sheet
    responseSheet.appendRow(rowData);
    
    // Send confirmation email to user
    try {
      sendSubmissionConfirmation(formData, enhancedData.recordId);
      Logger.log(`Confirmation email sent for Record ID: ${enhancedData.recordId}`);
    } catch (emailError) {
      Logger.log(`Email sending failed: ${emailError.message}`);
      // Don't fail the entire submission if email fails
    }
    
    // Log successful submission
    Logger.log(`Form submitted successfully. Record ID: ${enhancedData.recordId}, Quality Score: ${validation.qualityScore}%`);
    
    return {
      success: true,
      message: `Evaluation submitted successfully! Record ID: ${enhancedData.recordId}`,
      recordId: enhancedData.recordId,
      qualityScore: validation.qualityScore
    };
    
  } catch (error) {
    Logger.log(`Error in saveFormData: ${error.message}`);
    return {
      success: false,
      message: `Failed to save data. Error: ${error.message}`
    };
  }
}

// --- EMAIL NOTIFICATION FUNCTIONS ---

/**
 * Extracts a display name from an email address
 * @param {string} email - Email address to extract name from
 * @returns {string} Formatted display name
 */
function extractNameFromEmail(email) {
  try {
    if (!email || typeof email !== 'string') {
      return 'User';
    }
    
    const username = email.split('@')[0];
    // Convert john.doe or john_doe to John Doe
    return username
      .replace(/[._]/g, ' ')
      .split(' ')
      .map(word => word.charAt(0).toUpperCase() + word.slice(1).toLowerCase())
      .join(' ');
  } catch (error) {
    Logger.log(`Error extracting name from email: ${error.message}`);
    return 'User';
  }
}

/**
 * Generates Gmail-compatible HTML email content
 * @param {string} userName - Display name of the user
 * @param {string} recordId - Record ID of the evaluation
 * @param {Object} evaluationDetails - Key evaluation details
 * @param {boolean} isUpdate - Whether this is an update notification
 * @returns {string} HTML email content
 */
function generateEmailHTML(userName, recordId, evaluationDetails, isUpdate = false) {
  const actionText = isUpdate ? 'updated' : 'submitted';
  const actionPastTense = isUpdate ? 'Updated' : 'Submitted';
  
  return `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="utf-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>QA Evaluation ${actionPastTense}</title>
    </head>
    <body style="margin: 0; padding: 0; font-family: Arial, sans-serif; background-color: #f5f5f5;">
      <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%" style="background-color: #f5f5f5;">
        <tr>
          <td style="padding: 20px 0;">
            <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="600" style="margin: 0 auto; background-color: #ffffff; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
              
              <!-- Header -->
              <tr>
                <td style="padding: 30px 40px 20px 40px; text-align: center; background-color: #6366f1; border-radius: 8px 8px 0 0;">
                  <h1 style="margin: 0; color: #ffffff; font-size: 24px; font-weight: bold;">
                    ✅ QA New Hire Evaluation ${actionPastTense}
                  </h1>
                </td>
              </tr>
              
              <!-- Content -->
              <tr>
                <td style="padding: 30px 40px;">
                  <p style="margin: 0 0 25px 0; font-size: 16px; color: #333333; line-height: 1.5;">
                    Your QA evaluation has been successfully ${actionText} to the system.
                  </p>
                  
                  <!-- Evaluation Details Box -->
                  <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%" style="background-color: #f8fafc; border: 1px solid #e2e8f0; border-radius: 6px; margin: 25px 0;">
                    <tr>
                      <td style="padding: 20px;">
                        <h3 style="margin: 0 0 15px 0; font-size: 18px; color: #1f2937; font-weight: bold;">
                          Evaluation Details:
                        </h3>
                        
                        <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%">
                          <tr>
                            <td style="padding: 8px 0; font-size: 14px; color: #374151; font-weight: bold; width: 120px;">
                              • Record ID:
                            </td>
                            <td style="padding: 8px 0; font-size: 14px; color: #1f2937; font-family: monospace;">
                              ${recordId}
                            </td>
                          </tr>
                          <tr>
                            <td style="padding: 8px 0; font-size: 14px; color: #374151; font-weight: bold;">
                              • Team Member:
                            </td>
                            <td style="padding: 8px 0; font-size: 14px; color: #1f2937;">
                              ${evaluationDetails.teamMember || 'N/A'}
                            </td>
                          </tr>
                          <tr>
                            <td style="padding: 8px 0; font-size: 14px; color: #374151; font-weight: bold;">
                              • Sap ID:
                            </td>
                            <td style="padding: 8px 0; font-size: 14px; color: #1f2937;">
                              ${evaluationDetails.sapId || 'N/A'}
                            </td>
                          </tr>
                          <tr>
                            <td style="padding: 8px 0; font-size: 14px; color: #374151; font-weight: bold;">
                              • BAN:
                            </td>
                            <td style="padding: 8px 0; font-size: 14px; color: #1f2937;">
                              ${evaluationDetails.ban || 'N/A'}
                            </td>
                          </tr>
                          <tr>
                            <td style="padding: 8px 0; font-size: 14px; color: #374151; font-weight: bold;">
                              • FLP ID:
                            </td>
                            <td style="padding: 8px 0; font-size: 14px; color: #1f2937;">
                              ${evaluationDetails.flpId || 'N/A'}
                            </td>
                          </tr>
                        </table>
                      </td>
                    </tr>
                  </table>
                  
                  <p style="margin: 25px 0 0 0; font-size: 16px; color: #333333; line-height: 1.5;">
                    Your evaluation is now in the system and available for review.
                  </p>
                </td>
              </tr>
              
              <!-- Footer -->
              <tr>
                <td style="padding: 20px 40px 30px 40px; border-top: 1px solid #e5e7eb;">
                  <p style="margin: 0; font-size: 14px; color: #6b7280; line-height: 1.5;">
                    Best regards,<br>
                    <strong>QA Evaluation System</strong>
                  </p>
                  
                  <p style="margin: 15px 0 0 0; font-size: 12px; color: #9ca3af;">
                    This is an automated confirmation email. Please do not reply to this message.
                  </p>
                </td>
              </tr>
              
            </table>
          </td>
        </tr>
      </table>
    </body>
    </html>
  `;
}

/**
 * Sends a confirmation email to the user after form submission
 * @param {Object} formData - The submitted form data
 * @param {string} recordId - The generated record ID
 * @param {boolean} isUpdate - Whether this is an update notification
 */
function sendSubmissionConfirmation(formData, recordId, isUpdate = false) {
  try {
    // Get user email and name
    const userEmail = getCurrentUserEmail();
    const userName = extractNameFromEmail(userEmail);
    
    // Extract evaluation details for email
    const evaluationDetails = {
      teamMember: formData["Team Member"] || 'N/A',
      sapId: formData["Sap ID"] || 'N/A',
      ban: formData["Customer's BAN"] || 'N/A',
      flpId: formData["FLP Conversation ID"] || 'N/A'
    };
    
    // Generate email content
    const htmlBody = generateEmailHTML(userName, recordId, evaluationDetails, isUpdate);
    
    // Create subject line
    const actionText = isUpdate ? 'Updated' : 'Submitted';
    const subject = `✅ QA New Hire Evaluation ${actionText} Successfully - Record ID: ${recordId}`;
    
    // Send email
    MailApp.sendEmail({
      to: userEmail,
      subject: subject,
      htmlBody: htmlBody,
      name: "QA New Hire Evaluation System"
    });
    
    Logger.log(`Confirmation email sent to ${userEmail} for Record ID: ${recordId}`);
    
  } catch (error) {
    Logger.log(`Error sending confirmation email: ${error.message}`);
    throw error; // Re-throw to be caught by calling function
  }
}

/**
 * Test function to send a sample confirmation email
 * @returns {Object} Test result
 */
function testEmailNotification() {
  try {
    // Sample form data for testing
    const testFormData = {
      "Team Member": "John Doe",
      "Sap ID": "12345",
      "Customer's BAN": "987654321",
      "FLP Conversation ID": "FLP-2024-001234"
    };
    
    const testRecordId = "QA-09302024-TEST123";
    
    // Send test email
    sendSubmissionConfirmation(testFormData, testRecordId, false);
    
    return {
      success: true,
      message: "Test email sent successfully!",
      recordId: testRecordId
    };
    
  } catch (error) {
    Logger.log(`Test email failed: ${error.message}`);
    return {
      success: false,
      message: `Test email failed: ${error.message}`
    };
  }
}

/**
 * Diagnostic function to troubleshoot email issues
 * @returns {Object} Diagnostic information
 */
function diagnoseEmailSystem() {
  try {
    Logger.log("=== EMAIL SYSTEM DIAGNOSTICS START ===");
    
    const diagnostics = {
      timestamp: new Date().toISOString(),
      userEmail: null,
      emailPermissions: null,
      quotaInfo: null,
      testEmailResult: null,
      errors: []
    };
    
    // Test 1: Check user email
    try {
      const userEmail = getCurrentUserEmail();
      diagnostics.userEmail = userEmail;
      Logger.log(`✅ User email retrieved: ${userEmail}`);
    } catch (error) {
      diagnostics.errors.push(`Failed to get user email: ${error.message}`);
      Logger.log(`❌ Failed to get user email: ${error.message}`);
    }
    
    // Test 2: Check email quota
    try {
      const remainingQuota = MailApp.getRemainingDailyQuota();
      diagnostics.quotaInfo = {
        remainingQuota: remainingQuota,
        hasQuota: remainingQuota > 0
      };
      Logger.log(`✅ Email quota remaining: ${remainingQuota}`);
      
      if (remainingQuota === 0) {
        diagnostics.errors.push("No email quota remaining for today");
      }
    } catch (error) {
      diagnostics.errors.push(`Failed to check email quota: ${error.message}`);
      Logger.log(`❌ Failed to check email quota: ${error.message}`);
    }
    
    // Test 3: Try sending a simple test email
    try {
      if (diagnostics.userEmail) {
        MailApp.sendEmail({
          to: diagnostics.userEmail,
          subject: "🔧 Email System Diagnostic Test",
          body: `This is a diagnostic test email sent at ${new Date().toISOString()}.\n\nIf you receive this email, the basic email system is working.\n\nTest ID: DIAG-${Date.now()}`,
          name: "QA New Hire Evaluation System"
        });
        
        diagnostics.testEmailResult = {
          success: true,
          message: "Simple test email sent successfully"
        };
        Logger.log(`✅ Simple test email sent to ${diagnostics.userEmail}`);
      } else {
        diagnostics.testEmailResult = {
          success: false,
          message: "Cannot send test email - no user email available"
        };
      }
    } catch (error) {
      diagnostics.testEmailResult = {
        success: false,
        message: `Test email failed: ${error.message}`
      };
      diagnostics.errors.push(`Test email failed: ${error.message}`);
      Logger.log(`❌ Test email failed: ${error.message}`);
    }
    
    // Test 4: Check permissions
    try {
      // Try to access MailApp methods to check permissions
      const quota = MailApp.getRemainingDailyQuota();
      diagnostics.emailPermissions = {
        hasMailAppAccess: true,
        canCheckQuota: true
      };
      Logger.log(`✅ MailApp permissions verified`);
    } catch (error) {
      diagnostics.emailPermissions = {
        hasMailAppAccess: false,
        error: error.message
      };
      diagnostics.errors.push(`MailApp permission issue: ${error.message}`);
      Logger.log(`❌ MailApp permission issue: ${error.message}`);
    }
    
    // Summary
    const hasErrors = diagnostics.errors.length > 0;
    diagnostics.summary = {
      overallStatus: hasErrors ? "ISSUES_FOUND" : "HEALTHY",
      issueCount: diagnostics.errors.length,
      recommendations: []
    };
    
    if (diagnostics.quotaInfo && diagnostics.quotaInfo.remainingQuota === 0) {
      diagnostics.summary.recommendations.push("Wait until tomorrow for email quota to reset");
    }
    
    if (!diagnostics.userEmail) {
      diagnostics.summary.recommendations.push("Check Google Apps Script execution permissions");
    }
    
    if (diagnostics.testEmailResult && !diagnostics.testEmailResult.success) {
      diagnostics.summary.recommendations.push("Check spam folder and email filters");
    }
    
    if (diagnostics.errors.length === 0) {
      diagnostics.summary.recommendations.push("Email system appears healthy - check spam folder if emails not received");
    }
    
    Logger.log("=== EMAIL SYSTEM DIAGNOSTICS END ===");
    Logger.log(`Diagnostic summary: ${JSON.stringify(diagnostics.summary, null, 2)}`);
    
    return diagnostics;
    
  } catch (error) {
    Logger.log(`❌ Diagnostic function failed: ${error.message}`);
    return {
      success: false,
      error: error.message,
      message: "Diagnostic function encountered an error"
    };
  }
}

/**
 * Simple function to test basic email sending
 * @returns {Object} Test result
 */
function sendSimpleTestEmail() {
  try {
    const userEmail = getCurrentUserEmail();
    const testId = `TEST-${Date.now()}`;
    
    MailApp.sendEmail({
      to: userEmail,
      subject: "✅ Simple Email Test",
      body: `This is a simple test email.\n\nTest ID: ${testId}\nSent at: ${new Date().toISOString()}\n\nIf you receive this, basic email functionality is working.`,
      name: "QA Evaluation System"
    });
    
    Logger.log(`Simple test email sent to ${userEmail} with ID: ${testId}`);
    
    return {
      success: true,
      message: `Simple test email sent to ${userEmail}`,
      testId: testId,
      userEmail: userEmail
    };
    
  } catch (error) {
    Logger.log(`Simple test email failed: ${error.message}`);
    return {
      success: false,
      message: `Simple test email failed: ${error.message}`,
      error: error.toString()
    };
  }
}
