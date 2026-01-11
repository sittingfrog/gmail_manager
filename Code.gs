/**
 * @OnlyCurrentDoc
 */

/**
 * Creates a menu in the Google Sheet to open the web app.
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Gmail Manager')
      .addItem('Open Rule Manager', 'showSidebar')
      .addToUi();
}

/**
 * Shows the sidebar with the web app.
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Gmail Rule Manager');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Serves the main HTML page for the web app.
 * This allows the app to be deployed as a standalone web app.
 * @param {object} e The event parameter for a web app request.
 * @return {HtmlOutput} The HTML output for the web app.
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Gmail Rule Manager')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

/**
 * Retrieves the rules stored in script properties.
 * @return {Array<Object>} An array of rule objects.
 */
function getRules() {
  const userProperties = PropertiesService.getUserProperties();
  const rulesJson = userProperties.getProperty('gmailManagerRules');
  return rulesJson ? JSON.parse(rulesJson) : [];
}

/**
 * Saves a new rule to script properties.
 * @param {Object} rule The rule object to add.
 * @return {Array<Object>} The updated list of rules.
 */
function addRule(rule) {
  // Basic validation
  if (!rule.sender && !rule.subject) {
    throw new Error("A rule must have a sender or subject filter.");
  }
  const rules = getRules();
  rule.id = new Date().getTime().toString(); // Assign a unique ID
  rule.attachmentActions = rule.attachmentActions || []; // Ensure attachmentActions array exists
  rules.push(rule);
  const userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('gmailManagerRules', JSON.stringify(rules));
  return rules;
}

function saveRules(rules) {
  const userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('gmailManagerRules', JSON.stringify(rules));
  return rules;
}

/**
 * Deletes a rule based on its ID.
 * @param {string} ruleId The ID of the rule to delete.
 * @return {Array<Object>} The updated list of rules.
 */
function deleteRule(ruleId) {
  let rules = getRules();
  rules = rules.filter(rule => rule.id !== ruleId);
  return saveRules(rules);
}

/**
 * Adds an attachment action to an existing rule.
 * @param {string} ruleId The ID of the parent rule.
 * @param {Object} action The attachment action object to add.
 * @return {Array<Object>} The updated list of rules.
 */
function addAttachmentAction(ruleId, action) {
  const rules = getRules();
  const rule = rules.find(r => r.id === ruleId);
  if (!rule) throw new Error("Rule not found.");
  action.id = new Date().getTime().toString();
  rule.attachmentActions.push(action);
  return saveRules(rules);
}

/**
 * Deletes an attachment action from a rule.
 * @param {string} ruleId The ID of the parent rule.
 * @param {string} actionId The ID of the action to delete.
 * @return {Array<Object>} The updated list of rules.
 */
function deleteAttachmentAction(ruleId, actionId) {
  const rules = getRules();
  const rule = rules.find(r => r.id === ruleId);
  if (rule) {
    rule.attachmentActions = rule.attachmentActions.filter(a => a.id !== actionId);
  }
  return saveRules(rules);
}

/**
 * Scans for unread emails from a specific sender, saves specified attachments
 * to Google Drive, and marks the emails as read.
 * This function is intended to be run on a time-driven trigger.
 */
function processEmails() {
  const rules = getRules();
  if (rules.length === 0) {
    console.log("No rules to process.");
    return;
  }

  rules.forEach(rule => {
    // Build the search query from the rule
    let query = 'is:unread';
    if (rule.sender) query += ` from:"${rule.sender}"`;
    if (rule.subject) query += ` subject:("${rule.subject}")`;
    if (rule.attachmentActions && rule.attachmentActions.length > 0) query += ' has:attachment';

    try {
      const threads = GmailApp.search(query, 0, 10); // Process up to 10 threads per run to avoid timeout
      threads.forEach(thread => {
        const messages = thread.getMessages();
        messages.forEach(message => {
          if (message.isUnread() && rule.attachmentActions.length > 0) {
            rule.attachmentActions.forEach(action => {
              processAttachments(message, action);
            }); // This was the line with the missing parenthesis
          }
        });
        thread.markRead();
      });
    } catch (e) {
      console.error(`Error processing rule for sender "${rule.sender}": ${e.toString()}`);
    }
  });
}

/**
 * Processes attachments for a given message based on a rule.
 * @param {GoogleAppsScript.Gmail.GmailMessage} message The Gmail message.
 * @param {Object} action The attachment action to apply.
 */
function processAttachments(message, action) {
  const driveFolder = DriveApp.getFolderById(action.driveFolderId);
  if (!driveFolder) {
    console.error(`Drive folder with ID ${action.driveFolderId} not found.`);
    return;
  }

  const attachments = message.getAttachments();
  attachments.forEach(attachment => {
    const originalName = attachment.getName();
    const attachmentNameFilter = action.attachmentName || '*';

    // Check if the attachment name matches the action's specification
    const nameMatches = (attachmentNameFilter === '*' || attachmentNameFilter === originalName);

    if (nameMatches) {
      const outputName = getOutputFileName(action.outputFileName, originalName, message.getSubject(), message.getDate());
      
      // Check for duplicates before creating the file
      const existingFiles = driveFolder.getFilesByName(outputName);
      if (existingFiles.hasNext()) {
        console.log(`File "${outputName}" already exists in folder. Skipping.`);
        return;
      }

      driveFolder.createFile(attachment.copyBlob()).setName(outputName);
      console.log(`Saved attachment "${originalName}" as "${outputName}" to Drive.`);
    }
  });
}

/**
 * Generates the output file name based on a template.
 * @param {string} template The user-defined template string.
 * @param {string} originalName The original name of the attachment.
 * @param {string} subject The email subject.
 * @param {Date} date The date of the email.
 * @return {string} The formatted file name.
 */
function getOutputFileName(template, originalName, subject, date) {
  if (!template) {
    return originalName;
  }

  const yyyy = date.getFullYear();
  const mm = ('0' + (date.getMonth() + 1)).slice(-2);
  const dd = ('0' + date.getDate()).slice(-2);
  const dateString = `${yyyy}-${mm}-${dd}`;

  return template
    .replace(/{original_name}/g, originalName)
    .replace(/{subject}/g, subject.replace(/[/\\?%*:|"<>]/g, '-')) // Sanitize subject for file name
    .replace(/{date}/g, dateString);
}