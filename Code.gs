/**
 * @OnlyCurrentDoc
 *
 * The above comment directs Apps Script to limit the scope of file
 * access for this script to only the current document containing the script.
 */

// =================================================================
// == CONFIGURATION                                             ==
// =================================================================
// Please update these values to match your needs.

/**
 * The email address of the sender to look for.
 * @type {string}
 */
const SENDER_EMAIL = "Plaud.ai";

/**
 * The name of the summary attachment.
 * @type {string}
 */
const SUMMARY_ATTACHMENT_NAME = "summary.txt";

/**
 * The name of the transcript attachment.
 * @type {string}
 */
const TRANSCRIPT_ATTACHMENT_NAME = "transcript.txt";

/**
 * The ID of the Google Drive folder where summary files will be saved.
 * To get the folder ID, open the folder in Google Drive and copy the
 * last part of the URL.
 * e.g., for https://drive.google.com/drive/folders/12345ABCDE, the ID is "12345ABCDE".
 * @type {string}
 */
const SUMMARY_FOLDER_ID = "1Tg9toFVe-DbCEeBWMx3Uoo3h70O0MwOT";

/**
 * The ID of the Google Drive folder where transcript files will be saved.
 * @type {string}
 */
const TRANSCRIPT_FOLDER_ID = "1N_iTjBq0VXbx9whnHAnFyE0j1cYrmDQo";


/**
 * Scans for unread emails from a specific sender, saves specified attachments
 * to Google Drive, and marks the emails as read.
 * This function is intended to be run on a time-driven trigger.
 */
function processEmails() {
  const summaryFolder = DriveApp.getFolderById(SUMMARY_FOLDER_ID);
  const transcriptFolder = DriveApp.getFolderById(TRANSCRIPT_FOLDER_ID);

  // Search for unread emails from the specified sender
  const query = `from:${SENDER_EMAIL} is:unread`;
  const threads = GmailApp.search(query);

  threads.forEach(thread => {
    const messages = thread.getMessages();
    messages.forEach(message => {
      // Only process messages that are unread
      if (message.isUnread()) {
        const attachments = message.getAttachments();
        attachments.forEach(attachment => {
          const attachmentName = attachment.getName();
          if (attachmentName === SUMMARY_ATTACHMENT_NAME) {
            summaryFolder.createFile(attachment.copyBlob());
          } else if (attachmentName === TRANSCRIPT_ATTACHMENT_NAME) {
            transcriptFolder.createFile(attachment.copyBlob());
          }
        });
      }
    });
    // Mark the entire thread as read to avoid re-processing
    thread.markRead();
  });
}