/**
 * gmail_to_sheets.gs
  * ------------------
   * Google Apps Script: reads Gmail messages by label and logs
    * key fields (date, sender, subject, snippet) into a Google Sheet.
     *
      * Setup:
       *   1. Open Google Sheets > Extensions > Apps Script
        *   2. Paste this script
         *   3. Update LABEL_NAME and SHEET_NAME below
          *   4. Run setupTrigger() once to schedule automatic runs
           *
            * Permissions needed: Gmail (read), Sheets (read/write)
             */

             var LABEL_NAME = 'to-log';      // Gmail label to read from
             var SHEET_NAME = 'Email Log';   // Sheet tab name to write to
             var MAX_THREADS = 50;           // Max threads to process per run


             /**
              * Main function: fetch labeled emails and log them to the sheet.
               * Marks threads as processed by removing the label after logging.
                */
                function logEmailsToSheet() {
                  var label = GmailApp.getUserLabelByName(LABEL_NAME);
                    if (!label) {
                        Logger.log('Label not found: ' + LABEL_NAME);
                            return;
                              }

                                var threads = label.getThreads(0, MAX_THREADS);
                                  if (threads.length === 0) {
                                      Logger.log('No threads found with label: ' + LABEL_NAME);
                                          return;
                                            }

                                              var sheet = getOrCreateSheet(SHEET_NAME);
                                                ensureHeaders(sheet);

                                                  var rows = [];

                                                    for (var i = 0; i < threads.length; i++) {
                                                        var thread = threads[i];
                                                            var messages = thread.getMessages();
                                                                var firstMessage = messages[0];

                                                                    rows.push([
                                                                          Utilities.formatDate(firstMessage.getDate(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'),
                                                                                firstMessage.getFrom(),
                                                                                      firstMessage.getTo(),
                                                                                            thread.getFirstMessageSubject(),
                                                                                                  firstMessage.getPlainBody().substring(0, 300).replace(/\n/g, ' '),
                                                                                                        thread.getPermalink()
                                                                                                            ]);
                                                                                                            
                                                                                                                // Remove label so it isn't processed again
                                                                                                                    thread.removeLabel(label);
                                                                                                                      }
                                                                                                                      
                                                                                                                        if (rows.length > 0) {
                                                                                                                            var lastRow = sheet.getLastRow();
                                                                                                                                sheet.getRange(lastRow + 1, 1, rows.length, rows[0].length).setValues(rows);
                                                                                                                                    Logger.log('Logged ' + rows.length + ' thread(s) to sheet.');
                                                                                                                                      }
                                                                                                                                      }
                                                                                                                                      
                                                                                                                                      
                                                                                                                                      /**
                                                                                                                                       * Returns the sheet by name, or creates it if it doesn't exist.
                                                                                                                                        */
                                                                                                                                        function getOrCreateSheet(name) {
                                                                                                                                          var ss = SpreadsheetApp.getActiveSpreadsheet();
                                                                                                                                            var sheet = ss.getSheetByName(name);
                                                                                                                                              if (!sheet) {
                                                                                                                                                  sheet = ss.insertSheet(name);
                                                                                                                                                    }
                                                                                                                                                      return sheet;
                                                                                                                                                      }
                                                                                                                                                      
                                                                                                                                                      
                                                                                                                                                      /**
                                                                                                                                                       * Adds header row if the sheet is empty.
                                                                                                                                                        */
                                                                                                                                                        function ensureHeaders(sheet) {
                                                                                                                                                          if (sheet.getLastRow() === 0) {
                                                                                                                                                              sheet.appendRow(['Date', 'From', 'To', 'Subject', 'Preview', 'Link']);
                                                                                                                                                                  sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
                                                                                                                                                                    }
                                                                                                                                                                    }
                                                                                                                                                                    
                                                                                                                                                                    
                                                                                                                                                                    /**
                                                                                                                                                                     * Run this once to set up an hourly trigger.
                                                                                                                                                                      */
                                                                                                                                                                      function setupTrigger() {
                                                                                                                                                                        // Delete any existing triggers for this function
                                                                                                                                                                          var triggers = ScriptApp.getProjectTriggers();
                                                                                                                                                                            for (var i = 0; i < triggers.length; i++) {
                                                                                                                                                                                if (triggers[i].getHandlerFunction() === 'logEmailsToSheet') {
                                                                                                                                                                                      ScriptApp.deleteTrigger(triggers[i]);
                                                                                                                                                                                          }
                                                                                                                                                                                            }
                                                                                                                                                                                            
                                                                                                                                                                                              // Create a new hourly trigger
                                                                                                                                                                                                ScriptApp.newTrigger('logEmailsToSheet')
                                                                                                                                                                                                    .timeBased()
                                                                                                                                                                                                        .everyHours(1)
                                                                                                                                                                                                            .create();
                                                                                                                                                                                                            
                                                                                                                                                                                                              Logger.log('Hourly trigger created for logEmailsToSheet.');
                                                                                                                                                                                                              }
