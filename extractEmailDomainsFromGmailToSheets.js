//  ____          _____  _____   __          __
// |  _ \   /\   |  __ \|  __ \ /\ \        / /
// | |_) | /  \  | |__) | |__) /  \ \  /\  / / 
// |  _ < / /\ \ |  _  /|  ___/ /\ \ \/  \/ /  
// | |_) / ____ \| | \ \| |  / ____ \  /\  /   
// |____/_/    \_\_|  \_\_| /_/    \_\/  \/    
// 
// This script automatically creates a Google Sheets document containing a list of all domains from which the Gmail user has received emails.
// The maximum number of email threads (email domains) processed in one run is 500.
// The script processes 500 threads (email domains) every 5 minutes.
// To run this script, paste it into the Google Apps Script Editor and select the createTimeTriggerEvery5Minutes function as the entry point to add a trigger.
// The script will execute every 5 minutes and will continue until all emails have been processed or until the daily limit of Gmail API queries is exhausted.

// 🛠️ How to Run It:
// 1.) Log into your Google account and create a new Google Sheet.
// 2.) From the top menu, select Extensions > Apps Script.
// 3.) In the newly opened tab, the Google Apps Script editor will open. Erase the current content and paste the script content from this link: https://gist.github.com/barpaw/3792691a62b3877149795851fcb177c6
// 4.) Save the project and click Run.
// 5.) A message will appear indicating the need to grant permission to the app to access your data (permissions relate to Gmail, Google Sheets, and background operations).
// 6.) After authorization, the script will execute, and an automatic trigger will be created, which will cyclically update our sheet with data (domains) every 5 minutes.
// 7.) Emails are analyzed in batches of 500 threads, so after a few minutes, you should see the first batch of domains in your sheet.
// 8.) ☕ The solution operates in the cloud, so you can close the open tabs and return to your sheet later. Once all emails and their associated sender domains have been processed, the script will automatically delete the trigger and cease operation.

function createTimeTriggerEvery5Minutes() {
    // Creates a trigger that runs the 'main' function every 5 minutes
    ScriptApp.newTrigger("main")
        .timeBased()
        .everyMinutes(5)
        .create();
}

function main() {
    // Main script logic, processing emails

    Logger.log("Invoked function main()");

    var startIndex = PropertiesService.getScriptProperties().getProperty('startIndex'); // The index from which email messages will be retrieved.
    var count = 500; // Number of emails to process (500 is max in one run)

    Logger.log("Value of property startIndex: " + startIndex);
    Logger.log("Value of variable count: " + count);

    if (!startIndex) {
        startIndex = 0;
    }

    try {
        var threadsBatchCount = extractAllEmailDomainsInRange(parseInt(startIndex), parseInt(count));

        if (threadsBatchCount == count) {
            startIndex = parseInt(startIndex) + count;
            PropertiesService.getScriptProperties().setProperty('startIndex', startIndex.toString());
        } else {
            Logger.log("There are no more emails.");
            // remove trigger
            deleteTrigger();

            Logger.log("Done.");
        }

    } catch (e) {

        Logger.log("Error | main() yielded an error: " + e);
    }
}

function deleteTrigger() {
    // Function to remove the trigger after completing the task

    var triggers = ScriptApp.getProjectTriggers();

    for (var i = 0; i < triggers.length; i++) {

        if (triggers[i].getHandlerFunction() === "main") {

            ScriptApp.deleteTrigger(triggers[i]);
            Logger.log("Deleted trigger for function: main().");
        }
    }
}

function extractAllEmailDomainsInRange(startIndex, count) {
    // Function to extract domains from emails and save them in the sheet

    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getActiveSheet();

    Logger.log("Invoked function extractAllEmailDomainsInRange(" + startIndex + ", " + count + ").");

    // Create the sheet if it doesn't exist
    if (!sheet) {
        sheet = spreadsheet.insertSheet(sheetName);
        Logger.log("Sheet Info: Creating new sheet.");
        if (!sheet) {
            Logger.log("Sheet Info: Unable to create new sheet.");
            return;
        }
    } else {
        Logger.log("Sheet Info: Sheet exists.");
    }

    // Fetch threads in the specified range
    var threads = GmailApp.getInboxThreads(startIndex, count);
    var threadIndex = startIndex;
    var dataToInsert = []; // List to store domain and index

    for (var i = 0; i < threads.length; i++) {
        Logger.log("Progress: " + i + "/" + threads.length + " | " + threadIndex);
        var message = threads[i].getMessages()[0];

        var fromAddress = message.getFrom();

        var domain = fromAddress.substring(fromAddress.lastIndexOf("@") + 1).toLowerCase().replace(/>/g, "");

        // Store domain and thread index in list
        dataToInsert.push([domain, threadIndex]);
        threadIndex++;
    }

    // Insert data
    if (dataToInsert.length > 0) {
        var startRow = sheet.getLastRow() + 1;
        var range = sheet.getRange(startRow, 1, dataToInsert.length, 2);
        range.setValues(dataToInsert);
    }

    Logger.log("Current iteration: Done.");

    return threads.length;
}
