// This Google Sheets script will post to a Teams channel when a user submits data to a Google Forms Spreadsheet
// View the README for installation instructions. Don't forget to add the required Teams information below.

// Originally based on: https://github.com/markfguerra/google-forms-to-slack

/////////////////////////
// Begin customization //
/////////////////////////

// Alter this to match the incoming webhook url provided by Slack
var teamsIncomingWebhookUrl = 'https://o365coloradoedu.webhook.office.com/webhookb2/card-here';
var cardTitle = "Teams Card Title"

///////////////////////
// End customization //
///////////////////////
// This Google Sheets script will post to a Teams channel when a user submits data to a Google Forms Spreadsheet
// View the README for installation instructions. Don't forget to add the required slack information below.

// In the Script Editor, run initialize() at least once to make your code execute on form submit
function initialize() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i in triggers) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  ScriptApp.newTrigger("submitValuesToTeams")
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onFormSubmit()
    .create();
}

// Running the code in initialize() will cause this function to be triggered this on every Form Submit
function submitValuesToTeams(e) {
  var fields = makeFields(e.values);
  var toSend = fields.join('\n-  ');
  var toSend = '-  '.concat('', toSend);
  
  var payload = {
    "title": cardTitle,
    "text": toSend
  };

  var options = {
    'method': 'post',
    'payload': JSON.stringify(payload)
  };

  var response = UrlFetchApp.fetch(teamsIncomingWebhookUrl, options);
}

// Creates an array of Teams fields containing the questions and answers
var makeFields = function(values) {
  var fields = [];

  var columnNames = getColumnNames();

  for (var i = 0; i < columnNames.length; i++) {
    var colName = columnNames[i];
    var val = values[i];
    if (typeof val !== "undefined") {
      var finalString = "**";
      finalString += colName;
      finalString += ":** ";
      finalString += val;
      fields.push(finalString);
    }
  }

  return fields;
}

// Extracts the column names from the first row of the spreadsheet
var getColumnNames = function() {
  var sheet = SpreadsheetApp.getActiveSheet();

  // Get the header row using A1 notation
  var headerRow = sheet.getRange("1:1");

  // Extract the values from it
  var headerRowValues = headerRow.getValues()[0];

  return headerRowValues;
}
