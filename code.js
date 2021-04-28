// This Google Sheets script will post to a slack channel when a user submits data to a Google Forms Spreadsheet
// View the README for installation instructions. Don't forget to add the required slack information below.

// Source: https://github.com/markfguerra/google-forms-to-slack

/////////////////////////
// Begin customization //
/////////////////////////

// Alter this to match the incoming webhook url provided by Slack
var teamsIncomingWebhookUrl = 'https://o365coloradoedu.webhook.office.com/webhookb2/card-here';

///////////////////////
// End customization //
///////////////////////

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
  // Test code. uncomment to debug in Google Script editor
  // if (typeof e === "undefined") {
  //    e = {namedValues: {"Question1": ["answer1"], "Question2" : ["answer2"]}};
  //    messagePretext = "Debugging our Sheets to Slack integration";
  // };

  var fields = makeFields(e.values);
  var toSend = fields.join('\n');
  
  var payload = {
    "title":"Add the following employee to Citrix",
    "text": toSend
  };

  var options = {
    'method': 'post',
    'payload': JSON.stringify(payload)
  };

  var response = UrlFetchApp.fetch(teamsIncomingWebhookUrl, options);
}

// Creates an array of Slack fields containing the questions and answers
var makeFields = function(values) {
  var fields = [];

  var columnNames = getColumnNames();

  for (var i = 0; i < columnNames.length; i++) {
    var colName = columnNames[i];
    var val = values[i];
    var thing = colName.concat(' ', val)
    fields.push(thing);
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
