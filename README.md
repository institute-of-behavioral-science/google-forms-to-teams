# Google Forms To Microsoft Teams

With thanks to https://github.com/markfguerra and his Slack notifier here: https://github.com/markfguerra/google-forms-to-slack

Do you have a Google Form? Did you ever wish that you could get a Teams message when someone submits the Google Form? Great, because that is exactly what this script does.

## Technology Info
 - This script is a web service client to the Teams Connector Webhook API; it posts Teams messsages using HTTP Post.
 - This is a Google Apps script. That is, a JavaScript environment that can automate Google Drive products such as Google Sheets and Google Forms. It also comes with a neat browser based IDE. https://developers.google.com/apps-script/

## Setup
You'll need a few things
- A Google Sheet attached to your Google Form.
- A connector URL for the channel in Teams that you want to post into: https://docs.microsoft.com/en-us/microsoftteams/platform/webhooks-and-connectors/how-to/connectors-using#:~:text=In%20Microsoft%20Teams%2C%20choose%20More,the%20webhook%2C%20and%20choose%20Create.
- Note the URL that the above gives you

Procedure
 - Make sure you're logged in as the owner of the Google Forms Spreadsheet. If you're logged into any other Google accounts, this won't work, so I suggest using an incognito or private browsing window.
 - Open your Google Forms Spreadsheet.
 - In the menu, click on "Tools" -> "Script Editor".
 - Paste the code.js script into the Script Editor.
 - Edit the code you pasted to include the Teams webhook URL from Setup above in the customizations block; change the variable `teamsIncomingWebhookUrl`
 - Change the title of the cards in the `var cardTitle = ` line.
 - Set up the event triggers by running the `initialize()` function. In the Script Editor's menu bar, select the function `initialize` and click Run. Agree to any permission requests.
 - You're done! Try it out by submitting a response to your Google Form. If successful, you'll see a new message in your Teams channel.
