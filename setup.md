Install necessary stuff 
https://tools.slack.dev/deno-slack-sdk/guides/getting-started



Create a Slack App:

Go to https://api.slack.com/apps and create a new app
Enable Socket Mode
Add permissions: reactions:read, chat:write, users:read
Install to your workspace and copy the tokens to your .env file


Register an Azure App:

Register an app in Azure Active Directory
Add Microsoft Graph API permissions for calendar management
Copy the client ID, secret, and tenant ID to your .env file


Install and run:
npm install @slack/bolt axios @azure/msal-node dotenv
node app.js