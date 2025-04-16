const { App } = require('@slack/bolt');
const axios = require('axios');
const msal = require('@azure/msal-node');

// Initialize configuration
require('dotenv').config();

// Slack App initialization
const app = new App({
  token: process.env.SLACK_BOT_TOKEN,
  signingSecret: process.env.SLACK_SIGNING_SECRET,
  socketMode: true,
  appToken: process.env.SLACK_APP_TOKEN
});

// Microsoft Authentication configuration
const msalConfig = {
  auth: {
    clientId: process.env.MS_CLIENT_ID,
    clientSecret: process.env.MS_CLIENT_SECRET,
    authority: `https://login.microsoftonline.com/${process.env.MS_TENANT_ID}`
  }
};

const msalClient = new msal.ConfidentialClientApplication(msalConfig);

// Store for meetings data - in production, use a database
const meetingsStore = {};

// Listen for message events containing Teams meeting links
app.message(/https:\/\/teams\.microsoft\.com\/l\/meetup-join\/[^\s]+/i, async ({ message, context, say }) => {
  try {
    // Extract meeting link
    const meetingLinkRegex = /(https:\/\/teams\.microsoft\.com\/l\/meetup-join\/[^\s]+)/i;
    const meetingLink = message.text.match(meetingLinkRegex)[1];

    // Extract meeting ID from link (simplified - actual parsing depends on link format)
    const meetingId = extractMeetingIdFromLink(meetingLink);
    
    if (!meetingId) {
      await say({ text: "Couldn't parse the Teams meeting link.", thread_ts: message.ts });
      return;
    }

    // Store meeting info with original message timestamp as key
    meetingsStore[message.ts] = {
      meetingId,
      meetingLink,
      createdBy: message.user,
      participants: new Set()
    };

    // Add reaction to indicate this message can be reacted to for joining
    await app.client.reactions.add({
      token: context.botToken,
      name: "calendar",
      channel: message.channel,
      timestamp: message.ts
    });

    await say({
      text: "I've detected a Teams meeting link! React with :raised_hand: to be added to this meeting.",
      thread_ts: message.ts
    });
  } catch (error) {
    console.error("Error handling Teams meeting link:", error);
  }
});

// Listen for reaction added events
app.event('reaction_added', async ({ event, context, client }) => {
  try {
    // Check if this is a reaction to a message with a stored meeting
    const messageTs = event.item.ts;
    if (!meetingsStore[messageTs]) {
      return; // Not a tracked meeting message
    }

    // Check if the reaction is the one we're looking for
    if (event.reaction !== 'raised_hand') {
      return; // Not the target reaction
    }

    // Get user info
    const userResult = await client.users.info({
      user: event.user
    });

    // Check if user is already in the meeting
    const meetingInfo = meetingsStore[messageTs];
    if (meetingInfo.participants.has(event.user)) {
      await client.chat.postEphemeral({
        channel: event.item.channel,
        user: event.user,
        text: "You're already added to this meeting!"
      });
      return;
    }

    // Add user to the meeting via Microsoft Graph API
    const userEmail = userResult.user.profile.email;
    if (!userEmail) {
      await client.chat.postEphemeral({
        channel: event.item.channel,
        user: event.user,
        text: "Couldn't find your email address. Make sure your Slack profile has an email set."
      });
      return;
    }

    // Add user to Teams meeting
    const success = await addUserToTeamsMeeting(meetingInfo.meetingId, userEmail);
    
    if (success) {
      // Update our store
      meetingInfo.participants.add(event.user);
      
      // Notify user
      await client.chat.postMessage({
        channel: event.item.channel,
        thread_ts: messageTs,
        text: `<@${event.user}> has been added to the Teams meeting!`
      });
      
      // Send the user a direct message with the meeting link
      await client.chat.postMessage({
        channel: event.user,
        text: `You've been added to a Teams meeting. Here's the link: ${meetingInfo.meetingLink}`
      });
    } else {
      await client.chat.postEphemeral({
        channel: event.item.channel,
        user: event.user,
        text: "Sorry, there was a problem adding you to the meeting. Please try again or contact the meeting organizer."
      });
    }
  } catch (error) {
    console.error("Error handling reaction event:", error);
  }
});

/**
 * Extract meeting ID from Teams meeting link
 * Note: This is a simplified implementation and may need to be adjusted based on actual Teams URL format
 */
function extractMeetingIdFromLink(link) {
  try {
    // Extract the meeting ID from the URL parameters
    const url = new URL(link);
    const meetingId = url.searchParams.get('meetingId');
    return meetingId;
  } catch (error) {
    console.error("Error extracting meeting ID:", error);
    return null;
  }
}

/**
 * Add a user to a Teams meeting using Microsoft Graph API
 */
async function addUserToTeamsMeeting(meetingId, userEmail) {
  try {
    // Get access token for Microsoft Graph API
    const tokenResponse = await msalClient.acquireTokenByClientCredential({
      scopes: ['https://graph.microsoft.com/.default']
    });
    
    if (!tokenResponse || !tokenResponse.accessToken) {
      console.error("Failed to acquire Microsoft Graph token");
      return false;
    }
    
    // Add user as an attendee to the meeting
    // Note: This is a simplified version - actual implementation depends on your Teams/Outlook environment
    const response = await axios.patch(
      `https://graph.microsoft.com/v1.0/me/events/${meetingId}`,
      {
        attendees: [{
          emailAddress: {
            address: userEmail
          },
          type: "required"
        }]
      },
      {
        headers: {
          'Authorization': `Bearer ${tokenResponse.accessToken}`,
          'Content-Type': 'application/json'
        }
      }
    );
    
    return response.status >= 200 && response.status < 300;
  } catch (error) {
    console.error("Error adding user to Teams meeting:", error);
    return false;
  }
}

// Start the app
(async () => {
  await app.start(process.env.PORT || 3000);
  console.log('⚡️ Bolt app is running!');
})();
