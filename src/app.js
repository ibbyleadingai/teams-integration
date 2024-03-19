const { MemoryStorage } = require("botbuilder");
const axios = require('axios');
const { Application } = require("@microsoft/teams-ai");
const config = require("./config");

// Define memory storage for Teams application
const storage = new MemoryStorage();
const app = new Application({ storage });

// Send message to your Python Quart backend specifically for Teams
async function sendMessageToBackend(context, message) {
    try {
        const response = await axios.post(config.backendEndpoint, { // Ensure this is the correct endpoint
            messages: [{ role: 'user', content: message }]
        });

        // Assuming the response from your backend now directly contains an array of messages 
        // Adjust based on your backend response structure
        const messages = response.data.messages; // Updated to directly use 'messages' from response
        for (const msg of messages) {
            await context.sendActivity(msg.text); // Sending each message back to the user
        }
    } catch (error) {
        console.error('Error sending message to backend:', error);
        await context.sendActivity("There was an error processing your message. Please try again.");
    }
}

// Main message handler
app.message(async (context) => {
    const userMessage = context.activity.text;
    await sendMessageToBackend(context, userMessage);
});

// Reset conversation state
app.message("/reset", async (context) => {
    await context.sendActivity("Okay, let's start over.");
});

module.exports = app;