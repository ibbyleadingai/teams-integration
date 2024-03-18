// Import necessary libraries
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

        // Here we assume the entire response from the backend is structured for direct use
        // Adjust based on your backend response structure
        const messages = response.data.choices.flatMap(choice => choice.messages);
        for (const msg of messages) {
            if (msg.role === 'assistant') {
                await context.sendActivity(msg.content); // Sending each message back to the user
            }
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
