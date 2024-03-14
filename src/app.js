// Import necessary libraries
const { MemoryStorage } = require("botbuilder");
const axios = require('axios');
const { Application } = require("@microsoft/teams-ai");
const config = require("./config");

if (!config.openAIKey || !config.openAIAssistantId) {
  throw new Error("Missing OPENAI_API_KEY or OPENAI_ASSISTANT_ID in the configuration.");
}

// Validate necessary configurations
if (!config.backendEndpoint) {
    throw new Error("Missing backend endpoint in the configuration.");
}

// Define memory storage for Teams application
const storage = new MemoryStorage();
const app = new Application({
    storage
});

// Send message to your Python Quart backend
async function sendMessageToBackend(context, message) {
    console.log(`Sending message to backend: ${message}`); // Logging the message being sent
    try {
        const response = await axios.post(config.backendEndpoint, {
            messages: [{ role: 'user', content: message }]
        });
        console.log(`Received response from backend: ${JSON.stringify(response.data)}`); // Logging the response
        const replyText = response.data; // Adjust based on your backend response structure
        await context.sendActivity(replyText);
    } catch (error) {
        console.error('Error sending message to backend:', error);
        console.log(`Error details: ${error.message}`);
        if (error.response) {
            // The request was made and the server responded with a status code
            // that falls out of the range of 2xx
            console.log(error.response.data);
            console.log(error.response.status);
            console.log(error.response.headers);
        } else if (error.request) {
            // The request was made but no response was received
            console.log(error.request);
        } else {
            // Something happened in setting up the request that triggered an Error
            console.log('Error', error.message);
        }
        // Send error message back to user
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

