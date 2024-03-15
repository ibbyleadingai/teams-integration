// Import necessary libraries
const { MemoryStorage } = require("botbuilder");
const axios = require('axios');
const { Application } = require("@microsoft/teams-ai");
const config = require("./config");

if (!config.openAIKey || !config.openAIAssistantId) {
  throw new Error("Missing OPENAI_API_KEY or OPENAI_ASSISTANT_ID in the configuration.");
}

// Validate necessary configurations
// if (!config.backendEndpoint) {
//     throw new Error("Missing backend endpoint in the configuration.");
// }

// Define memory storage for Teams application
const storage = new MemoryStorage();
const app = new Application({
    storage
});

function cleanText(text) {
    let cleanedText = text;

    // Replace sequences of spaces with a single space
    cleanedText = cleanedText.replace(/\s+/g, ' ');

    // Attempt to fix common encoding issues
    cleanedText = cleanedText.replace(/â€™|â€˜|â€˜|â€™|â€œ|â€/g, "'"); // Replaces some common misinterpreted characters
    cleanedText = cleanedText.replace(/â€�/g, '"'); // Replaces incorrectly encoded double quotes
    cleanedText = cleanedText.replace(/â€”/g, '-'); // Replaces incorrectly encoded dashes

    return cleanedText;
}


// Send message to your Python Quart backend
async function sendMessageToBackend(context, message) {
    console.log(`Sending message to backend: ${message}`); // Logging the message being sent
    try {
        const response = await axios.post(config.backendEndpoint, {
            messages: [{ role: 'user', content: message }]
        }, {
            responseType: 'text'  // Ensures the response data is treated as a plain string
        });
        console.log(`Received raw response from backend: ${response.data}`); // Logging the raw response

        // Split the response by newlines and filter out any empty lines
        const responseParts = response.data.trim().split('\n').filter(part => part);
        let fullResponse = '';

        // Process each part of the response
        responseParts.forEach(part => {
            try {
                const jsonResponse = JSON.parse(part); // Parse each part as JSON
                if (jsonResponse.choices) {
                    jsonResponse.choices.forEach(choice => {
                        choice.messages.forEach(msg => {
                            if (msg.role === 'assistant') { // Concatenate messages from 'assistant'
                                fullResponse += msg.content + ' '; // Add a space for readability between messages
                            }
                        });
                    });
                }
            } catch (error) {
                console.error('Error parsing response part:', error);
                // Log the part that couldn't be parsed for debugging
                console.error('Part that caused the error:', part);
            }
        });

        // Log and send the concatenated message
        console.log("Full response:", fullResponse.trim());
        fullResponse = cleanText(fullResponse);
        await context.sendActivity(fullResponse.trim()); // Send the full, concatenated response to the user
    } catch (error) {
        console.error('Error sending message to backend:', error);
        // Log more details if available
        if (error.response) {
            console.log('Error response data:', error.response.data);
            console.log('Error response status:', error.response.status);
        }
        // Send a fallback error message to the user
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

