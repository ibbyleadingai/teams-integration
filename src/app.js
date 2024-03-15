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
    console.log(`Sending message to backend: ${message}`); // Log the message being sent to the backend for debugging.
    try {
        // Send the user's message to your backend and set the response type to stream for handling streamed responses.
        const response = await axios.post(config.backendEndpoint, {
            messages: [{ role: 'user', content: message }]
        }, {
            responseType: 'stream'  // This tells Axios to handle the response as a stream.
        });

        let runningText = ''; // This will store text chunks that may not represent complete JSON strings.
        let fullResponse = ''; // This will accumulate the full response text.

        // Process the response stream.
        const streamProcessed = new Promise((resolve, reject) => {
            response.data.on('data', (chunk) => {
                runningText += chunk.toString(); // Add the new chunk to any previous incomplete text.
                const parts = runningText.split("\n"); // Assuming the backend sends JSON objects separated by newlines.
                runningText = parts.pop(); // The last item might be incomplete, so we'll wait for the next chunk to complete it.

                parts.forEach(part => {
                    if (part) { // Ignore empty parts which can occur with consecutive newlines.
                        try {
                            const jsonResponse = JSON.parse(part); // Try parsing each complete part as JSON.
                            jsonResponse.choices?.forEach(choice => {
                                choice.messages.forEach(msg => {
                                    if (msg.role === 'assistant') { // Only concatenate messages from the 'assistant'.
                                        fullResponse += msg.content + ' '; // Add a space between concatenated messages for readability.
                                    }
                                });
                            });
                        } catch (error) {
                            console.error('Error parsing part of the response:', part, error); // Log parsing errors for debugging.
                        }
                    }
                });
            });

            response.data.on('end', () => {
                resolve(fullResponse.trim()); // Resolve the promise with the full, concatenated response when the stream ends.
            });

            response.data.on('error', (error) => {
                reject(error); // Reject the promise if there's an error processing the stream.
            });
        });

        // Wait for the stream to be fully processed.
        fullResponse = await streamProcessed;
        console.log("Final full response:", fullResponse); // Log the full response for debugging.
        await context.sendActivity(fullResponse); // Send the full response back to the user in Teams.
    } catch (error) {
        console.error('Error sending message to backend:', error); // Log any errors that occur during the process.
        // Inform the user in Teams that there was an error processing their message.
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

