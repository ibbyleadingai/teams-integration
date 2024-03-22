const { MemoryStorage } = require("botbuilder");
const config = require("./config");
const axios = require('axios');

// See https://aka.ms/teams-ai-library to learn more about the Teams AI library.
const { Application, AI, preview } = require("@microsoft/teams-ai");

// See README.md to prepare your own OpenAI Assistant
if (!config.openAIKey || !config.openAIAssistantId) {
  throw new Error(
    "Missing OPENAI_API_KEY or OPENAI_ASSISTANT_ID. See README.md to prepare your own OpenAI Assistant."
  );
}

// Create AI components
// Use OpenAI
const planner = new preview.AssistantsPlanner({
  apiKey: config.openAIKey,
  assistant_id: config.openAIAssistantId,
});

// Define storage and application
const storage = new MemoryStorage();
const app = new Application({
  storage,
  ai: {
    planner,
  },
});

app.message("/reset", async (context, state) => {
  state.deleteConversationState();
  await context.sendActivity("Ok lets start this over.");
});

app.message(async (context, state) => {
  // Extract the message text from the incoming request
  const userMessage = context.activity.text;

  // Prepare the request body for the /conversation_teams endpoint
  const requestBody = {
    messages: [{
      role: 'user',
      content: userMessage
    }]
  };

  // Send the request to your backend
  try {
    const response = await axios.post(`https://matassist.azurewebsites.net/conversation_teams`, requestBody, {
        headers: {
            'Content-Type': 'application/json',
            // Include any other headers your backend requires
        }
    });

    // Check if the response has the expected structure and messages
    if (response.data && Array.isArray(response.data.choices)) {
        const messages = response.data.choices[0].messages;
        if (Array.isArray(messages)) {
            // Filter out any messages where the role is not 'assistant'
            const filteredMessages = messages.filter(msg => msg.role === 'assistant');

            // Join the content of the remaining messages and send back to the user
            const backendMessage = filteredMessages.map(msg => msg.content).join('\n');
            await context.sendActivity(backendMessage);
        } else {
            await context.sendActivity('Received unexpected message format from backend.');
        }
    } else {
        console.error('Unexpected response structure:', response.data);
        await context.sendActivity('Received unexpected response format from backend.');
    }

} catch (error) {
    console.error('Error calling the backend:', error.message);
    if (error.response) {
        // The request was made and the server responded with a status code
        // that falls out of the range of 2xx
        console.error('Response data:', error.response.data);
        console.error('Response status:', error.response.status);
        console.error('Response headers:', error.response.headers);
        await context.sendActivity(`Backend error: ${error.response.status} - ${JSON.stringify(error.response.data)}`);
    } else if (error.request) {
        // The request was made but no response was received
        console.error('No response received:', error.request);
        await context.sendActivity('No response received from the backend, please check the backend service.');
    } else {
        // Something happened in setting up the request that triggered an Error
        console.error('Error setting up the request:', error.message);
        await context.sendActivity('There was an error setting up the request to the backend.');
    }
}
});

app.ai.action(AI.HttpErrorActionName, async (context, state, data) => {
  await context.sendActivity("An AI request failed. Please try again later.");
  return AI.StopCommandName;
});

module.exports = app;
