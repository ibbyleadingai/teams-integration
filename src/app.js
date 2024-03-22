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
    const response = await axios.post(`${config.backendEndpoint}/conversation_teams`, requestBody, {
      headers: {
        'Content-Type': 'application/json',
        // Include any other headers your backend requires
      }
    });

    // Send the backend's response back to the user in Teams
    const backendMessage = response.data.messages.map(msg => msg.content).join('\n');
    await context.sendActivity(backendMessage);

  } catch (error) {
    // Handle errors, e.g. if the backend is not reachable
    console.error('Error calling the backend:', error);
    await context.sendActivity('Sorry, I am having trouble reaching the backend. Please try again later.');
  }
});

app.ai.action(AI.HttpErrorActionName, async (context, state, data) => {
  await context.sendActivity("An AI request failed. Please try again later.");
  return AI.StopCommandName;
});

module.exports = app;
