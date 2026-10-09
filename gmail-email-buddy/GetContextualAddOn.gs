/**
 * @OnlyCurrentDoc
 */

// Constants
var ADDON_TITLE = 'Email AI Assistant';
var OPENAI_API_KEY_PROPERTY = "OPENAI_API_KEY";
var OPENAI_CHAT_URL = "https://api.openai.com/v1/chat/completions";
var OPENAI_MODEL = "gpt-4o-mini"; // Using a more capable model like gpt-4o-mini might help with context
// var OPENAI_MODEL = "gpt-3.5-turbo";


// Constants for Chat Feature
var CHAT_HISTORY_KEY = "CURRENT_CHAT_HISTORY";
var CHAT_MESSAGE_ID_KEY = "CURRENT_CHAT_MESSAGE_ID"; // Stores the ID of the message that initiated the chat (for thread context)
var CHAT_THREAD_ID_KEY = "CURRENT_CHAT_THREAD_ID"; // Stores the thread ID for the chat
var CHAT_SUBJECT_KEY = "CURRENT_CHAT_SUBJECT";
var MAX_EMAIL_BODY_SNIPPET_IN_THREAD = 800; // Max characters from each email body to include in thread context for chat
var MAX_CONVERSATION_CONTEXT_FOR_CHAT_LENGTH = 7000; // Max total characters for the conversation context in chat

// Constants for Summaries
var MAX_MESSAGES_FOR_CONV_SUMMARY = 25; // Max messages to include in a conversation summary
var MAX_BODY_SNIPPET_FOR_CONV_SUMMARY = 1500; // Max characters from each email body for conversation summary context


/**
 * Returns the contextual add-on card.
 * @param {Object} event Event object.
 * @returns {Card[]}
 */
function getContextualAddOn(event) {
  var card = CardService.newCardBuilder().setHeader(CardService.newCardHeader().setTitle(ADDON_TITLE));
  var messageId = null;
  var threadId = null;
  var canProcessEmail = false;

  if (event && event.gmail) {
    if (event.gmail.messageId) {
      messageId = event.gmail.messageId;
      threadId = event.gmail.threadId; // Gmail event object provides threadId
      canProcessEmail = true;
      Logger.log("getContextualAddOn: Using messageId: " + messageId + ", threadId: " + threadId);
    } else {
      Logger.log("getContextualAddOn: event.gmail.messageId not found.");
    }
  } else {
    Logger.log("getContextualAddOn: event.gmail object not found.");
  }

  var userProps = PropertiesService.getUserProperties();
  var activeChatThreadId = userProps.getProperty(CHAT_THREAD_ID_KEY);

  // If currently in a chat for this thread, show the chat card
  if (threadId && activeChatThreadId === threadId) {
    var chatHistoryJson = userProps.getProperty(CHAT_HISTORY_KEY);
    var chatHistory = chatHistoryJson ? JSON.parse(chatHistoryJson) : [];
    var chatSubject = userProps.getProperty(CHAT_SUBJECT_KEY);
    // Pass messageId as well, as it's used in buildChatCardView for the subtitle
    var initiatingMessageIdForChat = userProps.getProperty(CHAT_MESSAGE_ID_KEY);
    return [buildChatCardView(chatHistory, chatSubject, initiatingMessageIdForChat, threadId).build()];
  }

  var emailActionsSection = CardService.newCardSection().setHeader("Email Tools");
  if (canProcessEmail && messageId && threadId) {
    emailActionsSection.addWidget(CardService.newTextParagraph()
      .setText("Use AI to analyze, respond, or chat about this email/conversation."));

    var emailActionButtons = CardService.newButtonSet()
      .addButton(CardService.newTextButton()
        .setText("Summarize This Email") // Clarified: "This Email"
        .setOnClickAction(CardService.newAction()
          .setFunctionName("generateSingleEmailSummary")
          .setParameters({ messageId: messageId })))
      .addButton(CardService.newTextButton()
        .setText("Summarize Conversation")
        .setOnClickAction(CardService.newAction()
          .setFunctionName("summarizeEntireConversation")
          .setParameters({ threadId: threadId }))) // Use threadId
      .addButton(CardService.newTextButton()
        .setText("Create Reply (to this email)")
        .setOnClickAction(CardService.newAction()
          .setFunctionName("createReply")
          .setParameters({ messageId: messageId })))
      .addButton(CardService.newTextButton()
        .setText("Chat about this Conversation") // Clarified: "this Conversation"
        .setOnClickAction(CardService.newAction()
          .setFunctionName("startChat")
          .setParameters({ messageId: messageId, threadId: threadId }))); // Pass both
    emailActionsSection.addWidget(emailActionButtons);
  } else {
    emailActionsSection.addWidget(CardService.newTextParagraph()
      .setText("Open an email to enable AI-powered actions and chat."));
  }
  card.addSection(emailActionsSection);

  var settingsSection = CardService.newCardSection();
  settingsSection.addWidget(CardService.newButtonSet()
    .addButton(CardService.newTextButton()
      .setText("Configure OpenAI API Key")
      .setOnClickAction(CardService.newAction()
        .setFunctionName("showSettings"))));
  card.addSection(settingsSection);

  return [card.build()];
}

/**
 * Initializes a new chat session about the current email conversation/thread.
 * @param {Object} event The event object containing messageId and threadId.
 * @return {ActionResponse} The ActionResponse to display the chat interface.
 */
function startChat(event) {
  var messageId = event.parameters.messageId; // ID of the currently viewed message
  var threadId = event.parameters.threadId;   // ID of the conversation thread

  if (!messageId || !threadId) {
    return createErrorCard("Error: Message ID or Thread ID missing. Cannot start chat.");
  }

  try {
    var thread = GmailApp.getThreadById(threadId);
    if (!thread) {
      return createErrorCard("Error: Could not retrieve email thread to start chat.");
    }

    var messagesInThread = thread.getMessages();
    var conversationContextText = "";
    var currentLength = 0;
    var subject = messagesInThread.length > 0 ? messagesInThread[0].getSubject() : "N/A"; // Get subject from first message

    // Iterate messages in chronological order to build context
    // However, if we need to truncate, we want to keep the most recent ones.
    // So, build the full text then truncate if needed, or build in reverse and then reverse string.
    // For simplicity now: build chronologically, then take the tail if too long.

    var tempConversationArray = [];
    for (var i = 0; i < messagesInThread.length; i++) {
      var msg = messagesInThread[i];
      var bodySnippet = msg.getPlainBody().substring(0, MAX_EMAIL_BODY_SNIPPET_IN_THREAD);
      if (msg.getPlainBody().length > MAX_EMAIL_BODY_SNIPPET_IN_THREAD) {
        bodySnippet += "...";
      }
      var messageText = "\n\n---\nMessage " + (i + 1) + " of " + messagesInThread.length +
                        "\nFrom: " + msg.getFrom() +
                        "\nDate: " + msg.getDate().toUTCString() +
                        "\nSubject: " + msg.getSubject() + // Individual message subject can vary slightly
                        "\n\nBody Snippet:\n" + bodySnippet;
      tempConversationArray.push(messageText);
    }
    
    conversationContextText = tempConversationArray.join("");

    if (conversationContextText.length > MAX_CONVERSATION_CONTEXT_FOR_CHAT_LENGTH) {
      conversationContextText = "...(Context Truncated due to length)...\n" +
                                conversationContextText.substring(conversationContextText.length - MAX_CONVERSATION_CONTEXT_FOR_CHAT_LENGTH);
      Logger.log("Chat context truncated. Original length: " + tempConversationArray.join("").length + ", Truncated length: " + conversationContextText.length);
    }


    var initialSystemPrompt = "You are an AI assistant. This chat session is dedicated to discussing the ENTIRE email conversation provided below. " +
                              "The following is a transcript of the email thread, with messages separated by '---'.\n\n" +
                              "--- BEGIN EMAIL THREAD TRANSCRIPT ---\n" +
                              conversationContextText + "\n" +
                              "--- END EMAIL THREAD TRANSCRIPT ---\n\n" +
                              "YOUR TASK: Answer the user's questions based EXCLUSIVELY on the email thread transcript provided above. " +
                              "When the user asks about 'the initial email', 'what a specific email said', 'the content of the email titled X', or any similar question about the content of any email in this thread, you MUST search through the entire provided transcript to find the relevant message and answer based on its content. " +
                              "If part of an email was truncated in the transcript (indicated by '...'), acknowledge that you only have that snippet. " +
                              "Do not claim information was not provided if it is within this transcript. Focus all responses on this specific email conversation.";

    var initialMessages = [
      { role: "system", content: initialSystemPrompt },
      { role: "assistant", content: "I've loaded the email conversation: '" + subject + "'. What would you like to discuss or ask about it?" }
    ];

    var userProps = PropertiesService.getUserProperties();
    userProps.setProperty(CHAT_HISTORY_KEY, JSON.stringify(initialMessages));
    userProps.setProperty(CHAT_MESSAGE_ID_KEY, messageId); // Store the message ID that initiated the chat
    userProps.setProperty(CHAT_THREAD_ID_KEY, threadId);   // Store the active chat thread ID
    userProps.setProperty(CHAT_SUBJECT_KEY, subject);

    var chatCard = buildChatCardView(initialMessages, subject, messageId, threadId);
    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().pushCard(chatCard.build()))
      .build();

  } catch (e) {
    Logger.log("Error in startChat: " + e.toString() + "\nStack: " + e.stack);
    return createErrorCard(refineGmailErrorMessage(e, "Error starting chat session"));
  }
}

/**
 * Handles a new message submitted by the user in the chat interface.
 * @param {Object} event The event object from the chat card.
 * @return {ActionResponse} The ActionResponse to update the chat interface.
 */
function handleUserChatMessage(event) {
  var userInput = event.formInputs && event.formInputs.chatInput ? event.formInputs.chatInput[0].trim() : "";
  var initiatingMessageId = event.parameters.messageId; // ID of the message that initiated chat
  var threadId = event.parameters.threadId;
  var subject = event.parameters.subject;


  if (!userInput) {
    var existingHistory = JSON.parse(PropertiesService.getUserProperties().getProperty(CHAT_HISTORY_KEY) || "[]");
    var chatCardWithError = buildChatCardView(existingHistory, subject, initiatingMessageId, threadId, "Please type a message.");
    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().updateCard(chatCardWithError.build()))
      .build();
  }
  if (!initiatingMessageId || !threadId || !subject) {
      return createErrorCard("Chat context lost (messageId, threadId, or subject missing). Please end this chat and start a new one.");
  }

  try {
    var userProps = PropertiesService.getUserProperties();
    var chatHistoryJson = userProps.getProperty(CHAT_HISTORY_KEY);
    var messages = chatHistoryJson ? JSON.parse(chatHistoryJson) : [];

    messages.push({ role: "user", content: userInput });

    var aiResponseText = callOpenAIChatAPI(messages);
    messages.push({ role: "assistant", content: aiResponseText });

    userProps.setProperty(CHAT_HISTORY_KEY, JSON.stringify(messages));

    var updatedChatCard = buildChatCardView(messages, subject, initiatingMessageId, threadId);
    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().updateCard(updatedChatCard.build()))
      .build();

  } catch (e) {
    Logger.log("Error in handleUserChatMessage: " + e.toString());
    var existingHistoryOnError = JSON.parse(PropertiesService.getUserProperties().getProperty(CHAT_HISTORY_KEY) || "[]");
    var errorChatCard = buildChatCardView(existingHistoryOnError, subject, initiatingMessageId, threadId, "Error: " + e.message);
     return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().updateCard(errorChatCard.build()))
      .build();
  }
}

/**
 * Builds the card interface for the chat.
 * @param {Array<Object>} chatHistoryArray The array of chat messages.
 * @param {string} currentSubject The subject of the email being discussed.
 * @param {string} currentMessageId The ID of the email that initiated the chat.
 * @param {string} currentThreadId The ID of the thread being discussed.
 * @param {string} [errorMessage] Optional error message to display.
 * @return {CardBuilder} The CardBuilder object for the chat interface.
 */
function buildChatCardView(chatHistoryArray, currentSubject, currentMessageId, currentThreadId, errorMessage) {
  var chatCard = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle("Chat: " + currentSubject)
      // Using messageId (that initiated chat) for subtitle, or threadId if preferred
      .setSubtitle("Discussing conversation (Thread ID: ..." + currentThreadId.slice(-6) + ")"));

  var chatSection = CardService.newCardSection();

  if (errorMessage) {
    chatSection.addWidget(CardService.newTextParagraph().setText("<b>Status:</b> <font color=\"red\">" + errorMessage + "</font>"));
  }

  chatHistoryArray.forEach(function(msg) {
    if (msg.role === "system" && msg.content.includes("--- BEGIN EMAIL THREAD TRANSCRIPT ---")) {
        // Optionally, don't display the huge system prompt with the full thread in the chat UI
        // Or display a condensed version / placeholder
        // For now, skipping it to keep UI cleaner if it's too long.
        // chatSection.addWidget(CardService.newTextParagraph().setText("<i>[System: Email thread context loaded.]</i>"));
        return; // Skip showing the full system prompt in UI
    }
    if (msg.role === "user") {
      chatSection.addWidget(CardService.newTextParagraph().setText("<b>You:</b> " + msg.content));
    } else if (msg.role === "assistant") {
      chatSection.addWidget(CardService.newTextParagraph().setText("<b>AI:</b> " + msg.content));
    }
  });

  chatSection.addWidget(CardService.newTextInput()
    .setFieldName("chatInput")
    .setTitle("Your Message")
    .setHint("Type your message to the AI..."));

  var buttonSet = CardService.newButtonSet()
    .addButton(CardService.newTextButton()
      .setText("Send")
      .setOnClickAction(CardService.newAction()
        .setFunctionName("handleUserChatMessage")
        .setParameters({ messageId: currentMessageId, threadId: currentThreadId, subject: currentSubject }))
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED))
    .addButton(CardService.newTextButton()
      .setText("End Chat & Go Back")
      .setOnClickAction(CardService.newAction()
        .setFunctionName("endChat")
        // Pass threadId to ensure main card rebuilds correctly for this thread
        .setParameters({ originalThreadId: currentThreadId, originalMessageIdIfAny: currentMessageId }))
      .setTextButtonStyle(CardService.TextButtonStyle.TEXT));
  chatSection.addWidget(buttonSet);

  chatCard.addSection(chatSection);
  return chatCard;
}

/**
 * Ends the current chat session and clears history.
 * @param {Object} event The event object.
 * @return {ActionResponse} ActionResponse to navigate back to the main card.
 */
function endChat(event) {
  var userProps = PropertiesService.getUserProperties();
  userProps.deleteProperty(CHAT_HISTORY_KEY);
  userProps.deleteProperty(CHAT_MESSAGE_ID_KEY);
  userProps.deleteProperty(CHAT_THREAD_ID_KEY);
  userProps.deleteProperty(CHAT_SUBJECT_KEY);

  Logger.log("Chat session ended.");
  var originalThreadId = event.parameters.originalThreadId;
  var originalMessageIdIfAny = event.parameters.originalMessageIdIfAny;

  var refreshEvent = {};
  // To rebuild the main card, we need messageId and threadId if available
  if (originalThreadId && originalMessageIdIfAny) {
    refreshEvent = { gmail: { messageId: originalMessageIdIfAny, threadId: originalThreadId } };
  } else if (originalThreadId) { // Fallback if only threadId was passed (shouldn't happen with current flow)
      // We need a messageId to properly trigger contextual addon.
      // This part might need refinement if we can end chat from a state where messageId isn't clear.
      // For now, assume originalMessageIdIfAny will be present from chat card params.
      Logger.log("EndChat: originalMessageIdIfAny not found, only threadId. Main card might not fully refresh context.");
      refreshEvent = { gmail: { threadId: originalThreadId } }; // May not be enough for getContextualAddon
  }


  var mainCard = getContextualAddOn(refreshEvent); // Rebuild main card for the original context
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(mainCard[0])) // Update current card to main
    .build();
}

/**
 * Calls the OpenAI Chat API with a full message history.
 * @param {Array<Object>} messagesArray The array of message objects (role, content).
 * @return {string} The AI's response text.
 */
function callOpenAIChatAPI(messagesArray) {
  var apiKey = PropertiesService.getScriptProperties().getProperty(OPENAI_API_KEY_PROPERTY);
  if (!apiKey) {
    Logger.log("OpenAI API key not found for chat.");
    throw new Error("OpenAI API key not configured. Please set it in the add-on settings.");
  }

  var payload = {
    model: OPENAI_MODEL,
    messages: messagesArray,
    temperature: 0.5, // Slightly lower for more factual recall based on provided text
    max_tokens: 1000  // Increased slightly if needed for longer AI responses
  };

  var options = {
    method: "post",
    contentType: "application/json",
    headers: { Authorization: "Bearer " + apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  Logger.log("OpenAI Request Payload (messages part only): " + JSON.stringify(messagesArray.map(function(m){return {role:m.role, content: (m.content || "").substring(0,100) + "..."};})));


  var response = UrlFetchApp.fetch(OPENAI_CHAT_URL, options);
  var responseCode = response.getResponseCode();
  var responseBody = response.getContentText();

  if (responseCode >= 400) {
    Logger.log("OpenAI API Chat Error - Status: " + responseCode + " Response: " + responseBody);
    var errorMsg = "OpenAI API Error (" + responseCode + ")";
    try {
      var errorData = JSON.parse(responseBody);
      if (errorData && errorData.error && errorData.error.message) {
        errorMsg += ": " + errorData.error.message;
      }
      if (errorData && errorData.error && errorData.error.type === 'insufficient_quota') {
        errorMsg = "OpenAI API Error: Insufficient quota. Please check your OpenAI billing and plan.";
      }
    } catch (e) { /* Ignore parsing error */ }
    throw new Error(errorMsg);
  }

  try {
    var responseData = JSON.parse(responseBody);
    if (responseData.choices && responseData.choices.length > 0 && responseData.choices[0].message && responseData.choices[0].message.content) {
      Logger.log("OpenAI Response (content snippet): " + responseData.choices[0].message.content.substring(0,200) + "...");
      return responseData.choices[0].message.content.trim();
    } else {
      Logger.log("OpenAI API Chat - Unexpected response structure: " + responseBody);
      throw new Error("AI response format error. No content found.");
    }
  } catch (e) {
    Logger.log("OpenAI API Chat - Failed to parse JSON response: " + responseBody + "\nError: " + e.toString());
    throw new Error("Invalid JSON response from AI.");
  }
}

/**
 * Handles Gmail API errors by providing more specific messages.
 * @param {Error} e The error object.
 * @param {string} defaultMessage The default message prefix.
 * @returns {string} A refined error message.
 */
function refineGmailErrorMessage(e, defaultMessage) {
    var errorMessage = e.message || String(e);
    if (errorMessage.includes("Metadata scope") || errorMessage.includes("format FULL")) {
        return defaultMessage + ": Access denied. The add-on may lack permission to read full email content or headers. Please ensure all permissions are granted. (Details: " + errorMessage + ")";
    } else if (errorMessage.includes("Missing access token") || errorMessage.includes("MailboxService.GetMessage") || errorMessage.includes("No item with that id")) {
        return defaultMessage + ": Authorization or access issue. The email/thread might be inaccessible, moved, or permissions are insufficient. Try reloading. (Details: " + errorMessage + ")";
    }
    return defaultMessage + ": " + errorMessage;
}

/**
 * Generates a summary of the single current email.
 * @param {Object} event The event object with messageId.
 * @returns {ActionResponse}
 */
function generateSingleEmailSummary(event) {
  var messageId = event.parameters.messageId;
  if (!messageId) return createErrorCard("Error: Message ID missing. Cannot generate summary for a single email.");

  try {
    var message = GmailApp.getMessageById(messageId);
    if (!message) return createErrorCard("Error: Could not retrieve email. It might be inaccessible or moved.");

    var body = message.getPlainBody();
    var subject = message.getSubject();
    var from = message.getFrom();
    var date = message.getDate().toUTCString();

    var prompt = "Concisely summarize this single email (2-4 sentences is ideal):\n\n" +
                 "From: " + from + "\n" +
                 "Date: " + date + "\n" +
                 "Subject: " + subject + "\n\n" +
                 "Body:\n" + body.substring(0, 4000); // Max length for summary prompt
    var summary = callOpenAI(prompt, "You are an expert email summarizer. Be brief and factual.");

    var summaryCard = CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle('Single Email Summary'))
      .addSection(CardService.newCardSection()
        .addWidget(CardService.newKeyValue().setTopLabel("From").setContent(from))
        .addWidget(CardService.newKeyValue().setTopLabel("Date").setContent(date))
        .addWidget(CardService.newKeyValue().setTopLabel("Subject").setContent(subject))
        .addWidget(CardService.newTextParagraph().setText("<b>Summary:</b>"))
        .addWidget(CardService.newTextParagraph().setText(summary))
        .addWidget(CardService.newButtonSet()
          .addButton(CardService.newTextButton()
            .setText("Create Reply From This")
            .setOnClickAction(CardService.newAction()
              .setFunctionName("createReply")
              .setParameters({ messageId: messageId })))))
      .build();
    return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().pushCard(summaryCard)).build();
  } catch (e) {
    Logger.log("Error in generateSingleEmailSummary: " + e.toString() + "\nStack: " + e.stack);
    return createErrorCard(refineGmailErrorMessage(e, "Error generating single email summary"));
  }
}

/**
 * Summarizes the entire email conversation thread.
 * @param {Object} event The event object with threadId.
 * @returns {ActionResponse}
 */
function summarizeEntireConversation(event) {
  var threadId = event.parameters.threadId;
  if (!threadId) return createErrorCard("Error: Thread ID missing. Cannot summarize conversation.");

  try {
    var thread = GmailApp.getThreadById(threadId);
    if (!thread) return createErrorCard("Error: Could not retrieve thread for conversation summary.");

    var messagesInThread = thread.getMessages();
    var conversationSubject = messagesInThread.length > 0 ? messagesInThread[0].getSubject() : "N/A";

    var messagesToProcess = messagesInThread.slice(-MAX_MESSAGES_FOR_CONV_SUMMARY); // Take last N messages
    var conversationText = "";

    messagesToProcess.forEach(function(msg, index) {
      var bodySnippet = msg.getPlainBody().substring(0, MAX_BODY_SNIPPET_FOR_CONV_SUMMARY);
      if (msg.getPlainBody().length > MAX_BODY_SNIPPET_FOR_CONV_SUMMARY) bodySnippet += "...";
      conversationText += "Message " + (messagesInThread.length - messagesToProcess.length + index + 1) + " of " + messagesInThread.length + ":\n";
      conversationText += "From: " + msg.getFrom() + "\n";
      conversationText += "Date: " + msg.getDate().toUTCString() + "\n";
      conversationText += "Subject: " + msg.getSubject() + "\n";
      conversationText += "Body Snippet:\n" + bodySnippet + "\n\n---\n\n";
    });

    if (messagesInThread.length > MAX_MESSAGES_FOR_CONV_SUMMARY) {
        conversationText = "[Summarizing the last " + MAX_MESSAGES_FOR_CONV_SUMMARY + " messages of " + messagesInThread.length + " total in this conversation]\n\n" + conversationText;
    }

    var prompt = "Provide a concise summary of the key points, main topics, questions asked, decisions made, and any action items from the following email conversation. Organize the summary logically.\n\nConversation Subject: " + conversationSubject + "\n\n" + conversationText;
    var summary = callOpenAI(prompt, "You are an expert conversation summarizer. Focus on clarity and actionable insights.");

    var summaryCard = CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle('Conversation Summary'))
      .addSection(CardService.newCardSection()
        .addWidget(CardService.newKeyValue().setTopLabel("Overall Subject").setContent(conversationSubject))
        .addWidget(CardService.newTextParagraph().setText("<b>Summary of Conversation:</b>"))
        .addWidget(CardService.newTextParagraph().setText(summary)))
      .build();
    return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().pushCard(summaryCard)).build();
  } catch (e) {
    Logger.log("Error in summarizeEntireConversation: " + e.toString() + "\nStack: " + e.stack);
    return createErrorCard(refineGmailErrorMessage(e, "Error summarizing conversation"));
  }
}


/**
 * Creates a reply generation card based on the single current email.
 * @param {Object} event The event object.
 * @returns {ActionResponse}
 */
function createReply(event) {
  var messageId = event.parameters.messageId;
  if (!messageId) return createErrorCard("Error: Message ID missing. Cannot create reply.");

  try {
    var message = GmailApp.getMessageById(messageId);
    if (!message) return createErrorCard("Error: Could not retrieve email for reply.");

    var body = message.getPlainBody();
    var subject = message.getSubject();
    var from = message.getFrom();

    var prompt = "You are writing an email reply. Based on the original email below, generate a professional and helpful reply body. " +
                 "Address key points from the original email. Only provide the body of the reply, without your own salutation (like 'Hello John,') or closing (like 'Best regards, AI'). The user will add those.\n\n" +
                 "Original Email From: " + from + "\n" +
                 "Original Email Subject: " + subject + "\n\n" +
                 "Original Email Body:\n" + body.substring(0, 4000); // Max length for reply context
    var generatedReply = callOpenAI(prompt, "You are an expert email drafter. Your replies are polite, concise, and address all key points of the original email.");

    var replyCard = CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle('Generated Reply Suggestion'))
      .addSection(CardService.newCardSection()
        .addWidget(CardService.newTextParagraph().setText("<b>Suggested Reply Body (edit as needed):</b>"))
        .addWidget(CardService.newTextInput()
          .setFieldName("replyContent")
          .setTitle("Edit Reply Body")
          .setValue(generatedReply)
          .setMultiline(true))
        .addWidget(CardService.newButtonSet()
          .addButton(CardService.newTextButton()
            .setText("Create Draft with This")
            .setOnClickAction(CardService.newAction()
              .setFunctionName("createDraftWithGeneratedReply")
              .setParameters({ messageId: messageId })))
          .addButton(CardService.newTextButton()
            .setText("Send This Reply")
            .setOnClickAction(CardService.newAction()
              .setFunctionName("sendGeneratedReply")
              .setParameters({ messageId: messageId })))))
      .build();
    return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().pushCard(replyCard)).build();
  } catch (e) {
    Logger.log("Error in createReply: " + e.toString() + "\nStack: " + e.stack);
    return createErrorCard(refineGmailErrorMessage(e, "Error creating reply"));
  }
}

/**
 * Sends an email reply directly using the generated text.
 * @param {Object} event The event object.
 * @returns {ActionResponse}
 */
function sendGeneratedReply(event) {
  var messageId = event.parameters.messageId;
  var replyText = event.formInputs && event.formInputs.replyContent ? event.formInputs.replyContent[0] : '';

  if (!messageId) return createErrorCard("Error: Message ID missing for sending reply.");
  if (!replyText || replyText.trim() === "") return createErrorCard("Error: Reply text is empty.");

  try {
    var message = GmailApp.getMessageById(messageId);
    if (!message) return createErrorCard("Error: Original message not found to send reply.");
    message.reply(replyText.trim()); // Consider .replyAll() or options for more control if needed
    return createSuccessNotificationResponse("Reply sent successfully!", event);
  } catch (e) {
    Logger.log("Error in sendGeneratedReply: " + e.toString());
    return createErrorCard(refineGmailErrorMessage(e, "Error sending reply"));
  }
}

/**
 * Creates a draft reply using the generated text.
 * @param {Object} event The event object.
 * @returns {ActionResponse}
 */
function createDraftWithGeneratedReply(event) {
  var messageId = event.parameters.messageId;
  var replyText = event.formInputs && event.formInputs.replyContent ? event.formInputs.replyContent[0] : '';

  if (!messageId) return createErrorCard("Error: Message ID missing for creating draft.");
  if (!replyText || replyText.trim() === "") return createErrorCard("Error: Reply text is empty for draft.");

  try {
    var message = GmailApp.getMessageById(messageId);
    if (!message) return createErrorCard("Error: Original message not found to create draft reply.");
    message.createDraftReply(replyText.trim());
    return createSuccessNotificationResponse("Draft created successfully! Find it in your Drafts folder.", event);
  } catch (e) {
    Logger.log("Error in createDraftWithGeneratedReply: " + e.toString());
    return createErrorCard(refineGmailErrorMessage(e, "Error creating draft"));
  }
}

/**
 * Wrapper for single-turn OpenAI calls, uses callOpenAIChatAPI.
 * @param {string} prompt The user's prompt/query.
 * @param {string} [systemContent] Optional system message content.
 * @returns {string} The generated text.
 */
function callOpenAI(prompt, systemContent) {
  // API key check is done by callOpenAIChatAPI
  var effectiveSystemContent = systemContent || "You are a helpful AI assistant. Provide concise and relevant responses.";

  var messagesToAPI = [
    { role: "system", content: effectiveSystemContent },
    { role: "user", content: prompt }
  ];
  return callOpenAIChatAPI(messagesToAPI);
}


/**
 * Shows the settings card for API key configuration.
 * @param {Object} event Optional event object.
 * @returns {ActionResponse}
 */
function showSettings(event) {
  var currentApiKey = PropertiesService.getScriptProperties().getProperty(OPENAI_API_KEY_PROPERTY) || "";
  var maskedKey = currentApiKey ? "••••••••••••" + currentApiKey.slice(-4) : "Not Set";

  var card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader().setTitle('Settings'));
  var section = CardService.newCardSection().setHeader("OpenAI API Configuration");
  section.addWidget(CardService.newTextParagraph().setText("Configure your OpenAI API key. Get your API key from platform.openai.com. This add-on uses models like '" + OPENAI_MODEL +"'."));
  section.addWidget(CardService.newKeyValue().setTopLabel("Current API Key Status").setContent(maskedKey).setIcon(CardService.Icon.TICKET));
  section.addWidget(CardService.newTextInput()
    .setFieldName("apiKey")
    .setTitle("Enter or Update OpenAI API Key")
    .setHint("Paste your API key here (e.g., sk-...)"));
  section.addWidget(CardService.newButtonSet()
    .addButton(CardService.newTextButton()
      .setText("Save API Key")
      .setOnClickAction(CardService.newAction().setFunctionName("saveApiKey"))
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)));
  card.addSection(section);
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().pushCard(card.build()))
    .build();
}

/**
 * Saves the OpenAI API key.
 * @param {Object} event The event object.
 * @returns {ActionResponse}
 */
function saveApiKey(event) {
  var apiKeyInput = (event.formInputs && event.formInputs.apiKey && event.formInputs.apiKey[0])
                   ? event.formInputs.apiKey[0].trim() : "";

  if (!apiKeyInput || !apiKeyInput.startsWith("sk-") || apiKeyInput.length < 20) {
    Logger.log("saveApiKey: Invalid API key format.");
    var errorCard = CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle("Settings - Error"))
      .addSection(CardService.newCardSection()
        .addWidget(CardService.newTextParagraph().setText("<b>Error:</b> Please enter a valid OpenAI API key. It should start with 'sk-' and be an appropriate length."))
        .addWidget(CardService.newTextInput().setFieldName("apiKey").setTitle("New OpenAI API Key").setHint("Enter your OpenAI API key (sk-...)").setValue(apiKeyInput)) // Show what they entered
        .addWidget(CardService.newButtonSet().addButton(CardService.newTextButton()
          .setText("Save API Key")
          .setOnClickAction(CardService.newAction().setFunctionName("saveApiKey"))
          .setTextButtonStyle(CardService.TextButtonStyle.FILLED))))
      .build();
    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().updateCard(errorCard)) // Update existing settings card with error
      .build();
  }
  try {
    PropertiesService.getScriptProperties().setProperty(OPENAI_API_KEY_PROPERTY, apiKeyInput);
    Logger.log("OpenAI API Key saved successfully.");
    
    // Try to return to the main contextual card
    var homeCardEvent = {};
    if (event && event.gmail && event.gmail.messageId && event.gmail.threadId) {
        homeCardEvent.gmail = { messageId: event.gmail.messageId, threadId: event.gmail.threadId };
    }
    var homeCard = getContextualAddOn(homeCardEvent); // Rebuild based on current context
    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().updateCard(homeCard[0])) // Update current card
      .setNotification(CardService.newNotification().setText("API Key saved successfully."))
      .build();
  } catch (e) {
    Logger.log("Error saving API Key: " + e.toString());
    return createErrorCard("Error saving API key: " + (e.message || String(e)));
  }
}

/**
 * Creates an error card to display issues to the user.
 * @param {string} errorMessage The error message.
 * @returns {ActionResponse}
 */
function createErrorCard(errorMessage) {
  Logger.log("Creating error card: " + errorMessage);
  var card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader()
    .setTitle('Error Occurred')
    .setImageUrl('https://ssl.gstatic.com/images/icons/material/system/1x/error_red_24dp.png'));
  var section = CardService.newCardSection();
  section.addWidget(CardService.newTextParagraph().setText(errorMessage));
  // No dismiss button here, rely on universal navigation or back button
  card.addSection(section);
  
  // Universal navigation behavior: if an error card is pushed, user can use back arrow
  // If we want to replace current view: CardService.newNavigation().updateCard(card.build())
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().pushCard(card.build()))
    .build();
}

/**
 * Creates a success notification and optionally navigates.
 * Typically pops the current card or updates to main.
 * @param {string} message The success message.
 * @param {Object} [event] Optional event object to help rebuild the main card after success.
 * @returns {ActionResponse}
 */
function createSuccessNotificationResponse(message, event) {
  Logger.log("Creating success notification: " + message);
  var navigation = CardService.newNavigation();

  var mainCardEvent = {};
  // Try to get context from the event to rebuild the main card correctly
  if (event && event.parameters && event.parameters.messageId) {
    // This structure comes from actions like createDraftWithGeneratedReply
    // We need threadId as well for the main card.
    // This part is tricky as not all events will have full original context.
    try {
        var tempMsg = GmailApp.getMessageById(event.parameters.messageId);
        if (tempMsg) {
            mainCardEvent.gmail = { messageId: event.parameters.messageId, threadId: tempMsg.getThread().getId() };
        }
    } catch (e) {
        Logger.log("Could not get threadId from messageId in createSuccessNotificationResponse: " + e);
    }
  } else if (event && event.gmail && event.gmail.messageId && event.gmail.threadId) {
    // This structure comes from the add-on entry point (getContextualAddOn)
    mainCardEvent.gmail = { messageId: event.gmail.messageId, threadId: event.gmail.threadId };
  }

  if (mainCardEvent.gmail && typeof getContextualAddOn === 'function') {
    var mainAddOnCard = getContextualAddOn(mainCardEvent);
    if (mainAddOnCard && mainAddOnCard.length > 0) {
        navigation.updateCard(mainAddOnCard[0]); // Go back to the main contextual card
    } else {
        navigation.popCard(); // Fallback
    }
  } else {
    navigation.popCard(); // Default: pop current card (e.g., reply suggestion card)
  }

  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification().setText(message))
    .setNavigation(navigation)
    .build();
}
