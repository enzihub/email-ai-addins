/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// Ensure Office is ready
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    // DOMContentLoaded might be too early if Office.js hasn't fully initialized its internal state.
    // It's generally safer to run initialization logic directly if onReady is the entry point.
    run();
  }
});

// Name of the Outlook roaming setting that stores the user's own OpenAI key.
// The key itself is never in the code; users paste it in the Settings view.
const ROAMING_SETTING_NAME = "openai_api_key_v1";
let currentEmailItemId: string | null = null; // To store the current item ID for context

// Define a type for mailProps for clarity, matching Office.ReplyFormData structure for htmlBody
interface MailProperties {
  htmlBody: string;
  // Other properties like 'to', 'cc', 'subject' could be added if needed
  // but for displayReplyAllForm, only htmlBody is being used from this custom object.
  // For full functionality, use Office.ReplyFormData.
}

// UI Elements Cache
let ui = {
  loadingIndicator: null as HTMLElement | null,
  messageBar: null as HTMLElement | null,
  mainView: null as HTMLElement | null,
  summaryView: null as HTMLElement | null,
  replyView: null as HTMLElement | null,
  settingsView: null as HTMLElement | null,
  emailNotSelectedMessage: null as HTMLElement | null,
  emailToolsActions: null as HTMLElement | null,
  generateSummaryBtn: null as HTMLButtonElement | null,
  createReplyBtn: null as HTMLButtonElement | null,
  summaryFrom: null as HTMLElement | null,
  summarySubject: null as HTMLElement | null,
  summaryText: null as HTMLElement | null,
  createReplyFromSummaryBtn: null as HTMLButtonElement | null,
  backToMainSummaryBtn: null as HTMLButtonElement | null,
  replyContentTextarea: null as HTMLTextAreaElement | null,
  createDraftBtn: null as HTMLButtonElement | null,
  sendReplyBtn: null as HTMLButtonElement | null,
  backToMainReplyBtn: null as HTMLButtonElement | null,
  apiKeyStatus: null as HTMLElement | null,
  apiKeyInput: null as HTMLInputElement | null,
  saveApiKeyBtn: null as HTMLButtonElement | null,
  backToMainSettingsBtn: null as HTMLButtonElement | null,
  configureApiKeyBtn: null as HTMLButtonElement | null,
};

function cacheDOMElements() {
  ui.loadingIndicator = document.getElementById("loadingIndicator");
  ui.messageBar = document.getElementById("messageBar");
  ui.mainView = document.getElementById("mainView");
  ui.summaryView = document.getElementById("summaryView");
  ui.replyView = document.getElementById("replyView");
  ui.settingsView = document.getElementById("settingsView");
  ui.emailNotSelectedMessage = document.getElementById("emailNotSelectedMessage");
  ui.emailToolsActions = document.getElementById("emailToolsActions");
  ui.generateSummaryBtn = document.getElementById("generateSummaryBtn") as HTMLButtonElement;
  ui.createReplyBtn = document.getElementById("createReplyBtn") as HTMLButtonElement;
  ui.summaryFrom = document.getElementById("summaryFrom");
  ui.summarySubject = document.getElementById("summarySubject");
  ui.summaryText = document.getElementById("summaryText");
  ui.createReplyFromSummaryBtn = document.getElementById(
    "createReplyFromSummaryBtn"
  ) as HTMLButtonElement;
  ui.backToMainSummaryBtn = document.getElementById("backToMainSummaryBtn") as HTMLButtonElement;
  ui.replyContentTextarea = document.getElementById("replyContentTextarea") as HTMLTextAreaElement;
  ui.createDraftBtn = document.getElementById("createDraftBtn") as HTMLButtonElement;
  ui.sendReplyBtn = document.getElementById("sendReplyBtn") as HTMLButtonElement;
  ui.backToMainReplyBtn = document.getElementById("backToMainReplyBtn") as HTMLButtonElement;
  ui.apiKeyStatus = document.getElementById("apiKeyStatus");
  ui.apiKeyInput = document.getElementById("apiKeyInput") as HTMLInputElement;
  ui.saveApiKeyBtn = document.getElementById("saveApiKeyBtn") as HTMLButtonElement;
  ui.backToMainSettingsBtn = document.getElementById("backToMainSettingsBtn") as HTMLButtonElement;
  ui.configureApiKeyBtn = document.getElementById("configureApiKeyBtn") as HTMLButtonElement;
}

function run() {
  cacheDOMElements();
  showView("mainView");
  checkEmailSelectionAndUpdateUI();
  loadApiKeyStatus();
  registerEventHandlers();

  // Listen for item changes to update UI contextually
  Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, (eventArgs) => {
    checkEmailSelectionAndUpdateUI();
  });
}

function registerEventHandlers() {
  ui.generateSummaryBtn?.addEventListener("click", handleGenerateSummary);
  ui.createReplyBtn?.addEventListener("click", handleCreateReply);
  ui.configureApiKeyBtn?.addEventListener("click", showSettingsView);
  ui.saveApiKeyBtn?.addEventListener("click", handleSaveApiKey);

  ui.createReplyFromSummaryBtn?.addEventListener("click", handleCreateReply); // Re-uses createReply logic
  ui.backToMainSummaryBtn?.addEventListener("click", () => showView("mainView"));

  ui.createDraftBtn?.addEventListener("click", handleCreateDraftWithGeneratedReply);
  ui.sendReplyBtn?.addEventListener("click", handleSendGeneratedReply);
  ui.backToMainReplyBtn?.addEventListener("click", () => showView("mainView"));
  ui.backToMainSettingsBtn?.addEventListener("click", () => {
    showView("mainView");
    loadApiKeyStatus(); // Refresh API key status when returning to main view
  });
}

function showView(viewId: string) {
  document
    .querySelectorAll(".view")
    .forEach((view) => (view as HTMLElement).classList.remove("active"));
  ui.loadingIndicator?.classList.remove("active");
  hideMessageBar(); // Clear messages when switching views

  const activeView = document.getElementById(viewId);
  if (activeView) {
    activeView.classList.add("active");
  }
}

function checkEmailSelectionAndUpdateUI() {
  const item = Office.context.mailbox.item;
  currentEmailItemId = item?.itemId; // Store current item ID

  if (item && item.itemType === Office.MailboxEnums.ItemType.Message) {
    ui.emailToolsActions.style.display = "block";
    ui.emailNotSelectedMessage.style.display = "none";
    ui.generateSummaryBtn.disabled = false;
    ui.createReplyBtn.disabled = false;
  } else {
    ui.emailToolsActions.style.display = "none";
    ui.emailNotSelectedMessage.style.display = "block";
    ui.generateSummaryBtn.disabled = true;
    ui.createReplyBtn.disabled = true;
    showView("mainView"); // If no email selected, revert to main view
  }
}

interface EmailDetails {
  subject: string;
  body: string;
  from: string; // email address
  itemId: string;
  toRecipients?: Office.EmailAddressDetails[];
  ccRecipients?: Office.EmailAddressDetails[];
}

async function getEmailDetails(): Promise<EmailDetails | null> {
  const item = Office.context.mailbox.item;
  if (
    !item ||
    item.itemType !== Office.MailboxEnums.ItemType.Message ||
    item.itemId !== currentEmailItemId
  ) {
    showMessageBar("Please select an email message, or the selected email has changed.", "error");
    checkEmailSelectionAndUpdateUI(); // Re-validate
    return null;
  }

  const message = item as Office.MessageRead;
  const subject = message.subject;
  const from = message.from.emailAddress;
  const itemId = message.itemId;
  const toRecipients = message.to;
  const ccRecipients = message.cc;

  return new Promise((resolve, reject) => {
    message.body.getAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve({ subject, body: result.value.trim(), from, itemId, toRecipients, ccRecipients });
      } else {
        console.error("Error getting email body:", result.error);
        showMessageBar("Error getting email body: " + result.error.message, "error");
        reject(result.error);
      }
    });
  });
}

async function handleGenerateSummary() {
  showLoading("Analyzing email for key points..."); // Updated loading message
  const emailDetails = await getEmailDetails();
  if (!emailDetails) {
    hideLoading();
    return;
  }

  // --- Prompt 1: Extract Key Points ---
  const prompt1 = `Analyze the following email and extract the key points, main topic, any decisions made, and any explicit action items or questions asked. Present these as a concise list or bullet points.

Subject: ${emailDetails.subject}
From: ${emailDetails.from}

Body:
${emailDetails.body}

Key Information:`;

  let keyPoints = "";
  try {
    keyPoints = await callOpenAI(prompt1);
    console.log("OpenAI Response (Key Points):", keyPoints); // For debugging
    showLoading("Generating descriptive summary..."); // Update loading message
  } catch (error) {
    console.error("Error extracting key points:", error);
    showMessageBar(`Failed to extract key points: ${error.message || error}`, "error");
    hideLoading();
    return; // Stop if the first step fails
  }

  // --- Prompt 2: Generate Descriptive Summary from Key Points ---
  const prompt2 = `Based on the following key information extracted from an email, write a descriptive summary paragraph (around 3-5 sentences) that captures the main essence, decisions, and required actions or questions. Ensure the summary is coherent and easy to understand for someone who hasn't read the full email. Avoid starting with phrases like "The email discusses..." or "This email is about...". Directly summarize the content.

Key Information:
---
${keyPoints}
---

Descriptive Summary:`;

  try {
    const descriptiveSummary = await callOpenAI(prompt2);
    console.log("OpenAI Response (Descriptive Summary):", descriptiveSummary); // For debugging

    // Update UI with the final summary
    ui.summaryFrom.textContent = emailDetails.from;
    ui.summarySubject.textContent = emailDetails.subject;
    ui.summaryText.textContent = descriptiveSummary; // Use the result from the second prompt
    showView("summaryView");
  } catch (error) {
    console.error("Error generating descriptive summary:", error);
    showMessageBar(`Failed to generate descriptive summary: ${error.message || error}`, "error");
  } finally {
    hideLoading();
  }
}

async function handleCreateReply() {
  showLoading("Preparing reply suggestion...");
  const emailDetails = await getEmailDetails();
  if (!emailDetails) {
    hideLoading();
    return;
  }

  const prompt = `You are an AI assistant helping to draft an email reply.
Original Email Subject: ${emailDetails.subject}
Original Email From: ${emailDetails.from}
Original Email Body:
---
${emailDetails.body}
---
Based on the original email, write a professional and helpful reply.
The reply should be concise, address key points, and maintain a polite tone.
ONLY provide the body of the reply, without any greetings (like "Hi [Name],") or signatures (like "Best regards,").
The user will add those manually.`;

  try {
    const generatedReply = await callOpenAI(prompt);
    ui.replyContentTextarea.value = generatedReply;
    showView("replyView");
  } catch (error) {
    console.error("Error creating reply:", error);
    showMessageBar(`Failed to generate reply: ${error.message || error}`, "error");
  } finally {
    hideLoading();
  }
}

async function handleCreateDraftWithGeneratedReply() {
  showLoading("Creating draft...");
  const replyText = ui.replyContentTextarea.value;
  const originalEmail = await getEmailDetails(); // Get original email details for context

  if (!originalEmail) {
    showMessageBar("Could not get original email details to create draft.", "error");
    hideLoading();
    return;
  }

  if (!replyText || replyText.trim() === "") {
    showMessageBar("Reply text is empty.", "error");
    hideLoading();
    return;
  }

  // Prepare reply properties
  // The `Office.ReplyFormDATA` type used in original code was custom.
  // Office.js `displayReplyAllForm` takes a string (HTML body) or an `Office.ReplyFormData` object.
  // We will pass the HTML body string directly.
  const htmlBodyContent = replyText.trim().replace(/\n/g, "<br />");

  // Fix: Use the appropriate API format for displayReplyAllForm
  Office.context.mailbox.item.displayReplyAllForm({ htmlBody: htmlBodyContent });

  showMessageBar(
    "Draft reply form opened with generated content. Please review and save or send.",
    "info"
  );
  hideLoading();
  showView("mainView"); // Revert to main view
}

async function handleSendGeneratedReply() {
  showLoading("Opening reply form...");
  const replyText = ui.replyContentTextarea.value;

  if (
    !Office.context.mailbox.item ||
    Office.context.mailbox.item.itemType !== Office.MailboxEnums.ItemType.Message
  ) {
    showMessageBar("Cannot send reply. Please ensure an email is selected.", "error");
    hideLoading();
    return;
  }

  if (!replyText || replyText.trim() === "") {
    showMessageBar("Reply text is empty.", "error");
    hideLoading();
    return;
  }

  // Outlook add-ins typically don't send email directly in the background for security/UX reasons.
  // They populate a compose form (reply, reply all, or forward).
  // The `Office.ReplyFormDATA` type used in original code was custom.
  // We will pass the HTML body string directly.
  const htmlBodyContent = replyText.trim().replace(/\n/g, "<br />");

  // Fix: Use the appropriate API format for displayReplyAllForm
  (Office.context.mailbox.item as Office.MessageRead).displayReplyAllForm({
    htmlBody: htmlBodyContent,
  });

  showMessageBar("Reply form opened with generated content. Please review and send.", "info");
  hideLoading();
  showView("mainView");
}

async function callOpenAI(prompt: string): Promise<string> {
  const apiKey = Office.context.roamingSettings.get(ROAMING_SETTING_NAME) as string | undefined;
  if (!apiKey) {
    throw new Error("OpenAI API key not configured. Please go to Settings.");
  }

  const url = "https://api.openai.com/v1/chat/completions";
  const payload = {
    model: "gpt-3.5-turbo", // Or "gpt-4" if you have access
    messages: [
      {
        role: "system",
        content:
          "You are a helpful AI assistant. For email replies, strictly provide only the body of the reply, without greetings or signatures.",
      },
      { role: "user", content: prompt },
    ],
    temperature: 0.6, // Slightly lower for more deterministic professional replies
    max_tokens: 600, // Increased slightly
    top_p: 1.0,
    frequency_penalty: 0.0,
    presence_penalty: 0.0,
  };

  const options = {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: "Bearer " + apiKey,
    },
    body: JSON.stringify(payload),
  };

  try {
    const response = await fetch(url, options);
    const responseBodyText = await response.text(); // Get text for robust error parsing

    if (!response.ok) {
      console.error(
        "OpenAI API Error - Status:",
        response.status,
        "Response Body:",
        responseBodyText
      );
      let errorMsg = `OpenAI API Error (${response.status})`;
      try {
        const errorData = JSON.parse(responseBodyText); // Try to parse error
        if (errorData && errorData.error && errorData.error.message) {
          errorMsg += `: ${errorData.error.message}`;
        } else if (errorData && errorData.message) {
          // Sometimes error is directly in message
          errorMsg += `: ${errorData.message}`;
        } else {
          errorMsg += `. Response: ${responseBodyText.substring(0, 100)}...`; // Include part of response if parsing fails
        }
      } catch (e) {
        errorMsg += `. Non-JSON response: ${responseBodyText.substring(0, 100)}...`;
      }
      throw new Error(errorMsg);
    }

    const responseData = JSON.parse(responseBodyText);
    if (
      responseData.choices &&
      responseData.choices.length > 0 &&
      responseData.choices[0].message &&
      responseData.choices[0].message.content
    ) {
      return responseData.choices[0].message.content.trim();
    } else {
      console.error("OpenAI API - Unexpected response structure:", responseBodyText);
      throw new Error("OpenAI API Error: Unexpected response data structure.");
    }
  } catch (e) {
    console.error("OpenAI API Call Failed:", e);
    // Check for network errors specifically
    if (
      e instanceof TypeError &&
      (e.message.toLowerCase().includes("failed to fetch") ||
        e.message.toLowerCase().includes("networkerror"))
    ) {
      throw new Error(
        "Network error: Could not connect to OpenAI API. Please check your internet connection and firewall settings."
      );
    }
    throw e; // Re-throw other errors (could be the custom error from above or other exceptions)
  }
}

function showSettingsView() {
  loadApiKeyStatus();
  ui.apiKeyInput.value = ""; // Clear old input value
  showView("settingsView");
}

function loadApiKeyStatus() {
  const apiKey = Office.context.roamingSettings.get(ROAMING_SETTING_NAME) as string | undefined;
  if (apiKey && typeof apiKey === "string" && apiKey.length > 4) {
    ui.apiKeyStatus.textContent = `********${apiKey.slice(-4)}`;
  } else {
    ui.apiKeyStatus.textContent = "Not Set";
  }
}

function handleSaveApiKey() {
  const apiKeyInputValue = ui.apiKeyInput.value.trim();
  if (!apiKeyInputValue || !apiKeyInputValue.startsWith("sk-") || apiKeyInputValue.length < 30) {
    // Basic validation
    showMessageBar(
      "Invalid API key format. It should start with 'sk-' and be a valid length.",
      "error"
    );
    return;
  }

  Office.context.roamingSettings.set(ROAMING_SETTING_NAME, apiKeyInputValue);
  Office.context.roamingSettings.saveAsync((asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      showMessageBar("API Key saved successfully.", "info");
      loadApiKeyStatus(); // Update the status display
      setTimeout(() => showView("mainView"), 1500); // Go back after a short delay
    } else {
      console.error("Error saving API key:", asyncResult.error);
      showMessageBar("Error saving API key: " + asyncResult.error.message, "error");
    }
  });
}

// UI Utility Functions
function showLoading(message: string = "Processing...") {
  ui.loadingIndicator.textContent = message;
  ui.loadingIndicator.classList.add("active");
  // Disable all action buttons in the current view
  document
    .querySelectorAll(".view.active button")
    .forEach((button) => ((button as HTMLButtonElement).disabled = true));
  ui.configureApiKeyBtn.disabled = true; // Also disable global settings button
}

function hideLoading() {
  ui.loadingIndicator.classList.remove("active");
  // Re-enable all action buttons
  document
    .querySelectorAll("button")
    .forEach((button) => ((button as HTMLButtonElement).disabled = false));
  checkEmailSelectionAndUpdateUI(); // Re-apply disabled state for email actions if no email is selected
}

function showMessageBar(message: string, type: "info" | "error") {
  ui.messageBar.textContent = message;
  ui.messageBar.className = "message-bar"; // Reset classes
  if (type === "error") {
    ui.messageBar.classList.add("error-message");
  } else {
    ui.messageBar.classList.add("info-message");
  }
  ui.messageBar.style.display = "block";

  // Auto-hide after a few seconds
  setTimeout(
    () => {
      hideMessageBar();
    },
    type === "error" ? 7000 : 5000
  );
}

function hideMessageBar() {
  ui.messageBar.style.display = "none";
  ui.messageBar.textContent = "";
}
