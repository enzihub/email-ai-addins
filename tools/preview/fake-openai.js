// Local stand-in for the OpenAI Chat Completions endpoint, used only by the
// preview. It returns fixed, hand-written answers about the invented demo
// emails so screenshots are reproducible and no API key or network is needed.
"use strict";

function lastUser(messages) {
  for (let i = messages.length - 1; i >= 0; i--) if (messages[i].role === "user") return messages[i].content || "";
  return "";
}

function answer(body) {
  const messages = body.messages || [];
  const user = lastUser(messages);
  const system = (messages[0] && messages[0].content) || "";
  const isChat = system.includes("BEGIN EMAIL THREAD TRANSCRIPT");

  // ---- Outlook AI assistant (venue email) ----
  if (user.startsWith("Analyze the following email and extract the key points")) {
    return [
      "- Topic: choosing a venue for the Q3 offsite on 6–7 November (38 people).",
      "- Harbourline Hall: $4,200, no breakout rooms.",
      "- The Granary: $5,900, two breakout rooms, on-site catering.",
      "- Pine Ridge Lodge: $8,300, includes rooms, 90 minutes out of town.",
      "- Deadline: pick a venue by Thursday; holds expire Friday at noon.",
      "- Questions: projector in every room? Dietary list from Jonah?",
    ].join("\n");
  }
  if (user.startsWith("Based on the following key information extracted from an email")) {
    return (
      "Marta has narrowed the Q3 offsite (6–7 November, 38 people) to three venues: Harbourline Hall at $4,200 with no breakout rooms, " +
      "The Granary at $5,900 with two breakout rooms and catering, and Pine Ridge Lodge at $8,300 with overnight rooms but a 90-minute drive. " +
      "All three are holding the dates until Friday at noon, so she needs a decision by Thursday. " +
      "She also wants to know whether every room needs a projector and is waiting on the dietary list from Jonah."
    );
  }
  if (user.startsWith("You are an AI assistant helping to draft an email reply.")) {
    return (
      "Thanks for pulling this together so quickly. Let's go with The Granary: the two breakout rooms matter more to us than the overnight stay, and on-site catering keeps things simple.\n\n" +
      "We only need a projector in the main room. The breakout rooms can use screens. I've asked Jonah to send you the dietary list today.\n\n" +
      "No need for a call unless something changes. Please confirm the booking once The Granary sends the contract."
    );
  }

  // ---- Gmail Email Buddy (print-proofs thread) ----
  if (user.startsWith("Concisely summarize this single email")) {
    return (
      "Priya confirms a 3,000-copy run is fine: the total rises from $6,850 to $7,940, printing stays on Monday 10 March, " +
      "and delivery moves to Saturday 15 March, so someone must sign for it at the warehouse. She will lock the files once she has the photo credit and your written approval."
    );
  }
  if (user.startsWith("Provide a concise summary of the key points")) {
    return [
      "**Topic:** final proofs and print run for the 48-page spring catalogue.",
      "",
      "**Decisions:**",
      "• Page 14 price corrected to $179 (Tarn rain shell).",
      "• Run increased from 2,500 to 3,000 copies; total now $7,940.",
      "",
      "**Dates:** print Monday 10 March, delivery Saturday 15 March.",
      "",
      "**Open actions:**",
      "• Jonah: send the photo credit for page 22.",
      "• Sam: send written approval so Priya can lock the files.",
      "• Arrange someone to sign for delivery on Saturday.",
    ].join("\n");
  }
  if (user.startsWith("You are writing an email reply.")) {
    return (
      "Thanks Priya, 3,000 copies at $7,940 works for us. Please go ahead with printing on Monday 10 March.\n\n" +
      "Jonah has the photo credit for page 22 and will send it this morning. Please treat this email as our written approval once that is in.\n\n" +
      "For Saturday 15 March, Leo from our warehouse team will be there to sign for the delivery."
    );
  }
  if (isChat) {
    const q = user.toLowerCase();
    if (/(need|todo|to do|action|waiting|block)/.test(q)) {
      return "Two things are blocking the print: the photo credit for page 22 (Jonah is getting it) and your written approval. Once Priya has both, she locks the files.";
    }
    if (/(cost|price|total|\$|pay)/.test(q)) {
      return "The 3,000-copy run comes to $7,940, up from $6,850 for 2,500 copies (Priya, 4 March). It prints on Monday 10 March and now arrives on Saturday 15 March, one day later than planned.";
    }
    if (/(when|date|deliver|arrive|print)/.test(q)) {
      return "Printing is on Monday 10 March. With the larger run, delivery to the warehouse moves from Friday 14 March to Saturday 15 March, and someone needs to be there to sign for it.";
    }
    return "Based on the thread, the catalogue is ready to print once the page 22 photo credit and your approval reach Priya. Ask me about cost, dates or open actions.";
  }
  return "This is the local preview model. It only knows the demo emails.";
}

function handle(req, res, raw) {
  let body = {};
  try { body = JSON.parse(raw || "{}"); } catch (e) {}
  const content = answer(body);
  const payload = {
    id: "chatcmpl-preview",
    object: "chat.completion",
    created: Math.floor(Date.now() / 1000),
    model: body.model || "preview",
    choices: [{ index: 0, message: { role: "assistant", content }, finish_reason: "stop" }],
  };
  const delay = Number(process.env.PREVIEW_MODEL_DELAY_MS || 700);
  setTimeout(() => {
    res.writeHead(200, { "Content-Type": "application/json", "Access-Control-Allow-Origin": "*" });
    res.end(JSON.stringify(payload));
  }, delay);
}

module.exports = { handle, answer };
