// Headless capture of the preview hosts. Usage: node capture.mjs <outdir> [step]
import { chromium } from "playwright";
import fs from "fs";
const OUT = process.argv[2] || "./shots";
const ONLY = process.argv[3];
fs.mkdirSync(OUT, { recursive: true });
const BASE = "http://127.0.0.1:4310/";
const W = 1440, H = 900;

const browser = await chromium.launch({ headless: true, executablePath: process.env.CHROME_BIN || undefined });
const ctx = await browser.newContext({ viewport: { width: W, height: H }, deviceScaleFactor: 2 });
const page = await ctx.newPage();
const errors = [];
page.on("pageerror", (e) => errors.push("pageerror: " + e.message));
page.on("console", (m) => { if (m.type() === "error") errors.push("console: " + m.text()); });
page.on("request", (r) => { const u = r.url(); if (!u.startsWith("http://127.0.0.1") && !u.startsWith("data:")) errors.push("EXTERNAL REQUEST " + u); });

const panel = async (name, sel) => { await page.waitForTimeout(250); await page.locator(sel).screenshot({ path: `${OUT}/panel-${name}.png` }); console.log("panel", name); };
const shot = async (name, opts = {}) => { await page.waitForTimeout(250); await page.screenshot({ path: `${OUT}/${name}.png`, ...opts }); console.log("shot", name); };
const want = (s) => !ONLY || ONLY === s;

// ---------- Gmail / Email Buddy ----------
if (want("gmail")) {
  await page.goto(BASE + "?host=gmail&clean", { waitUntil: "networkidle" });
  await page.waitForFunction(() => window.__ready);
  await shot("gmail-1-home");
  const click = async (text) => {
    await page.locator("#card button", { hasText: text }).first().click();
    await page.waitForFunction(() => !document.getElementById("progress").classList.contains("on"), null, { timeout: 15000 });
    await page.waitForTimeout(300);
  };
  await click("Summarize Conversation");
  await shot("gmail-2-conversation-summary");
  await panel("gmail-summary", "aside.side");
  await page.locator("#card [data-back]").click();
  await click("Create Reply (to this email)");
  await shot("gmail-3-reply");
  await panel("gmail-reply", "aside.side");
  await page.locator("#card [data-back]").click();
  await click("Chat about this Conversation");
  await page.fill('#card input[name="chatInput"]', "What does the bigger print run cost, and when does it arrive?");
  await click("Send");
  await page.fill('#card input[name="chatInput"]', "What is still blocking the print?");
  await click("Send");
  await shot("gmail-4-chat");
  await panel("gmail-chat", "aside.side");
}

// ---------- Outlook AI assistant ----------
if (want("outlook-ai")) {
  await page.goto(BASE + "?host=outlook-ai&clean", { waitUntil: "networkidle" });
  await page.waitForFunction(() => window.__ready);
  const f = page.frameLocator("#addin");
  await f.locator("#generateSummaryBtn").waitFor();
  await shot("outlook-ai-1-home");
  await f.locator("#generateSummaryBtn").click();
  await f.locator("#summaryView.active").waitFor({ timeout: 15000 });
  await shot("outlook-ai-2-summary");
  await panel("outlook-summary", "aside.pane");
  await f.locator("#createReplyFromSummaryBtn").click();
  await f.locator("#replyView.active").waitFor({ timeout: 15000 });
  await shot("outlook-ai-3-reply");
  await panel("outlook-reply", "aside.pane");
  await f.locator("#sendReplyBtn").click();
  await page.waitForTimeout(700);
  await shot("outlook-ai-4-compose");
}

// ---------- Outlook summarizer (Next.js) ----------
if (want("summarizer")) {
  await page.goto(BASE + "?host=summarizer&clean", { waitUntil: "networkidle" });
  await page.waitForFunction(() => window.__ready);
  const f = page.frameLocator("#addin");
  await f.locator("button", { hasText: "Summarize Email" }).waitFor({ timeout: 15000 });
  await shot("summarizer-1-home");
  await f.locator("button", { hasText: "Summarize Email" }).click();
  await page.waitForTimeout(600);
  await shot("summarizer-2-summary");
  await panel("summarizer", "aside.pane");
}

console.log(errors.length ? errors.join("\n") : "no page errors, no external requests");
await browser.close();
