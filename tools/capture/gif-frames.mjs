// Frame-by-frame capture for the README demo GIF. Usage: node gif-frames.mjs <framesdir>
// Frames are PNG screenshots of a fixed clip; ffmpeg turns them into the GIF.
import { chromium } from "playwright";
import fs from "fs";
const OUT = process.argv[2] || "./.frames";
fs.rmSync(OUT, { recursive: true, force: true });
fs.mkdirSync(OUT, { recursive: true });
const BASE = "http://127.0.0.1:4310/";
// Each host gets a viewport sized so the crop starts on a column edge (no words cut).
const CW = 840, CH = 760;
let CLIP = { x: 0, y: 0, width: CW, height: CH };

const browser = await chromium.launch({ headless: true, executablePath: process.env.CHROME_BIN || undefined });
const ctx = await browser.newContext({ viewport: { width: CW, height: CH }, deviceScaleFactor: 1 });
const page = await ctx.newPage();
let n = 0;
const frame = async () => { await page.screenshot({ path: `${OUT}/f${String(n++).padStart(4, "0")}.png`, clip: CLIP }); };
const hold = async (count) => { await frame(); for (let i = 1; i < count; i++) { fs.copyFileSync(`${OUT}/f${String(n - 1).padStart(4, "0")}.png`, `${OUT}/f${String(n++).padStart(4, "0")}.png`); } };
const film = async (ms) => { const end = Date.now() + ms; while (Date.now() < end) { await frame(); await page.waitForTimeout(60); } };
const cursorTo = async (x, y) => {
  await page.evaluate(([x, y]) => {
    let c = document.getElementById("__cursor");
    if (!c) {
      c = document.createElement("div");
      c.id = "__cursor";
      c.style.cssText = "position:fixed;z-index:9999;width:22px;height:22px;border-radius:50%;background:rgba(15,108,189,.28);border:2px solid #0f6cbd;pointer-events:none;transition:left .35s cubic-bezier(.2,.7,.2,1),top .35s cubic-bezier(.2,.7,.2,1);left:700px;top:600px";
      document.body.appendChild(c);
    }
    c.style.left = x - 11 + "px"; c.style.top = y - 11 + "px";
  }, [x, y]);
  await film(420);
};
const centerOf = async (loc, frameOffset = { x: 0, y: 0 }) => { const b = await loc.boundingBox(); return [b.x + b.width / 2 + frameOffset.x, b.y + b.height / 2 + frameOffset.y]; };

// --- Outlook AI assistant ---
await page.goto(BASE + "?host=outlook-ai&clean&narrow", { waitUntil: "networkidle" });
await page.waitForFunction(() => window.__ready);
const f = page.frameLocator("#addin");
await f.locator("#generateSummaryBtn").waitFor();
await hold(14);
let [x, y] = await centerOf(f.locator("#generateSummaryBtn"));
await cursorTo(x, y);
await f.locator("#generateSummaryBtn").click();
await film(400);
await f.locator("#summaryView.active").waitFor({ timeout: 15000 });
await film(200);
await hold(26);
[x, y] = await centerOf(f.locator("#createReplyFromSummaryBtn"));
await cursorTo(x, y);
await f.locator("#createReplyFromSummaryBtn").click();
await f.locator("#replyView.active").waitFor({ timeout: 15000 });
await film(200);
await hold(24);
[x, y] = await centerOf(f.locator("#sendReplyBtn"));
await cursorTo(x, y);
await f.locator("#sendReplyBtn").click();
await film(700);
await hold(26);

// --- Gmail Email Buddy ---
await page.goto(BASE + "?host=gmail&clean&narrow", { waitUntil: "networkidle" });
await page.waitForFunction(() => window.__ready);
await hold(12);
const btn = (t) => page.locator("#card button", { hasText: t }).first();
const done = () => page.waitForFunction(() => !document.getElementById("progress").classList.contains("on"), null, { timeout: 15000 });
[x, y] = await centerOf(btn("Summarize Conversation"));
await cursorTo(x, y);
await btn("Summarize Conversation").click();
await film(300);
await done();
await film(120);
await hold(28);
await page.locator("#card [data-back]").click();
[x, y] = await centerOf(btn("Chat about this Conversation"));
await cursorTo(x, y);
await btn("Chat about this Conversation").click();
await done();
await film(120);
const input = page.locator('#card input[name="chatInput"]');
[x, y] = await centerOf(input);
await cursorTo(x, y);
const q = "What is still blocking the print?";
for (let i = 1; i <= q.length; i += 3) { await input.fill(q.slice(0, i)); await frame(); }
await input.fill(q);
await frame();
[x, y] = await centerOf(btn("Send"));
await cursorTo(x, y);
await btn("Send").click();
await film(300);
await done();
await film(120);
await hold(34);

console.log("frames:", n);
await browser.close();
