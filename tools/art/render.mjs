// Render the art boards in art.html to PNG at 2x with transparent rounded corners.
// Usage: node tools/art/render.mjs   (writes assets/*.png and docs/og.png)
import { chromium } from "playwright";
import path from "path";
import { fileURLToPath } from "url";
const here = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(here, "..", "..");
const browser = await chromium.launch({ headless: true, executablePath: process.env.CHROME_BIN || undefined, args: ["--allow-file-access-from-files"] });
const page = await browser.newPage({ viewport: { width: 1300, height: 900 }, deviceScaleFactor: 2 });
const external = [];
page.on("request", (r) => { if (!r.url().startsWith("file:")) external.push(r.url()); });
await page.goto("file://" + path.join(here, "art.html"), { waitUntil: "load" });
await page.evaluate(() => document.fonts.ready);
await page.waitForTimeout(300);
const jobs = [
  ["#hero", "assets/hero.png", 2],
  ["#collage", "assets/screenshots.png", 2],
  ["#how", "assets/how-it-works.png", 2],
];
for (const [sel, out] of jobs) {
  await page.locator(sel).screenshot({ path: path.join(root, out), omitBackground: true });
  console.log("wrote", out);
}
const og = await browser.newPage({ viewport: { width: 1300, height: 900 }, deviceScaleFactor: 1 });
await og.goto("file://" + path.join(here, "art.html"), { waitUntil: "load" });
await og.evaluate(() => document.fonts.ready);
await og.waitForTimeout(300);
await og.locator("#og").scrollIntoViewIfNeeded();
const bb = await og.evaluate(() => { const r = document.querySelector("#og").getBoundingClientRect(); return { x: r.left + scrollX, y: r.top + scrollY }; });
await og.screenshot({ fullPage: true, path: path.join(root, "docs/og.png"), clip: { x: Math.round(bb.x), y: Math.round(bb.y), width: 1200, height: 630 } });
console.log("wrote docs/og.png");
console.log(external.length ? "EXTERNAL: " + external.join(", ") : "no external requests");
await browser.close();
