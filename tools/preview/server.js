#!/usr/bin/env node
// Local preview server: runs the real add-in UIs inside mock mail hosts.
//
//   1. build the add-ins:   (cd outlook-ai-assistant && npm run build)
//                           (cd outlook-summarizer && npx next build)
//   2. start Next:          (cd outlook-summarizer && npx next start -p 3102)
//   3. start this server:   node tools/preview/server.js
//   4. open http://127.0.0.1:4310/
//
// What is real: the add-in task panes (built from this repo) and Email Buddy's
// Apps Script code, loaded unchanged. What is mocked: the mail host frames, the
// mailbox (invented emails in demo-data.js), Office.js (office-stub.js), Apps
// Script services (apps-script-shim.js) and the model (fake-openai.js).
// Nothing leaves your machine.
"use strict";
const http = require("http");
const fs = require("fs");
const path = require("path");
const fakeOpenAI = require("./fake-openai");

const ROOT = path.resolve(__dirname, "..", "..");
const PORT = Number(process.env.PREVIEW_PORT || 4310);
const NEXT_PORT = Number(process.env.NEXT_PORT || 3102);
const TYPES = {
  ".html": "text/html; charset=utf-8", ".js": "text/javascript; charset=utf-8", ".gs": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8", ".png": "image/png", ".svg": "image/svg+xml", ".json": "application/json",
  ".woff2": "font/woff2", ".map": "application/json",
};
const OFFICE_TAG = /<script[^>]*appsforoffice\.microsoft\.com[^>]*><\/script>/g;
const CSP_TAG = /<meta[^>]*Content-Security-Policy[^>]*>/gi;
const STUB_TAGS = '<script src="/preview/demo-data.js"></script><script src="/preview/office-stub.js"></script>';

function swapOffice(html) {
  return html.replace(CSP_TAG, "").replace(OFFICE_TAG, STUB_TAGS);
}

function sendFile(res, file, transform) {
  fs.readFile(file, (err, buf) => {
    if (err) { res.writeHead(404); return res.end("not found"); }
    const type = TYPES[path.extname(file)] || "application/octet-stream";
    let body = buf;
    if (transform && type.startsWith("text/html")) body = Buffer.from(transform(buf.toString("utf8")));
    res.writeHead(200, { "Content-Type": type, "Cache-Control": "no-store" });
    res.end(body);
  });
}

function safeJoin(base, rel) {
  const p = path.normalize(path.join(base, rel));
  return p.startsWith(base) ? p : null;
}

function proxyToNext(req, res) {
  const headers = Object.assign({}, req.headers);
  delete headers["accept-encoding"]; // keep HTML uncompressed so the Office.js tag can be swapped
  const up = http.request({ host: "127.0.0.1", port: NEXT_PORT, path: req.url, method: req.method, headers }, (r) => {
    const type = r.headers["content-type"] || "";
    if (!type.includes("text/html")) { res.writeHead(r.statusCode, r.headers); return r.pipe(res); }
    const chunks = [];
    r.on("data", (c) => chunks.push(c));
    r.on("end", () => {
      const html = swapOffice(Buffer.concat(chunks).toString("utf8"));
      const out = Object.assign({}, r.headers);
      delete out["content-length"];
      delete out["content-encoding"];
      res.writeHead(r.statusCode, out);
      res.end(html);
    });
  });
  up.on("error", () => { res.writeHead(502); res.end("Start the summarizer first: cd outlook-summarizer && npx next start -p " + NEXT_PORT); });
  req.pipe(up);
}

const server = http.createServer((req, res) => {
  const url = new URL(req.url, "http://localhost");
  const p = decodeURIComponent(url.pathname);

  if (p === "/" || p === "/hosts" || p === "/hosts/") return sendFile(res, path.join(__dirname, "hosts", "index.html"));
  if (p.startsWith("/preview/openai/")) {
    let raw = "";
    req.on("data", (c) => (raw += c));
    req.on("end", () => fakeOpenAI.handle(req, res, raw));
    return;
  }
  if (p.startsWith("/preview/")) {
    const f = safeJoin(__dirname, p.slice("/preview/".length));
    return f ? sendFile(res, f) : (res.writeHead(400), res.end());
  }
  if (p.startsWith("/hosts/")) {
    const f = safeJoin(path.join(__dirname, "hosts"), p.slice("/hosts/".length));
    return f ? sendFile(res, f) : (res.writeHead(400), res.end());
  }
  if (p.startsWith("/outlook-ai/")) {
    const f = safeJoin(path.join(ROOT, "outlook-ai-assistant", "dist"), p.slice("/outlook-ai/".length));
    return f ? sendFile(res, f, swapOffice) : (res.writeHead(400), res.end());
  }
  if (p.startsWith("/gmail/")) {
    const f = safeJoin(path.join(ROOT, "gmail-email-buddy"), p.slice("/gmail/".length));
    return f ? sendFile(res, f) : (res.writeHead(400), res.end());
  }
  if (p.startsWith("/summarizer")) {
    req.url = req.url.replace(/^\/summarizer/, "") || "/";
    return proxyToNext(req, res);
  }
  if (p.startsWith("/_next/")) return proxyToNext(req, res);
  res.writeHead(404);
  res.end("not found");
});

server.listen(PORT, "127.0.0.1", () => {
  console.log(`Preview: http://127.0.0.1:${PORT}/`);
});
