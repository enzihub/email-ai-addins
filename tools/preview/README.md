# Local preview

Runs the three add-ins in a **mock mail host** with invented emails, so you can see and screenshot
them without a Google or Microsoft account or an OpenAI key. Nothing leaves 127.0.0.1.

| File | Role |
| --- | --- |
| `server.js` | Serves the host page, the built Outlook task panes (with the Office.js tag swapped for the stub), proxies the Next.js summarizer, and answers `/preview/openai/*` |
| `hosts/index.html` | The mock mail window. `?host=gmail`, `?host=outlook-ai`, `?host=summarizer`. Add `&clean` to hide the preview badge, `&narrow` for a compact layout |
| `office-stub.js` | Minimal stand-in for the parts of Office.js the add-ins call. Not Microsoft's library |
| `apps-script-shim.js` | Minimal CardService / GmailApp / PropertiesService / UrlFetchApp so `GetContextualAddOn.gs` runs unchanged in a browser, plus a card renderer that approximates Gmail's |
| `demo-data.js` | The invented mailbox (all `.example` addresses) |
| `fake-openai.js` | Fixed answers about the demo emails, in the OpenAI chat-completions shape |

See the main README for the start commands. `tools/capture/` holds the headless Playwright scripts
that made the README screenshots and GIF (`CHROME_BIN` can point at any Chromium build).
