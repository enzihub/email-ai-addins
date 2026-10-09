# Email Summarizer (Outlook, Next.js)

The first prototype of the Outlook add-in: a Next.js 13 task pane that reads the open email with
Office.js and shows its first few sentences. It makes no AI or network call. See
`createSimpleSummary()` in `pages/index.tsx` if you want to plug a model in.

## Run it in Outlook

Outlook only loads add-ins over HTTPS, so you need a trusted localhost certificate first.
No certificate or key is stored in this repo; Microsoft's `office-addin-dev-certs` makes one on your machine.

```bash
npm install
npm run certs        # office-addin-dev-certs install (creates + trusts ~/.office-addin-dev-certs)
npm run dev:https    # server.js: Next.js behind HTTPS on https://localhost:3000
```

Then in Outlook: **Add-ins → My add-ins → Custom add-ins → Add from file** and pick
`public/manifest.xml`. Open an email and click **Summarize Email** in the ribbon.

To remove the dev certificate later: `npx office-addin-dev-certs uninstall`.

## Other scripts

| Script | What it does |
| --- | --- |
| `npm run dev` | Plain `next dev` on http://localhost:3000 (fine for UI work in a browser, not for Outlook) |
| `npm run build` / `npm start` | Production build and server |
| `PORT=3100 npm run dev:https` | HTTPS server on another port (update `public/manifest.xml` to match) |

## Files

```
pages/index.tsx            the task pane UI and the summary logic
pages/_document.tsx        loads office.js from Microsoft's CDN
public/manifest.xml        Outlook add-in manifest (points at https://localhost:3000)
public/email-summarizer.html  an earlier static chat-style UI mock, kept for reference
public/assets/icon-*.png   placeholder ribbon icons (same as outlook-ai-assistant; the originals were corrupt)
server.js                  HTTPS dev server using office-addin-dev-certs
```
