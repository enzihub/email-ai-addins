# AI Email Assistant (Outlook, Office.js)

An Outlook task pane add-in in TypeScript, built from Microsoft's
[Office Add-in TaskPane template](https://github.com/OfficeDev/Office-Addin-TaskPane).

- **Generate Summary**: two OpenAI calls. The first pulls out key points, decisions, actions and
  questions. The second writes a 3-5 sentence summary from them.
- **Create Reply**: drafts the reply body only (no greeting or sign-off).
- **Open Reply Form with This** / **Create Draft with This Reply**: opens Outlook's reply-all form
  with the draft inside. Nothing is sent until you press Send.
- **Settings**: paste your OpenAI key. It is saved in Outlook roaming settings
  (`openai_api_key_v1`) for your mailbox, never in code.

## Run it

```bash
npm ci
npx office-addin-dev-certs install   # trusted localhost certificate, made on your machine
npm run dev-server                   # https://localhost:3000
```

Sideload `manifest.xml` (**Add-ins → My add-ins → Custom add-ins → Add from file**), or run
`npm start` to let `office-addin-debugging` sideload it into Outlook desktop for you.

| Script | What it does |
| --- | --- |
| `npm run build` | Production build into `dist/` (replaces `https://localhost:3000/` with `urlProd` from `webpack.config.js`) |
| `npm run validate` | Validate `manifest.xml` |
| `npm start` / `npm stop` | Sideload / remove the add-in with Microsoft's debugging tool |

The model is `gpt-3.5-turbo`, set in `callOpenAI()` in `src/taskpane/taskpane.ts`. Swap in a
current model before relying on it.
