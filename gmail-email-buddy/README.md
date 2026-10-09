# Email Buddy (Gmail add-on, Apps Script)

A Google Workspace add-on for Gmail. Open any email and the side panel offers:

- **Summarize This Email**: 2-4 sentences about the open message.
- **Summarize Conversation**: the topic, decisions, questions and action items from up to the last 25 messages in the thread.
- **Create Reply (to this email)**: an editable reply body, then **Create Draft with This** or **Send This Reply**.
- **Chat about this Conversation**: the thread is loaded as context (800 characters per message, 7,000 in total) and you can ask follow-up questions. **End Chat & Go Back** deletes the chat history.
- **Configure OpenAI API Key**: saves your key in Apps Script script properties.

It calls OpenAI `gpt-4o-mini` (the `OPENAI_MODEL` constant) through `UrlFetchApp`.

## Install for testing

1. Create a project at [script.google.com](https://script.google.com/).
2. Paste `GetContextualAddOn.gs` into the editor.
3. In **Project Settings**, tick **Show "appsscript.json" manifest file** and replace it with `appsscript.json` from this folder.
4. **Deploy → Test deployments → Application(s): Gmail → Install**, then accept the permissions.
5. Refresh Gmail, open an email, open the add-on and click **Configure OpenAI API Key**.

With [clasp](https://github.com/google/clasp): `clasp create --type standalone --title "Email Buddy"`, copy both files into the folder, then `clasp push`.

## Notes

- Script properties are shared by everyone who uses the same deployment, so one key serves all users. For a team, move the key to `PropertiesService.getUserProperties()` or to a proxy you control.
- The manifest asks for `https://mail.google.com/` because the add-on can send replies. Remove that scope (and `sendGeneratedReply`) if you only want drafts.
- Chat history is kept in user properties only while a chat is open.
