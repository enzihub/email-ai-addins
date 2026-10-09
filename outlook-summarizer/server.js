// HTTPS dev server for the Outlook summarizer add-in.
//
// Outlook only loads add-ins over HTTPS. This wraps the Next.js app in an
// HTTPS server that uses the trusted localhost certificate created by
// Microsoft's office-addin-dev-certs tool. No certificate or key is stored
// in this repository: run `npm run certs` once to create and trust one on
// your own machine (it lives in ~/.office-addin-dev-certs).

const https = require("https");
const { parse } = require("url");
const next = require("next");
const devCerts = require("office-addin-dev-certs");

const port = Number(process.env.PORT || 3000);
const dev = process.env.NODE_ENV !== "production";
const app = next({ dev });
const handle = app.getRequestHandler();

async function main() {
  const options = await devCerts.getHttpsServerOptions();
  await app.prepare();
  https
    .createServer(options, (req, res) => handle(req, res, parse(req.url, true)))
    .listen(port, () => {
      console.log(`Outlook summarizer running at https://localhost:${port}`);
      console.log(`Sideload public/manifest.xml in Outlook to open it as a task pane.`);
    });
}

main().catch((err) => {
  console.error(err);
  console.error("\nNo dev certificate yet? Run `npm run certs` first.");
  process.exit(1);
});
