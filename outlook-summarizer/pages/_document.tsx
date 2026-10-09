import { Html, Head, Main, NextScript } from 'next/document';

export default function Document() {
  return (
    <Html>
      <Head>
        {/* Office.js CDN */}
        <script
          type="text/javascript"
          src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"
        />
        {/* Add meta tag for content security policy */}
        <meta
          httpEquiv="Content-Security-Policy"
          content="default-src * 'self' 'unsafe-inline' 'unsafe-eval' data: https://appsforoffice.microsoft.com;"
        />
      </Head>
      <body>
        <Main />
        <NextScript />
      </body>
    </Html>
  );
}