import React, { useState, useEffect } from 'react';
import Head from 'next/head';
import styles from '../styles/Home.module.css';

const Home: React.FC = () => {
  const [summary, setSummary] = useState<string>('');
  const [isOfficeInitialized, setIsOfficeInitialized] = useState<boolean>(false);
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [isError, setIsError] = useState<boolean>(false);

  useEffect(() => {
    // Initialize Office.js
    const initializeOffice = async () => {
      try {
        if (typeof Office !== 'undefined') {
          // Using a Promise wrapper for better error handling
          await new Promise<void>((resolve, reject) => {
            try {
              Office.onReady((info) => {
                console.log(`Office.js is initialized. Host: ${info.host}`);
                setIsOfficeInitialized(true);
                resolve();
              });
            } catch (e) {
              reject(e);
            }
          });
        } else {
          console.error('Office.js is not available');
        }
      } catch (error) {
        console.error('Error initializing Office.js:', error);
      }
    };

    initializeOffice();
  }, []);

  const handleSummarizeEmail = async () => {
    console.log('Summarize button clicked');

    if (!isOfficeInitialized) {
      setIsError(true);
      setSummary('Office.js is not initialized yet.');
      return;
    }

    // Show loading state
    setIsLoading(true);
    setIsError(false);
    setSummary('');

    try {
      const item = Office.context.mailbox.item;
      if (!item) {
        setIsLoading(false);
        setIsError(true);
        setSummary('Open an email first, then click Summarize.');
        return;
      }

      // Get the email body content
      item.body.getAsync(
        'text',
        (result) => {
          setIsLoading(false);

          if (result.status === Office.AsyncResultStatus.Succeeded) {
            const emailBody = result.value;

            // For this example, we'll create a simple summary
            // In a real app, you might call an AI service here
            const simpleSummary = createSimpleSummary(emailBody);
            setSummary(simpleSummary);
          } else {
            setIsError(true);
            setSummary('Failed to read email content: ' + result.error.message);
          }
        }
      );
    } catch (error) {
      setIsLoading(false);
      setIsError(true);
      setSummary('Error: ' + (error instanceof Error ? error.message : String(error)));
    }
  };

  // Helper function to create a simple summary
  const createSimpleSummary = (text: string): string => {
    // Remove excess whitespace and limit length
    const cleanText = text.replace(/\s+/g, ' ').trim();

    if (cleanText.length === 0) {
      return 'This email appears to be empty.';
    }

    if (cleanText.length <= 150) {
      return 'Short email: ' + cleanText;
    }

    // Extract first 2-3 sentences for a simple summary
    const sentences = cleanText.split(/[.!?]+/).filter(s => s.trim().length > 0);

    if (sentences.length === 0) {
      return cleanText.substring(0, 200) + '...';
    }

    const firstSentences = sentences.slice(0, Math.min(3, sentences.length));
    let summary = firstSentences.join('. ');

    if (summary.length > 300) {
      summary = summary.substring(0, 297) + '...';
    }

    return summary + (sentences.length > 3 ? '...' : '');
  };

  return (
    <div className={styles.container}>
      <Head>
        <title>Email Summarizer</title>
        <meta name="description" content="Outlook Add-in for email summarization" />
        <meta name="viewport" content="width=device-width, initial-scale=1.0" />
        <link rel="icon" href="/favicon.ico" />
      </Head>

      <main className={styles.main}>
        <h1 className={styles.title}>Email Summarizer</h1>

        {!isOfficeInitialized ? (
          <div className={styles.error}>Loading Office.js components...</div>
        ) : (
          <>
            <div className={styles.buttonContainer}>
              <button
                className={styles.button}
                onClick={handleSummarizeEmail}
                disabled={isLoading}
              >
                {isLoading ? 'Processing...' : 'Summarize Email'}
              </button>
            </div>

            <div className={styles.summaryContainer}>
              {isLoading ? (
                <div className={styles.loading}>Analyzing email content...</div>
              ) : summary ? (
                <div className={isError ? styles.error : styles.summaryContent}>{summary}</div>
              ) : (
                <div className={styles.placeholder}>Click the button above to summarize the current email</div>
              )}
            </div>
          </>
        )}
      </main>
    </div>
  );
};

export default Home;