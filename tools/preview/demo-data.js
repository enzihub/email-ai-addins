// Invented demo mailbox used by the local preview hosts.
// Every person, company and address here is fictional (.example domains).
(function (root) {
  var DEMO = {
    me: { name: "Sam Rivera", email: "sam@fernway.example" },

    inbox: [
      { id: "m-venue", from: "Marta Kovač", subject: "Q3 offsite: venue shortlist, need your pick by Thursday", snippet: "Hi Sam, three venues made the cut. Harbourline Hall is the cheapest…", time: "9:42", unread: true },
      { id: "m-proofs", from: "Priya Natarajan", subject: "Spring catalogue: final proofs + print date", snippet: "Proofs v3 attached. Two fixes left on pages 14 and 22…", time: "8:15", unread: true },
      { id: "m-invoice", from: "Ledgerly Billing", subject: "Invoice INV-2291 is ready", snippet: "Your September invoice for 4 seats is now available…", time: "Yesterday" },
      { id: "m-standup", from: "Jonah Okafor", subject: "Re: Thursday stand-up moved to 10:30", snippet: "Works for me. I'll bring the shipping numbers…", time: "Yesterday" },
      { id: "m-trail", from: "Fernway Community", subject: "Trail clean-up day: 14 volunteers so far", snippet: "Thanks to everyone who signed up for Saturday…", time: "Mon" },
      { id: "m-quote", from: "Ana Lindqvist", subject: "Quote for 400 enamel mugs", snippet: "As promised, here is the quote with the two glaze options…", time: "Mon" }
    ],

    // The message the Outlook add-ins read.
    venue: {
      id: "m-venue",
      itemId: "AAMkADemo-venue-0001",
      subject: "Q3 offsite: venue shortlist, need your pick by Thursday",
      from: { displayName: "Marta Kovač", emailAddress: "marta@harbourline-events.example" },
      to: [{ displayName: "Sam Rivera", emailAddress: "sam@fernway.example" }],
      cc: [{ displayName: "Jonah Okafor", emailAddress: "jonah@fernway.example" }],
      date: "Tue 14 Oct, 9:42",
      body:
        "Hi Sam,\n\n" +
        "Thanks for sending the headcount. With 38 people confirmed, three venues made the cut for the Q3 offsite on 6–7 November.\n\n" +
        "Harbourline Hall is the cheapest at $4,200 for both days, but it has no breakout rooms. " +
        "The Granary has two breakout rooms and on-site catering for $5,900. " +
        "Pine Ridge Lodge is the nicest and includes rooms for the night, but it comes in at $8,300 and is 90 minutes out of town.\n\n" +
        "All three are holding the dates until Friday at noon, so I need your pick by Thursday. " +
        "Could you also confirm whether we need a projector in every room, and send me the dietary list from Jonah?\n\n" +
        "Happy to jump on a call tomorrow afternoon if that helps.\n\n" +
        "Best,\nMarta"
    },

    // The Gmail thread Email Buddy reads.
    thread: {
      threadId: "thread-proofs-7f3a91",
      subject: "Spring catalogue: final proofs + print date",
      messages: [
        {
          id: "msg-proofs-1",
          from: "Priya Natarajan <priya@larkspur-print.example>",
          date: "2025-03-03T09:10:00Z",
          body:
            "Hi Sam,\n\nProofs v3 for the spring catalogue are attached (48 pages). Two fixes are still open: the price on the Tarn rain shell on page 14 says $189 but your sheet says $179, and the photo credit on page 22 is missing.\n\nIf you can approve by Wednesday we can print on Monday 10 March and deliver 2,500 copies to the warehouse on Friday 14 March.\n\nPriya"
        },
        {
          id: "msg-proofs-2",
          from: "Sam Rivera <sam@fernway.example>",
          date: "2025-03-03T11:32:00Z",
          body:
            "Thanks Priya. $179 is correct, please fix page 14. Jonah is chasing the photo credit for page 22 and will send it today.\n\nOne change: can we bump the run to 3,000 copies? Let me know what that does to the price and the date.\n\nSam"
        },
        {
          id: "msg-proofs-3",
          from: "Priya Natarajan <priya@larkspur-print.example>",
          date: "2025-03-04T08:05:00Z",
          body:
            "Morning Sam,\n\n3,000 copies is fine. The total goes from $6,850 to $7,940 and the print date stays Monday 10 March. Delivery slips by one day to Saturday 15 March, so we would need someone at the warehouse to sign for it.\n\nOnce I have the photo credit and your written approval I will lock the files.\n\nPriya"
        }
      ]
    }
  };
  root.DEMO = DEMO;
  if (typeof module !== "undefined") module.exports = DEMO;
})(typeof window !== "undefined" ? window : globalThis);
