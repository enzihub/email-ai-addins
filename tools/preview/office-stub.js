// Minimal Office.js stand-in for the local preview.
//
// It implements only the parts of the Office JavaScript API that the two
// Outlook add-ins in this repo call, backed by the invented message in
// demo-data.js. It lets the real, built task panes run in a plain browser.
// It is NOT Microsoft's library and does not talk to any mailbox.
(function () {
  var DEMO = window.DEMO;
  var msg = DEMO.venue;
  var store = window.localStorage;
  var SETTINGS_KEY = "preview.roamingSettings";
  var settings = {};
  try { settings = JSON.parse(store.getItem(SETTINGS_KEY) || "{}"); } catch (e) {}
  // A fake key so the add-in skips its "configure your key" state. Not a real key.
  if (!settings.openai_api_key_v1) settings.openai_api_key_v1 = ["sk", "local", "preview", "only", "demo"].join("-");

  function ok(value) { return { status: "succeeded", value: value }; }

  var item = {
    itemType: "message",
    itemId: msg.itemId,
    subject: msg.subject,
    from: msg.from,
    to: msg.to,
    cc: msg.cc,
    body: {
      getAsync: function (coercion, options, cb) {
        if (typeof options === "function") cb = options;
        setTimeout(function () { cb(ok(msg.body)); }, 30);
      }
    },
    displayReplyAllForm: function (data) {
      var html = typeof data === "string" ? data : (data && data.htmlBody) || "";
      window.parent.postMessage({ type: "office-preview:replyAll", htmlBody: html, subject: "RE: " + msg.subject }, "*");
    },
    displayReplyForm: function (data) { item.displayReplyAllForm(data); },
    notificationMessages: { replaceAsync: function () {} }
  };

  var Office = {
    HostType: { Outlook: "Outlook" },
    CoercionType: { Text: "text", Html: "html" },
    AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
    EventType: { ItemChanged: "olkItemSelectedChanged" },
    MailboxEnums: {
      ItemType: { Message: "message", Appointment: "appointment" },
      ItemNotificationMessageType: { InformationalMessage: "informationalMessage" }
    },
    context: {
      host: "Outlook",
      mailbox: {
        item: item,
        addHandlerAsync: function (type, handler, cb) { if (cb) cb(ok()); }
      },
      roamingSettings: {
        get: function (k) { return settings[k]; },
        set: function (k, v) { settings[k] = v; },
        remove: function (k) { delete settings[k]; },
        saveAsync: function (cb) {
          store.setItem(SETTINGS_KEY, JSON.stringify(settings));
          if (cb) setTimeout(function () { cb(ok()); }, 20);
        }
      }
    },
    actions: { associate: function () {} },
    onReady: function (cb) {
      var info = { host: "Outlook", platform: "OfficeOnline" };
      var p = new Promise(function (resolve) { setTimeout(function () { resolve(info); }, 0); });
      if (cb) p.then(cb);
      return p;
    }
  };
  window.Office = Office;

  // Send the add-in's OpenAI calls to the local stand-in model instead.
  var realFetch = window.fetch.bind(window);
  window.fetch = function (url, init) {
    if (typeof url === "string" && url.indexOf("https://api.openai.com/") === 0) {
      url = url.replace("https://api.openai.com", "/preview/openai");
    }
    return realFetch(url, init);
  };
})();
