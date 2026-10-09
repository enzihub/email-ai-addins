// Apps Script stand-in for the local preview of Email Buddy (the Gmail add-on).
//
// gmail-email-buddy/GetContextualAddOn.gs is loaded unchanged after this file.
// The shim gives it just enough of CardService, GmailApp, PropertiesService,
// UrlFetchApp and Logger to run in a browser against the invented thread in
// demo-data.js, then draws the cards it builds in a Gmail-style side panel.
// The card styling is an approximation of Gmail's add-on renderer.
(function () {
  var DEMO = window.DEMO;
  var NS = "preview.gas.";

  // ---------- Logger ----------
  window.Logger = { log: function () { /* quiet in preview */ } };

  // ---------- PropertiesService (localStorage backed) ----------
  function props(scope) {
    return {
      getProperty: function (k) { return localStorage.getItem(NS + scope + k); },
      setProperty: function (k, v) { localStorage.setItem(NS + scope + k, String(v)); return this; },
      deleteProperty: function (k) { localStorage.removeItem(NS + scope + k); return this; }
    };
  }
  window.PropertiesService = {
    getUserProperties: function () { return props("user."); },
    getScriptProperties: function () { return props("script."); }
  };
  // Fake key so the add-on calls the local stand-in model. Not a real key.
  if (!localStorage.getItem(NS + "script.OPENAI_API_KEY")) {
    localStorage.setItem(NS + "script.OPENAI_API_KEY", ["sk", "local", "preview", "only", "demo"].join("-"));
  }

  // ---------- GmailApp ----------
  var thread = DEMO.thread;
  function wrapMessage(m) {
    return {
      getId: function () { return m.id; },
      getPlainBody: function () { return m.body; },
      getSubject: function () { return thread.subject; },
      getFrom: function () { return m.from; },
      getDate: function () { return new Date(m.date); },
      getThread: function () { return { getId: function () { return thread.threadId; } }; },
      reply: function (text) { window.parent.postMessage({ type: "gas-preview:sent", text: text }, "*"); },
      createDraftReply: function (text) { window.parent.postMessage({ type: "gas-preview:draft", text: text }, "*"); }
    };
  }
  window.GmailApp = {
    getMessageById: function (id) {
      var m = thread.messages.filter(function (x) { return x.id === id; })[0];
      return m ? wrapMessage(m) : null;
    },
    getThreadById: function (id) {
      if (id !== thread.threadId) return null;
      return { getId: function () { return id; }, getMessages: function () { return thread.messages.map(wrapMessage); } };
    }
  };

  // ---------- UrlFetchApp (synchronous, like the real one) ----------
  window.UrlFetchApp = {
    fetch: function (url, options) {
      var target = url.replace("https://api.openai.com", "/preview/openai");
      var xhr = new XMLHttpRequest();
      xhr.open((options && options.method) || "GET", target, false);
      xhr.setRequestHeader("Content-Type", (options && options.contentType) || "application/json");
      xhr.send(options && options.payload);
      return {
        getResponseCode: function () { return xhr.status; },
        getContentText: function () { return xhr.responseText; }
      };
    }
  };

  // ---------- CardService (builders produce plain objects) ----------
  function builder(kind, fields) {
    var o = { kind: kind };
    Object.keys(fields).forEach(function (name) {
      o[name] = function (v) { this["_" + fields[name]] = v; return this; };
    });
    return o;
  }
  window.CardService = {
    TextButtonStyle: { FILLED: "FILLED", TEXT: "TEXT", OUTLINED: "OUTLINED" },
    Icon: { TICKET: "TICKET" },
    newCardBuilder: function () {
      return {
        kind: "card", _sections: [],
        setHeader: function (h) { this._header = h; return this; },
        addSection: function (s) { this._sections.push(s); return this; },
        build: function () { return { header: this._header, sections: this._sections }; }
      };
    },
    newCardHeader: function () { return builder("header", { setTitle: "title", setSubtitle: "subtitle", setImageUrl: "image" }); },
    newCardSection: function () {
      return {
        kind: "section", _widgets: [],
        setHeader: function (h) { this._header = h; return this; },
        addWidget: function (w) { this._widgets.push(w); return this; }
      };
    },
    newTextParagraph: function () { return builder("text", { setText: "text" }); },
    newKeyValue: function () { return builder("kv", { setTopLabel: "top", setContent: "content", setIcon: "icon" }); },
    newTextInput: function () { return builder("input", { setFieldName: "name", setTitle: "title", setHint: "hint", setValue: "value", setMultiline: "multiline" }); },
    newButtonSet: function () {
      return { kind: "buttons", _buttons: [], addButton: function (b) { this._buttons.push(b); return this; } };
    },
    newTextButton: function () { return builder("button", { setText: "text", setOnClickAction: "action", setTextButtonStyle: "style" }); },
    newAction: function () { return builder("action", { setFunctionName: "fn", setParameters: "params" }); },
    newNotification: function () { return builder("notification", { setText: "text" }); },
    newNavigation: function () {
      return {
        _ops: [],
        pushCard: function (c) { this._ops.push(["push", c]); return this; },
        updateCard: function (c) { this._ops.push(["update", c]); return this; },
        popCard: function () { this._ops.push(["pop"]); return this; }
      };
    },
    newActionResponseBuilder: function () {
      return {
        setNavigation: function (n) { this._nav = n; return this; },
        setNotification: function (n) { this._note = n; return this; },
        build: function () { return { nav: this._nav, note: this._note }; }
      };
    }
  };

  // ---------- Renderer ----------
  var stack = [];
  var event = { gmail: { messageId: thread.messages[thread.messages.length - 1].id, threadId: thread.threadId } };

  function esc(s) {
    return String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  }
  // Apps Script text widgets accept a small HTML subset; allow just those tags.
  function richText(s) {
    var out = esc(s)
      .replace(/&lt;(\/?)(b|i|u)&gt;/g, "<$1$2>")
      .replace(/&lt;font color=\\?"(#?[a-zA-Z0-9]+)\\?"&gt;/g, '<font color="$1">')
      .replace(/&lt;\/font&gt;/g, "</font>")
      .replace(/\n/g, "<br>");
    // light markdown the model may return
    out = out.replace(/\*\*(.+?)\*\*/g, "<b>$1</b>");
    return out;
  }

  function renderCard(card) {
    var root = document.getElementById("card");
    root.innerHTML = "";
    var showBack = stack.length > 1;
    var head = document.createElement("div");
    head.className = "card-header";
    head.innerHTML =
      (showBack ? '<button class="back" aria-label="Back" data-back>' +
        '<svg viewBox="0 0 24 24" width="20" height="20"><path d="M20 11H7.8l5.6-5.6L12 4l-8 8 8 8 1.4-1.4L7.8 13H20z" fill="currentColor"/></svg></button>' : "") +
      '<div class="titles"><div class="title">' + esc(card.header && card.header._title) + "</div>" +
      (card.header && card.header._subtitle ? '<div class="subtitle">' + esc(card.header._subtitle) + "</div>" : "") +
      "</div>" +
      (card.header && card.header._image ? '<span class="err-dot" aria-hidden="true">!</span>' : "");
    root.appendChild(head);

    card.sections.forEach(function (sec) {
      var s = document.createElement("section");
      s.className = "card-section";
      if (sec._header) s.innerHTML = '<h3 class="section-header">' + esc(sec._header) + "</h3>";
      sec._widgets.forEach(function (w) { s.appendChild(renderWidget(w)); });
      root.appendChild(s);
    });
    var back = root.querySelector("[data-back]");
    if (back) back.onclick = function () { stack.pop(); renderCard(stack[stack.length - 1]); };
    root.scrollTop = root.scrollHeight;
    if (window.__onCardRendered) window.__onCardRendered();
  }

  function renderWidget(w) {
    var el = document.createElement("div");
    if (w.kind === "text") {
      el.className = "w-text";
      var t = w._text || "";
      if (/^<b>You:<\/b>/.test(t)) el.className += " chat-you";
      if (/^<b>AI:<\/b>/.test(t)) el.className += " chat-ai";
      el.innerHTML = richText(t);
    } else if (w.kind === "kv") {
      el.className = "w-kv";
      el.innerHTML = '<div class="kv-top">' + esc(w._top) + '</div><div class="kv-content">' + esc(w._content) + "</div>";
    } else if (w.kind === "input") {
      el.className = "w-input" + (w._multiline ? " multi" : "");
      var id = "f-" + w._name;
      var control = w._multiline
        ? '<textarea id="' + id + '" name="' + esc(w._name) + '" rows="8">' + esc(w._value || "") + "</textarea>"
        : '<input id="' + id + '" name="' + esc(w._name) + '" value="' + esc(w._value || "") + '" placeholder=" ">';
      el.innerHTML = control + '<label for="' + id + '">' + esc(w._title) + "</label>" +
        (w._hint ? '<div class="hint">' + esc(w._hint) + "</div>" : "");
    } else if (w.kind === "buttons") {
      el.className = "w-buttons";
      w._buttons.forEach(function (b) {
        var btn = document.createElement("button");
        btn.className = "btn " + (b._style === "FILLED" ? "filled" : "text");
        btn.textContent = b._text;
        btn.onclick = function () { runAction(b._action, btn); };
        el.appendChild(btn);
      });
    }
    return el;
  }

  function collectInputs() {
    var inputs = {};
    document.querySelectorAll("#card input[name], #card textarea[name]").forEach(function (f) {
      inputs[f.name] = [f.value];
    });
    return inputs;
  }

  function toast(text) {
    var t = document.getElementById("toast");
    t.textContent = text;
    t.classList.add("show");
    clearTimeout(t._h);
    t._h = setTimeout(function () { t.classList.remove("show"); }, 3200);
  }

  function runAction(action, btn) {
    var bar = document.getElementById("progress");
    bar.classList.add("on");
    document.querySelectorAll("#card button").forEach(function (b) { b.disabled = true; });
    var inputs = collectInputs();
    // Let the progress bar paint before the synchronous "server" call.
    setTimeout(function () {
      var ev = { parameters: action._params || {}, formInputs: inputs, gmail: event.gmail };
      var res;
      try { res = window[action._fn](ev); }
      catch (e) { res = null; console.error(e); }
      bar.classList.remove("on");
      apply(res);
    }, 60);
  }

  function apply(res) {
    if (!res) { renderCard(stack[stack.length - 1]); return; }
    if (res.nav) {
      res.nav._ops.forEach(function (op) {
        if (op[0] === "push") stack.push(op[1]);
        else if (op[0] === "update") stack[stack.length - 1] = op[1];
        else if (op[0] === "pop" && stack.length > 1) stack.pop();
      });
    }
    renderCard(stack[stack.length - 1]);
    if (res.note) toast(res.note._text);
  }

  window.GasPreview = {
    start: function () {
      var cards = window.getContextualAddOn(event);
      stack = [cards[0]];
      renderCard(stack[0]);
    },
    reset: function () {
      Object.keys(localStorage).forEach(function (k) { if (k.indexOf(NS + "user.") === 0) localStorage.removeItem(k); });
    }
  };
})();
