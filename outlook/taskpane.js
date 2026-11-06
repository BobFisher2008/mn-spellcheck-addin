// src/extension/content/textmap.ts
var BLOCK_TAGS = /* @__PURE__ */ new Set([
  "ADDRESS",
  "ARTICLE",
  "ASIDE",
  "BLOCKQUOTE",
  "BR",
  "DD",
  "DIV",
  "DL",
  "DT",
  "FIGCAPTION",
  "FIGURE",
  "FOOTER",
  "FORM",
  "H1",
  "H2",
  "H3",
  "H4",
  "H5",
  "H6",
  "HEADER",
  "HR",
  "LI",
  "MAIN",
  "NAV",
  "OL",
  "P",
  "PRE",
  "SECTION",
  "TABLE",
  "TD",
  "TH",
  "TR",
  "UL"
]);
function isSkippedElement(el2) {
  const tag = el2.tagName;
  if (tag === "SCRIPT" || tag === "STYLE" || tag === "CODE" || tag === "PRE" || tag === "KBD" || tag === "SAMP") return true;
  if (el2.getAttribute("contenteditable") === "false") return true;
  if (el2.getAttribute("spellcheck") === "false") return true;
  if (el2.getAttribute("data-mn-spell") === "off") return true;
  if (tag === "BLOCKQUOTE") return true;
  const cls = typeof el2.className === "string" ? el2.className : "";
  if (cls.includes("gmail_quote")) return true;
  if (el2.id === "divRplyFwdMsg" || el2.id === "appendonsend") return true;
  return false;
}
function buildTextMap(root) {
  let text = "";
  const segments = [];
  const walker = root.ownerDocument.createTreeWalker(root, NodeFilter.SHOW_ELEMENT | NodeFilter.SHOW_TEXT, {
    acceptNode(n2) {
      if (n2.nodeType === Node.ELEMENT_NODE && n2 !== root && isSkippedElement(n2)) {
        return NodeFilter.FILTER_REJECT;
      }
      return NodeFilter.FILTER_ACCEPT;
    }
  });
  const separate = () => {
    if (text.length > 0 && !text.endsWith("\n")) text += "\n";
  };
  let n = walker.currentNode;
  while (n = walker.nextNode()) {
    if (n.nodeType === Node.TEXT_NODE) {
      const data = n.data;
      if (data.length === 0) continue;
      segments.push({ node: n, start: text.length, end: text.length + data.length });
      text += data;
    } else if (BLOCK_TAGS.has(n.tagName)) {
      separate();
    }
  }
  return { text, segments };
}

// src/core/protocol.ts
var RpcClient = class {
  nextId = 1;
  pending = /* @__PURE__ */ new Map();
  send;
  constructor(send) {
    this.send = send;
  }
  /** Feed every response received from the transport into this. */
  receive(msg) {
    const p = this.pending.get(msg.id);
    if (!p) return;
    this.pending.delete(msg.id);
    if (msg.ok) p.resolve(msg.result);
    else p.reject(new Error(msg.error));
  }
  call(method, params) {
    const id = this.nextId++;
    return new Promise((resolve, reject) => {
      this.pending.set(id, { resolve, reject });
      this.send({ id, method, params });
    });
  }
};

// src/office/common.ts
function startEngine() {
  const worker = new Worker(new URL("../engine/worker.js", location.href), { type: "module" });
  const client = new RpcClient((m) => worker.postMessage(m));
  worker.onmessage = (e) => client.receive(e.data);
  worker.onerror = (e) => console.error("[mn-spell] engine worker error", e.message);
  return client;
}
var PERSONAL_KEY = "mn-spell.personalWords";
var personal = {
  load() {
    try {
      const v = JSON.parse(localStorage.getItem(PERSONAL_KEY) ?? "[]");
      return Array.isArray(v) ? v.filter((x) => typeof x === "string") : [];
    } catch {
      return [];
    }
  },
  save(words) {
    try {
      localStorage.setItem(PERSONAL_KEY, JSON.stringify([...new Set(words)].sort()));
    } catch {
    }
  }
};
var UI = {
  checking: "\u0428\u0430\u043B\u0433\u0430\u0436 \u0431\u0430\u0439\u043D\u0430\u2026",
  noErrors: "\u0410\u043B\u0434\u0430\u0430 \u043E\u043B\u0434\u0441\u043E\u043D\u0433\u04AF\u0439.",
  errors: (n) => `${n} \u0430\u043B\u0434\u0430\u0430 \u043E\u043B\u0434\u043B\u043E\u043E`,
  ignore: "\u0410\u043B\u0433\u0430\u0441\u0430\u0445",
  add: "\u0422\u043E\u043B\u0438\u043D\u0434 \u043D\u044D\u043C\u044D\u0445",
  noSuggestions: "\u0421\u0430\u043D\u0430\u043B \u0431\u043E\u043B\u0433\u043E\u0445 \u04AF\u0433 \u0430\u043B\u0433\u0430",
  loadingEngine: "\u0422\u043E\u043B\u044C \u0430\u0447\u0430\u0430\u043B\u0436 \u0431\u0430\u0439\u043D\u0430\u2026"
};
function el(tag, props = {}, ...children) {
  const e = document.createElement(tag);
  Object.assign(e, props);
  e.append(...children);
  return e;
}

// src/office/outlook/taskpane.ts
var engine = startEngine();
var items = [];
var selected = "";
var selectedSource = "";
var pollTimer = 0;
var suggestionCache = /* @__PURE__ */ new Map();
var $ = (id) => document.getElementById(id);
var asyncOf = (fn) => new Promise(
  (resolve, reject) => fn((r) => r.status === Office.AsyncResultStatus.Succeeded ? resolve(r.value) : reject(r.error))
);
function compose() {
  return Office.context.mailbox.item;
}
function suggest(word) {
  let p = suggestionCache.get(word);
  if (!p) {
    p = engine.call("suggest", { word });
    suggestionCache.set(word, p);
  }
  return p;
}
function setStatus(t) {
  $("status").textContent = t;
}
async function bodyText() {
  const item = compose();
  const opts = {};
  if (Office.context.requirements.isSetSupported("Mailbox", "1.14")) {
    opts.bodyMode = Office.MailboxEnums.BodyMode.HostConfig;
  }
  const html = await asyncOf((cb) => item.body.getAsync(Office.CoercionType.Html, opts, cb));
  const doc = new DOMParser().parseFromString(html, "text/html");
  for (const marker of ["appendonsend", "divRplyFwdMsg"]) {
    const m = doc.getElementById(marker);
    if (m) {
      let n = m;
      while (n && n !== doc.body) {
        while (n.nextSibling) n.parentNode.removeChild(n.nextSibling);
        n = n.parentNode;
      }
      m.remove();
    }
  }
  return buildTextMap(doc.body).text;
}
function snippet(text, issue) {
  const a = Math.max(0, issue.start - 30);
  const b = Math.min(text.length, issue.end + 30);
  return (a > 0 ? "\u2026" : "") + text.slice(a, b).replace(/\s+/g, " ") + (b < text.length ? "\u2026" : "");
}
async function check() {
  setStatus(UI.checking);
  try {
    const item = compose();
    const [body, subject] = await Promise.all([
      bodyText(),
      asyncOf((cb) => item.subject.getAsync(cb))
    ]);
    const [bodyIssues, subjectIssues] = await engine.call("checkMany", { texts: [body, subject] });
    items = [
      ...subjectIssues.map((issue) => ({ source: "subject", issue, context: snippet(subject, issue) })),
      ...bodyIssues.map((issue) => ({ source: "body", issue, context: snippet(body, issue) }))
    ];
    render();
    setStatus(items.length ? UI.errors(items.length) : UI.noErrors);
  } catch (e) {
    setStatus(`\u0410\u043B\u0434\u0430\u0430: ${String(e?.message ?? e)}`);
  }
}
async function readSelection() {
  const r = await asyncOf(
    (cb) => compose().getSelectedDataAsync(Office.CoercionType.Text, cb)
  );
  return { text: r.data ?? "", source: r.sourceProperty ?? "" };
}
async function pollSelection() {
  try {
    const s = await readSelection();
    if (s.text !== selected || s.source !== selectedSource) {
      selected = s.text;
      selectedSource = s.source;
      updateSelectionState();
    }
  } catch {
  }
}
function matchesSelection(it) {
  return selected.trim() === it.issue.text && selectedSource === it.source;
}
function updateSelectionState() {
  for (const li of $("issues").querySelectorAll("li")) {
    const it = items[Number(li.dataset.i)];
    const on = !!it && matchesSelection(it);
    li.classList.toggle("selected", on);
    for (const b of li.querySelectorAll("button.suggestion")) b.disabled = !on;
  }
}
async function replaceSelection(it, replacement) {
  const s = await readSelection();
  if (s.text.trim() !== it.issue.text || s.source !== it.source) {
    setStatus(`"${it.issue.text}" \u04AF\u0433\u0438\u0439\u0433 \u043C\u0435\u0441\u0441\u0435\u0436 \u0434\u043E\u0442\u0440\u043E\u043E\u0441 \u0441\u043E\u043D\u0433\u043E\u043D\u043E \u0443\u0443.`);
    return;
  }
  const lead = s.text.match(/^\s*/)[0];
  const trail = s.text.match(/\s*$/)[0];
  await asyncOf(
    (cb) => compose().setSelectedDataAsync(lead + replacement + trail, { coercionType: Office.CoercionType.Text }, cb)
  );
  await check();
}
function render() {
  const list = $("issues");
  list.replaceChildren();
  items.forEach((it, i) => {
    const sugg = el("div", { className: "suggestions" }, UI.checking);
    const ignoreBtn = el("button", { className: "secondary", type: "button" }, UI.ignore);
    ignoreBtn.addEventListener("click", async () => {
      await engine.call("ignore", { word: it.issue.word });
      await check();
    });
    const addBtn = el("button", { className: "secondary", type: "button" }, UI.add);
    addBtn.addEventListener("click", async () => {
      personal.save([...personal.load(), it.issue.word]);
      await engine.call("addWord", { word: it.issue.word });
      await check();
    });
    const where = it.source === "subject" ? "\u0413\u0430\u0440\u0447\u0438\u0433 \xB7 " : "";
    const li = el(
      "li",
      {},
      el("span", { className: "word" }, it.issue.text),
      el("div", { className: "context" }, where + it.context),
      sugg,
      el("div", { className: "hint" }, "\u0417\u0430\u0441\u0430\u0445\u044B\u043D \u0442\u0443\u043B\u0434 \u04AF\u0433\u0438\u0439\u0433 \u043C\u0435\u0441\u0441\u0435\u0436 \u0434\u043E\u0442\u0440\u043E\u043E\u0441 \u0441\u043E\u043D\u0433\u043E\u043D\u043E \u0443\u0443 (\u0445\u043E\u0451\u0440 \u0442\u043E\u0432\u0448\u0438\u043D\u043E)."),
      el("div", { className: "actions" }, ignoreBtn, addBtn)
    );
    li.dataset.i = String(i);
    list.append(li);
    void suggest(it.issue.word).then((s) => {
      sugg.replaceChildren(
        ...s.length ? s.map((x) => {
          const b = el("button", { className: "suggestion", type: "button", disabled: !matchesSelection(it) }, x);
          b.addEventListener("click", () => void replaceSelection(it, x));
          return b;
        }) : [el("span", { className: "none" }, UI.noSuggestions)]
      );
    });
  });
}
Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Outlook) {
    setStatus("\u042D\u043D\u044D \u0445\u0443\u0443\u0434\u0441\u044B\u0433 Outlook \u0434\u043E\u0442\u043E\u0440 \u043D\u044D\u044D\u043D\u044D \u04AF\u04AF.");
    return;
  }
  $("recheck").addEventListener("click", () => void check());
  setStatus(UI.loadingEngine);
  await engine.call("setPersonalWords", { words: personal.load() });
  await check();
  pollTimer = window.setInterval(() => void pollSelection(), 1e3);
  document.addEventListener("visibilitychange", () => {
    if (document.hidden) clearInterval(pollTimer);
    else pollTimer = window.setInterval(() => void pollSelection(), 1e3);
  });
});
//# sourceMappingURL=taskpane.js.map
