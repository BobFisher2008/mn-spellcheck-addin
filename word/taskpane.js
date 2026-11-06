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
function occurrenceIndex(haystack, needle, start) {
  let k = 0;
  let pos = 0;
  while (pos <= haystack.length) {
    const i = haystack.indexOf(needle, pos);
    if (i < 0 || i > start) return -1;
    if (i === start) return k;
    k++;
    pos = i + needle.length;
  }
  return -1;
}

// src/office/word/taskpane.ts
var engine = startEngine();
var has16 = false;
var has17 = false;
var has18 = false;
var items = [];
var checking = false;
var recheckQueued = false;
var changeTimer = 0;
var suggestionCache = /* @__PURE__ */ new Map();
var annotationsByPara = /* @__PURE__ */ new Map();
var annotationItem = /* @__PURE__ */ new Map();
var paraSignature = /* @__PURE__ */ new Map();
var $ = (id) => document.getElementById(id);
var isOn = (id) => $(id).checked;
function suggest(word) {
  let p = suggestionCache.get(word);
  if (!p) {
    p = engine.call("suggest", { word });
    suggestionCache.set(word, p);
  }
  return p;
}
function snippet(text, issue) {
  const a = Math.max(0, issue.start - 30);
  const b = Math.min(text.length, issue.end + 30);
  return (a > 0 ? "\u2026" : "") + text.slice(a, b).replace(/\s+/g, " ") + (b < text.length ? "\u2026" : "");
}
function setStatus(text) {
  $("status").textContent = text;
}
async function checkDocument() {
  if (checking) {
    recheckQueued = true;
    return;
  }
  checking = true;
  setStatus(UI.checking);
  try {
    await Word.run(async (ctx) => {
      const paras = ctx.document.body.paragraphs;
      paras.load(has16 ? "items/text,items/uniqueLocalId" : "items/text");
      await ctx.sync();
      const texts = paras.items.map((p) => p.text);
      const results = await engine.call("checkMany", { texts });
      items = [];
      results.forEach((issues, i) => {
        for (const issue of issues) {
          items.push({ paraIndex: i, paraId: has16 ? paras.items[i].uniqueLocalId : void 0, issue, context: snippet(texts[i], issue) });
        }
      });
      render();
      if (has17 && isOn("annotate")) await annotate(ctx, paras.items, results);
    });
    setStatus(items.length ? UI.errors(items.length) : UI.noErrors);
  } catch (e) {
    setStatus(`\u0410\u043B\u0434\u0430\u0430: ${String(e)}`);
    console.error(e);
  } finally {
    checking = false;
    if (recheckQueued) {
      recheckQueued = false;
      void checkDocument();
    }
  }
}
async function annotate(ctx, paras, results) {
  const changed = [];
  const seen = /* @__PURE__ */ new Set();
  paras.forEach((p, i) => {
    const id = p.uniqueLocalId;
    seen.add(id);
    const sig = results[i].map((x) => `${x.start}:${x.text}`).join("|");
    if (paraSignature.get(id) === sig) return;
    paraSignature.set(id, sig);
    changed.push({ p, id, list: items.filter((it) => it.paraId === id) });
  });
  for (const id of [...paraSignature.keys()]) if (!seen.has(id)) paraSignature.delete(id);
  if (changed.length === 0) return;
  for (const { id } of changed) {
    for (const annId of annotationsByPara.get(id) ?? []) {
      ctx.document.getAnnotationById(annId).delete();
      annotationItem.delete(annId);
    }
    annotationsByPara.delete(id);
  }
  await ctx.sync().catch(() => void 0);
  const inserted = [];
  for (const { p, id, list } of changed) {
    if (list.length === 0) continue;
    const critiques = [];
    for (const it of list) {
      const c = {
        colorScheme: Word.CritiqueColorScheme.red,
        start: it.issue.start,
        length: it.issue.end - it.issue.start
      };
      if (has18) {
        c.popupOptions = {
          brandingTextResourceId: "MN.Brand",
          subtitleResourceId: "MN.Subtitle",
          titleResourceId: "MN.Title",
          suggestions: (await suggest(it.issue.word)).slice(0, 3)
          // Word shows at most 3
        };
      }
      critiques.push(c);
    }
    inserted.push({ id, list, ids: p.insertAnnotations({ critiques }) });
  }
  await ctx.sync();
  for (const { id, list, ids } of inserted) {
    annotationsByPara.set(id, ids.value);
    ids.value.forEach((annId, j) => annotationItem.set(annId, list[j]));
  }
}
async function clearAnnotations() {
  if (!has17) return;
  await Word.run(async (ctx) => {
    for (const annId of annotationItem.keys()) ctx.document.getAnnotationById(annId).delete();
    await ctx.sync().catch(() => void 0);
  });
  annotationItem.clear();
  annotationsByPara.clear();
  paraSignature.clear();
}
async function locate(ctx, it) {
  let p;
  if (it.paraId && has16) {
    p = ctx.document.getParagraphByUniqueLocalId(it.paraId);
  } else {
    const paras = ctx.document.body.paragraphs;
    paras.load("items");
    await ctx.sync();
    const found = paras.items[it.paraIndex];
    if (!found) return null;
    p = found;
  }
  p.load("text");
  await ctx.sync();
  const { start, end, text } = it.issue;
  if (p.text.slice(start, end) !== text) return null;
  const k = occurrenceIndex(p.text, text, start);
  if (k < 0) return null;
  const hits = p.search(text, { matchCase: true });
  hits.load("items");
  await ctx.sync();
  return hits.items[k] ?? null;
}
async function replace(it, replacement) {
  await Word.run(async (ctx) => {
    const r = await locate(ctx, it);
    if (!r) {
      setStatus("\u0411\u0430\u0440\u0438\u043C\u0442 \u04E9\u04E9\u0440\u0447\u043B\u04E9\u0433\u0434\u0441\u04E9\u043D \u0431\u0430\u0439\u043D\u0430 \u2014 \u0434\u0430\u0445\u0438\u043D \u0448\u0430\u043B\u0433\u0430\u0436 \u0431\u0430\u0439\u043D\u0430\u2026");
      return;
    }
    r.insertText(replacement, Word.InsertLocation.replace);
    await ctx.sync();
  });
  await checkDocument();
}
async function selectInDocument(it) {
  await Word.run(async (ctx) => {
    const r = await locate(ctx, it);
    if (r) {
      r.select();
      await ctx.sync();
    }
  });
}
async function ignore(word) {
  await engine.call("ignore", { word });
  await checkDocument();
}
async function addWord(word) {
  personal.save([...personal.load(), word]);
  await engine.call("addWord", { word });
  await checkDocument();
}
function render() {
  const list = $("issues");
  list.replaceChildren();
  for (const it of items) {
    const sugg = el("div", { className: "suggestions" }, UI.checking);
    const word = el("button", { className: "word", title: "\u0411\u0430\u0440\u0438\u043C\u0442\u0430\u0430\u0441 \u043E\u043B\u043E\u0445 (select in document)", type: "button" }, it.issue.text);
    word.addEventListener("click", () => void selectInDocument(it));
    const ignoreBtn = el("button", { className: "secondary", type: "button" }, UI.ignore);
    ignoreBtn.addEventListener("click", () => void ignore(it.issue.word));
    const addBtn = el("button", { className: "secondary", type: "button" }, UI.add);
    addBtn.addEventListener("click", () => void addWord(it.issue.word));
    list.append(
      el("li", {}, word, el("div", { className: "context" }, it.context), sugg, el("div", { className: "actions" }, ignoreBtn, addBtn))
    );
    void suggest(it.issue.word).then((s) => {
      sugg.replaceChildren(
        ...s.length ? s.map((x) => {
          const b = el("button", { className: "suggestion", type: "button" }, x);
          b.addEventListener("click", () => void replace(it, x));
          return b;
        }) : [el("span", { className: "none" }, UI.noSuggestions)]
      );
    });
  }
}
function scheduleRecheck() {
  if (!isOn("live")) return;
  clearTimeout(changeTimer);
  changeTimer = window.setTimeout(() => void checkDocument(), 800);
}
async function onPopupAction(args) {
  const it = annotationItem.get(args.id);
  if (!it) return;
  if (args.action === "Accept" && args.critiqueSuggestion) {
    await Word.run(async (ctx) => {
      const ann = ctx.document.getAnnotationById(args.id);
      const r = ann.critiqueAnnotation.range;
      r.load("text");
      await ctx.sync();
      if (r.text === it.issue.text) {
        r.insertText(args.critiqueSuggestion, Word.InsertLocation.replace);
        await ctx.sync();
      }
    }).catch((e) => console.warn("[mn-spell] accept", e));
  } else if (args.action === "Reject") {
    await engine.call("ignore", { word: it.issue.word });
  }
  scheduleRecheck();
}
Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Word) {
    setStatus("\u042D\u043D\u044D \u0445\u0443\u0443\u0434\u0441\u044B\u0433 Word \u0434\u043E\u0442\u043E\u0440 \u043D\u044D\u044D\u043D\u044D \u04AF\u04AF.");
    return;
  }
  const req = Office.context.requirements;
  has16 = req.isSetSupported("WordApi", "1.6");
  has17 = req.isSetSupported("WordApi", "1.7");
  has18 = req.isSetSupported("WordApi", "1.8");
  $("caps").textContent = `WordApi: 1.6 ${has16 ? "\u2713" : "\u2717"} \xB7 1.7 ${has17 ? "\u2713" : "\u2717"} \xB7 1.8 ${has18 ? "\u2713" : "\u2717"}`;
  $("live").disabled = !has16;
  $("live").checked = has16;
  $("annotate").disabled = !has17;
  $("annotate").checked = has17;
  $("recheck").addEventListener("click", () => void checkDocument());
  $("annotate").addEventListener("change", () => {
    if (isOn("annotate")) {
      paraSignature.clear();
      void checkDocument();
    } else void clearAnnotations();
  });
  setStatus(UI.loadingEngine);
  await engine.call("setPersonalWords", { words: personal.load() });
  await checkDocument();
  if (has16) {
    await Word.run(async (ctx) => {
      ctx.document.onParagraphChanged.add(async () => scheduleRecheck());
      ctx.document.onParagraphAdded.add(async () => scheduleRecheck());
      ctx.document.onParagraphDeleted.add(async () => scheduleRecheck());
      if (has18) ctx.document.onAnnotationPopupAction.add(onPopupAction);
      await ctx.sync();
    });
  }
});
//# sourceMappingURL=taskpane.js.map
