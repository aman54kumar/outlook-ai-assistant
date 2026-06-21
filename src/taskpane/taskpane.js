// taskpane.js
// Wires the UI to Office.js (read the email, insert results) and to the local
// /api/run endpoint (which does the actual AI work server-side).

const RUN_LABEL   = { summarize: "Summarize thread", reply: "Draft reply", tone: "Rewrite draft" };
const INSERT_LABEL = { summarize: "Insert", reply: "Set as draft", tone: "Replace draft" };

// Per-provider hardcoded defaults — base URL and recommended default model.
// Model IDs verified from official docs, June 2026.
const PROVIDER_DEFAULTS = {
  openai:    { baseUrl: "https://api.openai.com/v1",                                 model: "gpt-5.5" },
  anthropic: { baseUrl: "https://api.anthropic.com",                                 model: "claude-sonnet-4-6" },
  google:    { baseUrl: "https://generativelanguage.googleapis.com/v1beta/openai",   model: "gemini-2.5-flash" },
  custom:    { baseUrl: "",                                                            model: "" },
};

// Model lists shown in the datalist — verified API model IDs as of June 2026.
const MODEL_SUGGESTIONS = {
  // OpenAI: gpt-5.x is the current generation; o3/o4-mini are reasoning models.
  // gpt-4o / gpt-4.1 series still active but scheduled for deprecation Oct 2026.
  openai: [
    "gpt-5.5", "gpt-5.5-pro",
    "gpt-5.4", "gpt-5.4-mini", "gpt-5.4-nano",
    "o3", "o3-pro", "o4-mini",
    "gpt-4o", "gpt-4o-mini",
    "gpt-4.1", "gpt-4.1-mini", "gpt-4.1-nano",
  ],
  // Anthropic: 4.6 generation uses dateless IDs (e.g. claude-sonnet-4-6).
  // Earlier models use dated snapshots or -latest aliases.
  anthropic: [
    "claude-fable-5",
    "claude-opus-4-8", "claude-opus-4-7", "claude-opus-4-6",
    "claude-sonnet-4-6", "claude-sonnet-4-5",
    "claude-haiku-4-5",
    "claude-3-5-sonnet-latest", "claude-3-5-haiku-latest",
  ],
  // Google Gemini: use short stable names; gemini-3.5-flash is the latest stable.
  // Accessed via the OpenAI-compatible endpoint.
  google: [
    "gemini-3.5-flash",
    "gemini-3.1-flash-lite",
    "gemini-2.5-pro", "gemini-2.5-flash", "gemini-2.5-flash-lite",
    "gemini-2.0-flash", "gemini-2.0-flash-lite",
  ],
  custom: [],
};

// Runtime state (not persisted).
const state = {
  action:    "summarize",
  tone:      "Professional",
  provider:  "local",
  isCompose: false,
  lastResult: "",  // last AI output (Markdown)
};

// Settings persisted to localStorage so they survive pane reloads.
// Per-provider keys mean switching providers never wipes your other keys.
const settings = {
  provider:         "local",   // mirrors state.provider
  cloudProvider:    "openai",
  // Per-provider API keys
  apiKey_openai:    "",
  apiKey_anthropic: "",
  apiKey_google:    "",
  apiKey_custom:    "",
  // Per-provider model overrides (empty = PROVIDER_DEFAULTS fallback)
  model_openai:     "",
  model_anthropic:  "",
  model_google:     "",
  model_custom:     "",
  // Per-provider base URL overrides (empty = PROVIDER_DEFAULTS fallback)
  baseUrl_openai:   "",
  baseUrl_anthropic: "",
  baseUrl_google:   "",
  baseUrl_custom:   "",
  // Local
  ollamaModel:      "",
  ollamaBaseUrl:    "",
};

const $ = (id) => document.getElementById(id);

// ── Persistence ───────────────────────────────────────────────────────────────

function loadSettings() {
  try {
    const saved = localStorage.getItem("ama-settings");
    if (!saved) return;
    const data = JSON.parse(saved);
    Object.assign(settings, data);

    // ── Migrate from old single-key format (pre-per-provider) ──────────────
    // If user had a key/model/baseUrl saved with the old schema, carry it forward
    // to the appropriate per-provider slot so they don't lose it.
    if (data.apiKey && !data.apiKey_openai && !data.apiKey_anthropic) {
      settings[`apiKey_${settings.cloudProvider || "openai"}`] = data.apiKey;
    }
    if (data.cloudModel && !data.model_openai) {
      settings[`model_${settings.cloudProvider || "openai"}`] = data.cloudModel;
    }
    if (data.cloudBaseUrl && !data.baseUrl_custom) {
      settings[`baseUrl_${settings.cloudProvider || "custom"}`] = data.cloudBaseUrl;
    }
  } catch { /* first run or corrupt */ }
}

function saveSettings() {
  try { localStorage.setItem("ama-settings", JSON.stringify(settings)); } catch {}
}

// ── Boot ──────────────────────────────────────────────────────────────────────

Office.onReady((info) => {
  if (info.host !== Office.HostType.Outlook) {
    $("boot").textContent = "This add-in runs inside Outlook.";
    return;
  }
  state.isCompose =
    typeof Office.context.mailbox.item?.body?.setSelectedDataAsync === "function";
  loadSettings();
  initUI();
  loadSource();
  $("boot").hidden = true;
  $("pane").hidden = false;
});

// ── UI init ───────────────────────────────────────────────────────────────────

function initUI() {
  // Tabs
  document.querySelectorAll(".tab").forEach((tab) => {
    tab.addEventListener("click", () => setAction(tab.dataset.action));
  });

  // Tone chips
  const tones = $("tones");
  [
    "Professional", "Friendly", "Concise", "Assertive", "Apologetic",
    "Government", "Formal", "Diplomatic", "Direct", "Empathetic", "Casual",
  ].forEach((t, i) => {
    const b = document.createElement("button");
    b.className = "tone" + (i === 0 ? " is-on" : "");
    b.textContent = t;
    b.type = "button";
    b.addEventListener("click", () => {
      state.tone = t;
      document.querySelectorAll(".tone").forEach((x) => x.classList.toggle("is-on", x === b));
    });
    tones.appendChild(b);
  });

  // Run + result buttons
  $("runBtn").addEventListener("click", run);
  $("copyBtn").addEventListener("click", copyResult);
  $("insertBtn").addEventListener("click", insertResult);

  // Runtime chip opens drawer
  $("runtimeChip").addEventListener("click", openDrawer);
  $("drawerClose").addEventListener("click", () => ($("drawer").hidden = true));

  // Local / Cloud segment
  document.querySelectorAll(".seg__opt[data-provider]").forEach((opt) => {
    opt.addEventListener("click", () => setProvider(opt.dataset.provider));
  });

  // Cloud provider segment
  document.querySelectorAll(".seg__opt[data-cloud-provider]").forEach((opt) => {
    opt.addEventListener("click", () => setCloudProvider(opt.dataset.cloudProvider));
  });

  // Show / hide API key
  $("toggleKeyBtn").addEventListener("click", () => {
    const inp = $("apiKeyInput");
    inp.type = inp.type === "password" ? "text" : "password";
  });

  // Cloud field listeners — write into per-provider slots
  $("apiKeyInput").addEventListener("input", () => {
    settings[`apiKey_${settings.cloudProvider}`] = $("apiKeyInput").value;
    saveSettings();
    updateCloudStatus();
    scheduleFetchModels(settings.cloudProvider);  // re-fetch after user stops typing
  });
  $("cloudModelInput").addEventListener("input", () => {
    settings[`model_${settings.cloudProvider}`] = $("cloudModelInput").value;
    saveSettings();
  });
  $("cloudBaseUrlInput").addEventListener("input", () => {
    settings[`baseUrl_${settings.cloudProvider}`] = $("cloudBaseUrlInput").value;
    saveSettings();
    updateBaseUrlReset(settings.cloudProvider);
  });
  $("baseUrlReset").addEventListener("click", () => {
    const cp = settings.cloudProvider;
    const def = PROVIDER_DEFAULTS[cp]?.baseUrl ?? "";
    $("cloudBaseUrlInput").value = def;
    settings[`baseUrl_${cp}`] = "";   // empty = "use default"
    saveSettings();
    updateBaseUrlReset(cp);
  });

  // Local field listeners
  $("ollamaModelInput").addEventListener("input", () => {
    settings.ollamaModel = $("ollamaModelInput").value;
    saveSettings();
  });
  $("ollamaUrlInput").addEventListener("input", () => {
    settings.ollamaBaseUrl = $("ollamaUrlInput").value;
    saveSettings();
  });

  applySettingsToUI();
  setAction("summarize");
}

function applySettingsToUI() {
  $("ollamaModelInput").value  = settings.ollamaModel;
  $("ollamaUrlInput").value    = settings.ollamaBaseUrl;
  // setCloudProvider fills the API key, model, and base URL fields for the active provider
  setCloudProvider(settings.cloudProvider, false);
  // Restore the Local/Cloud toggle without triggering another saveSettings()
  if (settings.provider === "cloud") {
    state.provider = "cloud";
    document.querySelectorAll(".seg__opt[data-provider]").forEach((o) =>
      o.classList.toggle("is-on", o.dataset.provider === "cloud")
    );
    $("runtimeChip").classList.add("is-cloud");
    $("runtimeLabel").textContent = "Cloud";
  }
}

// ── Provider & settings UI ────────────────────────────────────────────────────

function openDrawer() {
  syncDrawerSections();
  $("drawer").hidden = false;
}

function syncDrawerSections() {
  const isCloud = state.provider === "cloud";
  $("localSettings").hidden  = isCloud;
  $("cloudSettings").hidden  = !isCloud;
}

function setProvider(provider) {
  state.provider = provider;
  settings.provider = provider;  // keep in sync so it survives pane reload
  saveSettings();
  document.querySelectorAll(".seg__opt[data-provider]").forEach((o) =>
    o.classList.toggle("is-on", o.dataset.provider === provider)
  );
  const cloud = provider === "cloud";
  $("runtimeChip").classList.toggle("is-cloud", cloud);
  $("runtimeLabel").textContent = cloud ? "Cloud" : "Local";
  syncDrawerSections();
}

function setCloudProvider(cp, save = true) {
  settings.cloudProvider = cp;
  if (save) saveSettings();

  // Highlight the right segment button
  document.querySelectorAll(".seg__opt[data-cloud-provider]").forEach((o) =>
    o.classList.toggle("is-on", o.dataset.cloudProvider === cp)
  );

  const def = PROVIDER_DEFAULTS[cp] ?? { baseUrl: "", model: "" };

  // Fill per-provider values; fall back to PROVIDER_DEFAULTS if not yet set
  $("apiKeyInput").value       = settings[`apiKey_${cp}`]    || "";
  $("cloudModelInput").value   = settings[`model_${cp}`]     || def.model;
  $("cloudBaseUrlInput").value = settings[`baseUrl_${cp}`]   || def.baseUrl;

  // Seed the datalist with the static fallback while the live fetch runs
  rebuildModelList(cp);

  $("googleNote").hidden = cp !== "google";
  updateBaseUrlReset(cp);
  updateCloudStatus();
  // Fetch live models if the key is already set for this provider
  if (settings[`apiKey_${cp}`]) scheduleFetchModels(cp, 0);
}

/** Shows the Reset button only when the base URL differs from the provider default. */
function updateBaseUrlReset(cp) {
  const def     = PROVIDER_DEFAULTS[cp]?.baseUrl ?? "";
  const current = $("cloudBaseUrlInput").value;
  $("baseUrlReset").hidden = current === def || (current === "" && def === "");
}

function updateCloudStatus() {
  const el = $("cloudStatus");
  const cp = settings.cloudProvider;
  const hasKey = (settings[`apiKey_${cp}`] || "").length > 0;
  el.textContent = hasKey ? "✓ API key entered" : "No key entered — will use server .env if set";
  el.className = "key-status " + (hasKey ? "key-status--ok" : "key-status--warn");
}

// ── Dynamic model fetching ────────────────────────────────────────────────────

let _modelFetchTimer = null;

/**
 * Schedules a model fetch with an optional debounce delay (default 900 ms).
 * Passing delay=0 fires immediately (used when switching providers with a key set).
 */
function scheduleFetchModels(cp, delay = 900) {
  clearTimeout(_modelFetchTimer);
  _modelFetchTimer = setTimeout(() => fetchModels(cp), delay);
}

/**
 * Calls /api/models on the server (which proxies the provider's model list
 * endpoint) and populates the datalist for the given cloud provider.
 * Falls back to MODEL_SUGGESTIONS if the fetch fails or the key is missing.
 */
async function fetchModels(cp) {
  const apiKey    = settings[`apiKey_${cp}`] || "";
  const baseUrl   = settings[`baseUrl_${cp}`] || PROVIDER_DEFAULTS[cp]?.baseUrl || "";
  const statusEl  = $("modelStatus");
  const list      = $("modelList");

  // Nothing to do if there's no key yet
  if (!apiKey) {
    statusEl.textContent = "";
    rebuildModelList(cp);
    return;
  }

  statusEl.textContent = "Loading models…";
  statusEl.className   = "model-status model-status--loading";

  try {
    const res = await fetch("/api/models", {
      method:  "POST",
      headers: { "Content-Type": "application/json" },
      body:    JSON.stringify({ cloudProvider: cp, apiKey, cloudBaseUrl: baseUrl }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Could not fetch models.");

    const models = data.models || [];
    if (!models.length) throw new Error("No chat models returned.");

    // Only update if the user hasn't switched away from this provider
    if (settings.cloudProvider !== cp) return;

    list.innerHTML = "";
    models.forEach((m) => {
      const opt = document.createElement("option");
      opt.value = m;
      list.appendChild(opt);
    });

    statusEl.textContent = `${models.length} models loaded`;
    statusEl.className   = "model-status model-status--ok";
    setTimeout(() => {
      if (statusEl.textContent.includes("loaded")) statusEl.textContent = "";
    }, 3000);
  } catch (err) {
    // Key wrong, network error, etc. — fall back to static list silently
    statusEl.textContent = "Could not load — using defaults";
    statusEl.className   = "model-status model-status--err";
    rebuildModelList(cp);
  }
}

/** Rebuilds the datalist from the static MODEL_SUGGESTIONS fallback. */
function rebuildModelList(cp) {
  const list = $("modelList");
  list.innerHTML = "";
  (MODEL_SUGGESTIONS[cp] ?? []).forEach((m) => {
    const opt = document.createElement("option");
    opt.value = m;
    list.appendChild(opt);
  });
}

// ── Email body processing ─────────────────────────────────────────────────────

/**
 * Reads the email body, strips signature/quoted content, returns plain text
 * for the AI. Tries HTML first for better noise removal; falls back to plain
 * text if the Office.js item doesn't support HTML coercion.
 */
async function getBody() {
  try {
    const html  = await getRawBodyHtml();
    const clean = stripEmailNoise(html);
    const text  = htmlToText(clean);
    // If stripping was too aggressive and removed everything, fall back to
    // the full plain-text body rather than sending the AI an empty string.
    if (text.trim()) return text;
    return htmlToText(html);
  } catch {
    // HTML coercion not supported (old Outlook / plain-text message) — use text.
    return new Promise((resolve, reject) => {
      Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, (r) => {
        if (r.status === Office.AsyncResultStatus.Succeeded) resolve((r.value || "").trim());
        else reject(new Error(r.error?.message || "Could not read the email body."));
      });
    });
  }
}

/**
 * Like getBody() but preserves the quoted email thread (strips only signatures).
 * Used by the reply action so the AI sees the full conversation history.
 */
async function getBodyWithThread() {
  try {
    const html  = await getRawBodyHtml();
    const clean = stripSignaturesOnly(html);
    const text  = htmlToText(clean);
    if (text.trim()) return text;
    return htmlToText(html);
  } catch {
    return new Promise((resolve, reject) => {
      Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, (r) => {
        if (r.status === Office.AsyncResultStatus.Succeeded) resolve((r.value || "").trim());
        else reject(new Error(r.error?.message || "Could not read the email body."));
      });
    });
  }
}

/**
 * Strips only signature blocks and the RFC 3676 "-- " delimiter,
 * leaving quoted/forwarded thread content intact.
 */
function stripSignaturesOnly(html) {
  if (!html) return "";
  const doc  = new DOMParser().parseFromString(html, "text/html");
  const body = doc.body;

  // Remove known signature containers
  [
    "#Signature", "#signature", "#x_Signature", "#divSignature",
    ".MsoSignature", ".gmail_signature", ".apple-mail-signature",
    "[data-signatureid]", ".x_gmail_signature",
  ].forEach((sel) => {
    try { body.querySelectorAll(sel).forEach((el) => el.remove()); } catch {}
  });

  // Strip RFC 3676 "-- " signature delimiter and everything after it
  const topChildren = [...body.children];
  let cutFrom = -1;
  for (let i = 0; i < topChildren.length; i++) {
    const txt = topChildren[i].textContent.replace(/\u00a0/g, " ").trim();
    if (txt === "--" || txt === "-- ") { cutFrom = i; break; }
  }
  if (cutFrom !== -1) topChildren.slice(cutFrom).forEach((el) => el.remove());

  return body.innerHTML;
}

/**
 * Returns the raw HTML of the current message body.
 */
function getRawBodyHtml() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, (r) => {
      if (r.status === Office.AsyncResultStatus.Succeeded) resolve(r.value || "");
      else reject(new Error(r.error?.message || "Could not read the email body."));
    });
  });
}

/**
 * Strips email signature blocks and quoted/forwarded content from HTML.
 * Best-effort: handles Outlook Web, Gmail, RFC 3676 "--" delimiter.
 */
function stripEmailNoise(html) {
  if (!html) return "";
  const doc = new DOMParser().parseFromString(html, "text/html");
  const body = doc.body;

  // Remove known signature containers
  [
    "#Signature", "#signature", "#x_Signature", "#divSignature",
    ".MsoSignature", ".gmail_signature", ".apple-mail-signature",
    "[data-signatureid]", ".x_gmail_signature",
  ].forEach((sel) => {
    try { body.querySelectorAll(sel).forEach((el) => el.remove()); } catch {}
  });

  // Remove Outlook / Gmail quoted / forwarded sections
  [
    ".gmail_quote", ".gmail_extra", ".yahoo_quoted",
    "[id^='divRplyFwdMsg']", ".moz-cite-prefix", "[id^='appendonsend']",
  ].forEach((sel) => {
    try { body.querySelectorAll(sel).forEach((el) => el.remove()); } catch {}
  });

  // Remove <hr> reply/forward separators ONLY when followed by email header text
  // (e.g. "From:", "Sent:", "To:") — avoids stripping content-only dividers.
  body.querySelectorAll("hr").forEach((hr) => {
    // Collect text from the next ~200 chars after the <hr>
    let sample = "";
    let node = hr.nextSibling;
    while (node && sample.length < 200) {
      sample += (node.textContent || "");
      node = node.nextSibling;
    }
    const looksLikeQuote = /\b(From|Sent|To|Subject|Date)\s*:/i.test(sample);
    if (looksLikeQuote) {
      let next = hr.nextSibling;
      while (next) { const n = next.nextSibling; next.remove(); next = n; }
      hr.remove();
    }
  });

  // Strip RFC 3676 "-- " signature delimiter and everything after it
  const topChildren = [...body.children];
  let cutFrom = -1;
  for (let i = 0; i < topChildren.length; i++) {
    const txt = topChildren[i].textContent.replace(/\u00a0/g, " ").trim();
    if (txt === "--" || txt === "-- ") { cutFrom = i; break; }
  }
  if (cutFrom !== -1) topChildren.slice(cutFrom).forEach((el) => el.remove());

  return body.innerHTML;
}

/**
 * Converts HTML to clean plain text, preserving paragraph and list structure
 * as line breaks.
 */
function htmlToText(html) {
  if (!html) return "";
  const doc = new DOMParser().parseFromString(html, "text/html");
  const body = doc.body;

  body.querySelectorAll("br").forEach((el) => el.replaceWith("\n"));
  body.querySelectorAll("p, div, h1, h2, h3, h4, h5, h6, tr, blockquote").forEach((el) => {
    el.prepend(doc.createTextNode("\n"));
    el.append(doc.createTextNode("\n"));
  });
  body.querySelectorAll("li").forEach((el) => {
    el.prepend(doc.createTextNode("• "));
    el.append(doc.createTextNode("\n"));
  });

  return (body.textContent || "")
    .replace(/\r\n/g, "\n")
    .replace(/[ \t]{2,}/g, " ")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

/**
 * Converts the AI's Markdown-formatted response to HTML for display in the
 * result panel (class-based styling via taskpane.css).
 *
 * Built-in converter — no external library required.
 * Supports: h1–h6, bold, italic, bold-italic, strikethrough, inline code,
 * links, bullet/ordered lists (with 2-level nesting), blockquotes,
 * horizontal rules, and blank-line-separated paragraphs.
 */
function markdownToHtml(md) {
  if (!md) return "";

  // ── Inline formatter ─────────────────────────────────────────────────────
  // HTML-escapes first, then applies Markdown span patterns so user content
  // can never break out of attribute values or introduce raw tags.
  function inline(text) {
    return text
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      // links before code so backticks inside link text work normally
      .replace(/\[([^\]]+)\]\(([^)]+)\)/g, '<a href="$2">$1</a>')
      .replace(/`([^`\n]+)`/g, "<code>$1</code>")
      .replace(/\*\*\*(.+?)\*\*\*/gs, "<strong><em>$1</em></strong>")
      .replace(/\*\*(.+?)\*\*/gs, "<strong>$1</strong>")
      .replace(/\*(.+?)\*/gs, "<em>$1</em>")
      .replace(/___(.+?)___/gs, "<strong><em>$1</em></strong>")
      .replace(/__(.+?)__/gs, "<strong>$1</strong>")
      .replace(/_([^_\s][^_\n]*)_/gs, "<em>$1</em>")
      .replace(/~~(.+?)~~/gs, "<del>$1</del>");
  }

  // ── Block parser ──────────────────────────────────────────────────────────
  const lines    = md.split("\n");
  const out      = [];
  const listStack = [];   // {type:"ul"|"ol", indent:number}
  let   paraLines = [];   // lines being accumulated into a <p>

  function flushPara() {
    if (!paraLines.length) return;
    out.push(`<p>${paraLines.join("<br>\n")}</p>`);
    paraLines = [];
  }

  // Close every open list whose indent level is ≥ the given threshold.
  function closeListsTo(threshold) {
    while (listStack.length && listStack[listStack.length - 1].indent >= threshold) {
      out.push(`</${listStack.pop().type}>`);
    }
  }

  function closeAllLists() {
    while (listStack.length) out.push(`</${listStack.pop().type}>`);
  }

  for (const rawLine of lines) {
    const trimmed = rawLine.trimStart();
    const indent  = rawLine.length - trimmed.length;

    // ── Blank line ────────────────────────────────────────────────────────
    if (!trimmed) {
      flushPara();
      // Lists stay open across a single blank line (standard Markdown)
      continue;
    }

    // ── ATX Heading: # … ###### ──────────────────────────────────────────
    const hMatch = trimmed.match(/^(#{1,6})\s+(.+?)(?:\s+#+\s*)?$/);
    if (hMatch) {
      flushPara();
      closeAllLists();
      const lvl = hMatch[1].length;
      out.push(`<h${lvl}>${inline(hMatch[2])}</h${lvl}>`);
      continue;
    }

    // ── Horizontal rule: ---, ***, ___ (3 or more identical chars) ───────
    if (/^([-*_])\s*(?:\1\s*){2,}$/.test(trimmed)) {
      flushPara();
      closeAllLists();
      out.push("<hr>");
      continue;
    }

    // ── Blockquote: > text ───────────────────────────────────────────────
    if (trimmed.startsWith("> ") || trimmed === ">") {
      flushPara();
      closeAllLists();
      out.push(`<blockquote><p>${inline(trimmed === ">" ? "" : trimmed.slice(2))}</p></blockquote>`);
      continue;
    }

    // ── Unordered list item: -, *, + ─────────────────────────────────────
    const ulMatch = trimmed.match(/^[-*+]\s+(.*)/);
    if (ulMatch) {
      flushPara();
      closeListsTo(indent + 1);           // close any lists deeper than this
      const top = listStack[listStack.length - 1];
      if (!top || top.indent < indent) {
        out.push("<ul>"); listStack.push({ type: "ul", indent });
      } else if (top.type !== "ul") {
        out.push("</ol>"); listStack.pop();
        out.push("<ul>"); listStack.push({ type: "ul", indent });
      }
      out.push(`<li>${inline(ulMatch[1])}</li>`);
      continue;
    }

    // ── Ordered list item: 1. 2. etc. ────────────────────────────────────
    const olMatch = trimmed.match(/^(\d+)\.\s+(.*)/);
    if (olMatch) {
      flushPara();
      closeListsTo(indent + 1);
      const top = listStack[listStack.length - 1];
      if (!top || top.indent < indent) {
        out.push("<ol>"); listStack.push({ type: "ol", indent });
      } else if (top.type !== "ol") {
        out.push("</ul>"); listStack.pop();
        out.push("<ol>"); listStack.push({ type: "ol", indent });
      }
      out.push(`<li>${inline(olMatch[2])}</li>`);
      continue;
    }

    // ── Regular paragraph text ────────────────────────────────────────────
    // A non-list line at indent 0 closes any open lists.
    if (listStack.length && indent === 0) closeAllLists();
    paraLines.push(inline(rawLine));
  }

  flushPara();
  closeAllLists();
  return out.join("\n");
}

/**
 * Produces Outlook-compatible HTML with inline styles so that bold, italics,
 * bullet points, and paragraph spacing survive insertion into the compose body.
 * Used for both the Insert button and the clipboard copy.
 */
function markdownToHtmlForOutlook(md) {
  if (!md) return "";
  const base = markdownToHtml(md);

  // Post-process: walk the DOM and attach inline styles to every element.
  // Outlook's Word-based engine strips class-based CSS but honours inline styles.
  const doc  = new DOMParser().parseFromString(`<div id="root">${base}</div>`, "text/html");
  const root = doc.getElementById("root");

  root.querySelectorAll("p").forEach((el) => {
    el.style.margin      = "0 0 10px 0";
    el.style.padding     = "0";
    el.style.lineHeight  = "1.6";
  });
  root.querySelectorAll("h1").forEach((el) => {
    el.style.fontSize   = "20px";
    el.style.fontWeight = "bold";
    el.style.margin     = "16px 0 8px 0";
    el.style.lineHeight = "1.3";
  });
  root.querySelectorAll("h2").forEach((el) => {
    el.style.fontSize   = "17px";
    el.style.fontWeight = "bold";
    el.style.margin     = "14px 0 6px 0";
    el.style.lineHeight = "1.3";
  });
  root.querySelectorAll("h3, h4, h5, h6").forEach((el) => {
    el.style.fontSize   = "15px";
    el.style.fontWeight = "bold";
    el.style.margin     = "12px 0 6px 0";
    el.style.lineHeight = "1.3";
  });
  root.querySelectorAll("ul").forEach((el) => {
    el.style.margin         = "8px 0 10px 24px";
    el.style.padding        = "0";
    el.style.listStyleType  = "disc";
  });
  root.querySelectorAll("ol").forEach((el) => {
    el.style.margin         = "8px 0 10px 24px";
    el.style.padding        = "0";
    el.style.listStyleType  = "decimal";
  });
  root.querySelectorAll("li").forEach((el) => {
    el.style.margin     = "4px 0";
    el.style.lineHeight = "1.6";
  });
  root.querySelectorAll("strong, b").forEach((el) => {
    el.style.fontWeight = "bold";
  });
  root.querySelectorAll("em, i").forEach((el) => {
    el.style.fontStyle = "italic";
  });
  root.querySelectorAll("u").forEach((el) => {
    el.style.textDecoration = "underline";
  });
  root.querySelectorAll("blockquote").forEach((el) => {
    el.style.borderLeft  = "3px solid #cccccc";
    el.style.margin      = "10px 0 10px 16px";
    el.style.paddingLeft = "12px";
    el.style.color       = "#555555";
    el.style.fontStyle   = "italic";
  });
  root.querySelectorAll("a").forEach((el) => {
    el.style.color          = "#0563C1";
    el.style.textDecoration = "underline";
  });

  return root.innerHTML;
}


// ── Office.js helpers ─────────────────────────────────────────────────────────

function getSubject() {
  const subj = Office.context.mailbox.item.subject;
  if (typeof subj === "string") return Promise.resolve(subj);
  return new Promise((resolve) => {
    try {
      subj.getAsync((r) => resolve(r.status === Office.AsyncResultStatus.Succeeded ? r.value : ""));
    } catch { resolve(""); }
  });
}

function getFrom() {
  const f = Office.context.mailbox.item.from;
  return f?.displayName || f?.emailAddress || "";
}

async function loadSource() {
  try {
    const body = await getBody();
    const preview = body.length > 600 ? body.slice(0, 600) + "…" : body || "(empty)";
    $("sourcePreview").textContent = preview;
  } catch (e) {
    $("sourcePreview").textContent = e.message;
  }
}

// ── Run ───────────────────────────────────────────────────────────────────────

async function run() {
  hideResult();
  setBusy(true);
  try {
    // For reply: read the full thread (including quoted messages) so the AI
    // has full conversation context. For other actions: stripped body only.
    const body = state.action === "reply"
      ? await getBodyWithThread()
      : await getBody();
    if (!body) throw new Error("There's no text in this email yet.");

    const payload = { body, subject: await getSubject(), from: getFrom() };
    if (state.action === "reply") {
      payload.intent    = $("intent").value.trim();
      payload.hasThread = true;
    }
    if (state.action === "tone") payload.tone = state.tone;

    // Build the settings override for this request.
    // For cloud: use per-provider stored values; fall back to PROVIDER_DEFAULTS
    // so the server always receives a concrete model and base URL.
    const cp  = settings.cloudProvider;
    const def = PROVIDER_DEFAULTS[cp] ?? {};
    const cloudSettingsPayload =
      state.provider === "cloud"
        ? {
            cloudProvider: cp,
            apiKey:        settings[`apiKey_${cp}`]    || undefined,
            cloudModel:    settings[`model_${cp}`]     || def.model    || undefined,
            cloudBaseUrl:  settings[`baseUrl_${cp}`]   || def.baseUrl  || undefined,
          }
        : {
            ollamaModel:   settings.ollamaModel   || undefined,
            ollamaBaseUrl: settings.ollamaBaseUrl || undefined,
          };

    const res = await fetch("/api/run", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        action: state.action,
        payload,
        provider: state.provider,
        cloudSettings: cloudSettingsPayload,
      }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "The model request failed.");

    showResult(data.result, data.ran);
  } catch (e) {
    showError(e.message);
  } finally {
    setBusy(false);
  }
}

// ── Result UI ─────────────────────────────────────────────────────────────────

function setBusy(busy) {
  $("runBtn").disabled = busy;
  $("runText").textContent = busy ? "Working…" : RUN_LABEL[state.action];
}

function setAction(action) {
  state.action = action;
  document.querySelectorAll(".tab").forEach((t) =>
    t.classList.toggle("is-active", t.dataset.action === action)
  );
  $("intentField").hidden  = action !== "reply";
  $("toneField").hidden    = action !== "tone";
  $("runText").textContent = RUN_LABEL[action];
  $("sourceWrap").hidden   = action === "tone";
  // Update the Insert/Replace button label
  $("insertBtn").textContent = INSERT_LABEL[action] || "Insert";
  hideResult();
}

function showResult(text, ran) {
  state.lastResult = text;
  $("resultTag").textContent =
    state.action === "summarize" ? "Summary"
    : state.action === "reply"  ? "Draft reply"
    : "Rewritten";

  // Render the Markdown as formatted HTML in the result panel
  $("resultBody").innerHTML = markdownToHtml(text);

  // "Insert" in compose; "Replace draft" for tone; hidden in read mode
  $("insertBtn").hidden      = !state.isCompose;
  $("insertBtn").textContent = INSERT_LABEL[state.action] || "Insert";
  $("result").hidden         = false;

  if (ran) showNote(`Ran on ${ran.provider} · ${ran.model}`, false);
}

function hideResult() {
  $("result").hidden = true;
  $("note").hidden   = true;
}

function showError(msg) { showNote(msg, true); }

function showNote(msg, isError) {
  const n = $("note");
  n.textContent = msg;
  n.classList.toggle("is-error", isError);
  n.hidden = false;
}

// ── Clipboard / insert ────────────────────────────────────────────────────────

async function copyResult() {
  try {
    // Write both HTML and plain text so pasting into Outlook or Word gives rich
    // text, while pasting into a plain-text field still works.
    const html = markdownToHtmlForOutlook(state.lastResult);
    const clip = new ClipboardItem({
      "text/html":  new Blob([html],              { type: "text/html" }),
      "text/plain": new Blob([state.lastResult],  { type: "text/plain" }),
    });
    await navigator.clipboard.write([clip]);
    flash($("copyBtn"), "Copied");
  } catch {
    // ClipboardItem not supported in this context — fall back to plain text
    try {
      await navigator.clipboard.writeText(state.lastResult);
      flash($("copyBtn"), "Copied");
    } catch {
      showError("Couldn't copy to the clipboard.");
    }
  }
}

function insertResult() {
  // Always use Outlook-compatible HTML (inline styles) when writing to the body.
  const html = markdownToHtmlForOutlook(state.lastResult);
  const body = Office.context.mailbox.item.body;

  if ((state.action === "tone" || state.action === "reply") && typeof body.setAsync === "function") {
    // Reply: set the full compose body so the AI draft is ready to send.
    // Tone: replace the entire draft with the rewritten version.
    const successLabel = state.action === "reply" ? "Draft set" : "Replaced";
    body.setAsync(
      html,
      { coercionType: Office.CoercionType.Html },
      (r) => {
        if (r.status === Office.AsyncResultStatus.Succeeded) flash($("insertBtn"), successLabel);
        else showError(r.error?.message || "Couldn't set the email body.");
      }
    );
  } else {
    // Summarize: insert formatted HTML at the current cursor position.
    body.setSelectedDataAsync(
      html,
      { coercionType: Office.CoercionType.Html },
      (r) => {
        if (r.status === Office.AsyncResultStatus.Succeeded) flash($("insertBtn"), "Inserted");
        else showError(r.error?.message || "Couldn't insert into the email.");
      }
    );
  }
}

function flash(btn, label) {
  const prev = btn.textContent;
  btn.textContent = label;
  setTimeout(() => (btn.textContent = prev), 1200);
}
