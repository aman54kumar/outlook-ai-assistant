// demo.js
// Two modes:
//   Demo  — returns realistic pre-written results instantly (no key, no setup).
//   Live  — calls a real provider with a key the visitor pastes (kept in this tab only).
// The Live path reuses the exact same prompt + provider modules the add-in uses.

import { buildMessages } from "../src/ai/prompts.js";
import { runChat } from "../src/ai/providers.js";

const $ = (id) => document.getElementById(id);
const state = { action: "summarize", tone: "Professional", mode: "demo" };

const TONES = ["Professional", "Friendly", "Concise", "Assertive", "Apologetic"];

// --- Canned results, matched to the sample email ----------------------------
const CANNED = {
  summarize:
    "Priya has reviewed the proposal and is close to approving, but attaches three " +
    "conditions and one request:\n\n" +
    "• SSO with Azure AD must ship at launch, not in a later phase.\n" +
    "• Role-based access is mandatory — finance must not see HR records.\n" +
    "• She needs confirmation that March 15 go-live still holds if SSO is pulled forward.\n" +
    "• Action this week: send a short data-flow note for her security team.",
  reply:
    "Hi Priya,\n\n" +
    "Thanks for the quick review — happy to confirm on each point:\n\n" +
    "- SSO with Azure AD will be in scope for day-one launch, not a later phase.\n" +
    "- Role-based access will be enforced, with finance and HR records fully separated.\n" +
    "- March 15 go-live still holds with SSO moved up; it's accounted for in the plan.\n\n" +
    "I'll send the data-flow note for your security team by Friday.\n\n" +
    "Best,\n[Your name]",
  tone: {
    Professional:
      "Hi Priya, could you confirm whether we're still on track for the March 15 launch " +
      "and whether the SSO work is resolved? I'm also following up on the data-flow note. Thank you.",
    Friendly:
      "Hi Priya! Just checking in — are we still good for the March 15 launch, and is the " +
      "SSO piece sorted? Also still keen to get that data-flow note whenever it's ready. Thanks so much!",
    Concise: "Still on for March 15? Is SSO resolved? Also need the data-flow note. Thanks.",
    Assertive:
      "Priya — I need confirmation that March 15 still holds and that SSO is resolved. " +
      "Please also send the data-flow note today.",
    Apologetic:
      "Hi Priya, sorry to chase — could you let me know if we're still on track for March 15 " +
      "and whether SSO is sorted? Apologies again, but I also still need that data-flow note. " +
      "Thanks for your patience.",
  },
};

init();

function init() {
  document.querySelectorAll(".tab").forEach((t) =>
    t.addEventListener("click", () => setAction(t.dataset.action))
  );

  const tones = $("tones");
  TONES.forEach((t, i) => {
    const b = document.createElement("button");
    b.className = "tone" + (i === 0 ? " is-on" : "");
    b.type = "button";
    b.textContent = t;
    b.addEventListener("click", () => {
      state.tone = t;
      document.querySelectorAll(".tone").forEach((x) => x.classList.toggle("is-on", x === b));
    });
    tones.appendChild(b);
  });

  $("runBtn").addEventListener("click", run);
  $("copyBtn").addEventListener("click", copyResult);
  $("runtimeChip").addEventListener("click", () => {
    // toggle quickly between demo/live from the chip
    setMode(state.mode === "demo" ? "live" : "demo");
  });

  document.querySelectorAll(".seg__opt").forEach((o) =>
    o.addEventListener("click", () => setMode(o.dataset.mode))
  );
  $("liveProvider").addEventListener("change", (e) => {
    $("liveBase").hidden = e.target.value !== "openai";
  });

  setAction("summarize");
  setMode("demo");
}

function setAction(action) {
  state.action = action;
  document.querySelectorAll(".tab").forEach((t) =>
    t.classList.toggle("is-active", t.dataset.action === action)
  );
  $("intentField").hidden = action !== "reply";
  $("toneField").hidden = action !== "tone";
  $("runText").textContent =
    action === "summarize" ? "Summarize thread" : action === "reply" ? "Draft reply" : "Rewrite text";
  hideResult();
}

function setMode(mode) {
  state.mode = mode;
  document.querySelectorAll(".seg__opt").forEach((o) =>
    o.classList.toggle("is-on", o.dataset.mode === mode)
  );
  const live = mode === "live";
  $("liveCfg").classList.toggle("is-open", live);
  $("runtimeChip").classList.toggle("is-live", live);
  $("runtimeLabel").textContent = live ? "Live" : "Demo";
  $("drawerHint").textContent = live
    ? "Live mode calls your provider directly from this tab. Your key is never stored or sent anywhere else."
    : "Demo mode returns realistic, pre-written results — no setup, no key.";
}

async function run() {
  hideResult();
  setBusy(true);
  try {
    if (state.mode === "demo") {
      await wait(450);
      showResult(demoResult());
    } else {
      const text = await liveResult();
      showResult(text);
    }
  } catch (e) {
    showError(e.message);
  } finally {
    setBusy(false);
  }
}

function demoResult() {
  if (state.action === "tone") return CANNED.tone[state.tone];
  return CANNED[state.action];
}

async function liveResult() {
  const key = $("liveKey").value.trim();
  if (!key) throw new Error("Add an API key, or switch back to Demo mode.");
  const provider = $("liveProvider").value;

  const payload = {
    body: $("emailBody").value,
    subject: $("subjLine").textContent,
    from: $("fromLine").textContent,
  };
  if (state.action === "reply") payload.intent = $("intent").value.trim();
  if (state.action === "tone") payload.tone = state.tone;

  const messages = buildMessages(state.action, payload);
  try {
    return await runChat(
      {
        provider,
        apiKey: key,
        model: $("liveModel").value.trim() || undefined,
        baseUrl: provider === "openai" ? $("liveBase").value.trim() || undefined : undefined,
        browser: true,
      },
      messages
    );
  } catch (e) {
    if (provider === "openai") {
      throw new Error(
        "OpenAI usually blocks direct browser calls (CORS). Try Anthropic here, or run the add-in locally."
      );
    }
    throw e;
  }
}

function setBusy(busy) {
  $("runBtn").disabled = busy;
  $("runText").textContent = busy
    ? "Working…"
    : state.action === "summarize"
    ? "Summarize thread"
    : state.action === "reply"
    ? "Draft reply"
    : "Rewrite text";
}

function showResult(text) {
  $("resultBody").textContent = text;
  $("resultTag").textContent =
    state.action === "summarize"
      ? "Summary"
      : state.action === "reply"
      ? "Draft reply"
      : `Rewritten · ${state.tone}`;
  $("result").hidden = false;
}

function hideResult() {
  $("result").hidden = true;
  $("note").hidden = true;
}

function showError(msg) {
  const n = $("note");
  n.textContent = msg;
  n.classList.add("is-error");
  n.hidden = false;
}

async function copyResult() {
  try {
    await navigator.clipboard.writeText($("resultBody").textContent);
    const b = $("copyBtn");
    b.textContent = "Copied";
    setTimeout(() => (b.textContent = "Copy"), 1200);
  } catch {
    showError("Couldn't copy to the clipboard.");
  }
}

const wait = (ms) => new Promise((r) => setTimeout(r, ms));
