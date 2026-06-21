// demo.js — fully self-contained (no external imports).
// Two modes:
//   Demo — returns realistic pre-written results instantly (no key, no setup).
//   Live — calls a real provider directly from the browser with a key the visitor pastes.

// ---------------------------------------------------------------------------
// Prompts (inlined from src/ai/prompts.js)
// ---------------------------------------------------------------------------

const TONES = [
  "Professional",
  "Friendly",
  "Concise",
  "Assertive",
  "Apologetic",
  "Government",
  "Formal",
  "Diplomatic",
  "Direct",
  "Empathetic",
  "Casual",
];

const TONE_GUIDES = {
  Professional:
    "Write in a polished, business-appropriate style. Use clear, respectful language. " +
    "Avoid slang. Greet with 'Hi [Name],' or 'Dear [Name],' and close with 'Best regards,' or 'Kind regards,'.",
  Friendly:
    "Write in a warm, approachable style. Be personable and upbeat. Contractions are welcome. " +
    "Avoid overly stiff phrasing. Greet with 'Hi [Name],' and close with 'Thanks,' or 'Best,'.",
  Concise:
    "Be extremely concise — strip the message to its essential meaning. Use short sentences and bullet points " +
    "for multiple items. Eliminate pleasantries, filler words, and repetition. Every word must earn its place.",
  Assertive:
    "Write with confidence and authority. State your position clearly without hedging or over-apologising. " +
    "Use active voice. Be direct about expectations and timelines. Close with 'Regards,'.",
  Apologetic:
    "Open by sincerely acknowledging any inconvenience, error, or concern. Express genuine empathy. " +
    "Then offer a constructive solution or a clear path forward. Avoid defensive language. " +
    "Close with 'With sincere apologies,' or 'Sincerely,'.",
  Government:
    "Write in a formal, official style appropriate for government or institutional correspondence. " +
    "Use no contractions. Use passive voice where it sounds natural for official language. " +
    "Reference processes, procedures, or policy when relevant. Use honorifics and proper titles. " +
    "Greet with 'Dear Sir/Madam,' or 'Dear [Title Last Name],'. " +
    "Close with 'Yours faithfully,' when the recipient is unknown, or 'Yours sincerely,' when named.",
  Formal:
    "Use elevated, precise vocabulary and traditional letter structure. No contractions or colloquial expressions. " +
    "Sentences should be well-constructed and complete. " +
    "Greet with 'Dear [Name],' or 'Dear Sir/Madam,'. Close with 'Yours sincerely,' or 'Yours truly,'.",
  Diplomatic:
    "Use tactful, balanced, and non-confrontational language. Acknowledge the other party's perspective " +
    "before presenting your own. Find common ground. Soften disagreements with phrases like " +
    "'I appreciate your perspective…' or 'While I understand…'. Avoid blame or accusatory language. " +
    "Close with 'With kind regards,'.",
  Direct:
    "Skip pleasantries entirely. Lead immediately with the core ask or key point. " +
    "Use short, declarative sentences. State what is needed and why in as few words as possible. " +
    "Omit filler. A minimal or no greeting/closing is appropriate.",
  Empathetic:
    "Begin by genuinely acknowledging the recipient's situation, feelings, or effort before addressing the matter. " +
    "Use warm, supportive language throughout. Show that you are listening. " +
    "Greet with 'Hi [Name],' and close with 'With warm regards,' or 'Take care,'.",
  Casual:
    "Write as you would speak to a familiar colleague. Contractions are encouraged. " +
    "Keep it light, human, and easy to read. Greet with 'Hey [Name],' or 'Hi [Name],' " +
    "and close with 'Cheers,' or 'Thanks,'.",
};

const FORMAT_INSTRUCTIONS = `
Format your response using Markdown:
- Separate paragraphs with a blank line between them.
- Use **bold** for key terms, names, dates, decisions, or action items (use sparingly).
- Use *italic* for emphasis or titles.
- Use bullet lists (- item) for action items, options, or enumerated points.
- Use numbered lists (1. item) only when order genuinely matters.
- Use > blockquote for any quoted or referenced text.
- Use ### for section headings only when the content clearly has multiple sections.
- Do not add a preamble like "Here is the draft:" — start directly with the greeting or content.
- Do not add a sign-off or closing unless explicitly writing a reply or email.
`.trim();

function emailContextBlock({ subject, from, body }) {
  const lines = [];
  if (subject) lines.push(`Subject: ${subject}`);
  if (from) lines.push(`From: ${from}`);
  lines.push("", body || "");
  return lines.join("\n");
}

function buildMessages(action, payload) {
  switch (action) {
    case "summarize":
      return [
        {
          role: "system",
          content: [
            "You summarize email threads for a busy professional.",
            "Write a concise summary (2–4 sentences) of what the thread is about,",
            "followed by a bullet list of action items, decisions, or open questions.",
            "If there are no action items or decisions, say so briefly.",
            "Never invent details that are not explicitly stated in the email.",
            "",
            FORMAT_INSTRUCTIONS,
          ].join("\n"),
        },
        { role: "user", content: emailContextBlock(payload) },
      ];

    case "reply": {
      const intent = (payload.intent || "").trim();
      const hasThread = Boolean(payload.hasThread);
      const instruction = intent
        ? `Write the reply around this intent: ${intent}`
        : "Write an appropriate, helpful reply.";
      const threadInstruction = hasThread
        ? "The message below contains the full email thread. Read all messages carefully to understand the full conversation context before composing your reply."
        : "";
      return [
        {
          role: "system",
          content: [
            "You draft email replies that are natural, professional, and ready to send.",
            threadInstruction,
            "Output only the reply body — no Subject line, no To/From headers.",
            "Start with a greeting appropriate to the conversation's register",
            "(e.g. 'Hi [Name],' for informal exchanges, 'Dear [Name],' for formal ones).",
            "End with a polite closing appropriate to the tone, followed by a blank line.",
            "Do not add a placeholder like [Your Name] in the sign-off.",
            "Keep the reply focused — do not restate what the original message already said.",
            "",
            FORMAT_INSTRUCTIONS,
          ]
            .filter(Boolean)
            .join("\n"),
        },
        {
          role: "user",
          content: `${instruction}\n\n--- Email thread ---\n${emailContextBlock(payload)}`,
        },
      ];
    }

    case "tone": {
      const tone = TONES.includes(payload.tone) ? payload.tone : "Professional";
      const guide = TONE_GUIDES[tone] || TONE_GUIDES.Professional;
      return [
        {
          role: "system",
          content: [
            `You rewrite email drafts in a specific tone. The requested tone is: **${tone}**.`,
            "",
            `Tone guidance: ${guide}`,
            "",
            "Rules:",
            "- Preserve the meaning, facts, names, numbers, commitments, and intent exactly.",
            "- Keep the same language as the original draft.",
            "- Adjust the greeting and closing to match the tone.",
            "- Output only the rewritten email body — no explanation, no commentary, no subject line.",
            "",
            FORMAT_INSTRUCTIONS,
          ].join("\n"),
        },
        {
          role: "user",
          content: `Rewrite the following email draft in a **${tone}** tone:\n\n${payload.body || ""}`,
        },
      ];
    }

    default:
      throw new Error(`Unknown action: ${action}`);
  }
}

// ---------------------------------------------------------------------------
// Providers (inlined from src/ai/providers.js)
// ---------------------------------------------------------------------------

const PROVIDER_DEFAULTS = {
  openai:    { baseUrl: "https://api.openai.com/v1",                               model: "gpt-5.5" },
  anthropic: { baseUrl: "https://api.anthropic.com",                               model: "claude-sonnet-4-6" },
  google:    { baseUrl: "https://generativelanguage.googleapis.com/v1beta/openai", model: "gemini-2.5-flash" },
};

async function runChat(config, messages) {
  const provider = config.provider;
  const d = PROVIDER_DEFAULTS[provider];
  if (!d) throw new Error(`Unsupported provider: ${provider}`);

  const model = config.model || d.model;
  const baseUrl = (config.baseUrl || d.baseUrl).replace(/\/$/, "");
  const maxTokens = config.maxTokens || 1024;

  if (provider === "openai" || provider === "google") {
    const res = await fetch(`${baseUrl}/chat/completions`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${config.apiKey}`,
      },
      body: JSON.stringify({ model, messages, max_tokens: maxTokens }),
    });
    if (!res.ok) {
      if (provider === "openai") {
        throw new Error(
          "OpenAI usually blocks direct browser calls (CORS). Try Anthropic or Google here, or run the add-in locally."
        );
      }
      throw new Error(`API error ${res.status}: ${await safeText(res)}`);
    }
    const data = await res.json();
    return data?.choices?.[0]?.message?.content?.trim() || "";
  }

  if (provider === "anthropic") {
    const system = messages
      .filter((m) => m.role === "system")
      .map((m) => m.content)
      .join("\n\n");
    const turns = messages.filter((m) => m.role !== "system");
    const res = await fetch(`${baseUrl}/v1/messages`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": config.apiKey,
        "anthropic-version": "2023-06-01",
        "anthropic-dangerous-direct-browser-access": "true",
      },
      body: JSON.stringify({ model, max_tokens: maxTokens, system, messages: turns }),
    });
    if (!res.ok) throw new Error(`Anthropic error ${res.status}: ${await safeText(res)}`);
    const data = await res.json();
    return data?.content?.map((b) => b.text || "").join("").trim() || "";
  }

  throw new Error(`Unsupported provider: ${provider}`);
}

async function safeText(res) {
  try {
    return (await res.text()).slice(0, 300);
  } catch {
    return "(no body)";
  }
}

// ---------------------------------------------------------------------------
// Canned demo results (matched to the sample email)
// ---------------------------------------------------------------------------

const CANNED = {
  summarize:
    "Priya has reviewed the proposal and is close to approving, but attaches three " +
    "conditions and one request:\n\n" +
    "- **SSO with Azure AD** must ship at launch, not in a later phase.\n" +
    "- **Role-based access** is mandatory — finance must not see HR records.\n" +
    "- She needs confirmation that the **March 15 go-live** still holds if SSO is pulled forward.\n" +
    "- **Action this week:** send a short data-flow note for her security team.",

  reply:
    "Hi Priya,\n\n" +
    "Thanks for the quick review — happy to confirm on each point:\n\n" +
    "- **SSO with Azure AD** will be in scope for day-one launch, not a later phase.\n" +
    "- **Role-based access** will be enforced, with finance and HR records fully separated.\n" +
    "- The **March 15 go-live** still holds with SSO moved up; it's accounted for in the plan.\n\n" +
    "I'll send the data-flow note for your security team by **Friday**.\n\n" +
    "Best regards,",

  tone: {
    Professional:
      "Hi Priya,\n\n" +
      "I wanted to follow up on your questions regarding the onboarding portal. Could you confirm " +
      "whether we are still on track for the March 15 launch, and whether the SSO requirement has been " +
      "resolved on your end? I am also following up on the data-flow note for your security team.\n\n" +
      "Thank you for your time.\n\nBest regards,",

    Friendly:
      "Hi Priya!\n\n" +
      "Just checking in — are we still good for the March 15 launch, and is the SSO piece all sorted? " +
      "Also still keen to get that data-flow note over to your security team whenever it's ready. " +
      "Let me know!\n\nThanks so much,",

    Concise:
      "Priya — still on for March 15? SSO resolved? Also need the data-flow note for security. Thanks.",

    Assertive:
      "Priya,\n\n" +
      "I need confirmation that March 15 still holds and that SSO is resolved on your end. " +
      "Please also send the data-flow note to your security team today so we can keep the timeline.\n\n" +
      "Regards,",

    Apologetic:
      "Hi Priya,\n\n" +
      "I'm sorry to chase on this — I just wanted to check whether we're still on track for March 15 " +
      "and whether the SSO requirement has been resolved. Apologies for the follow-up, but I also still " +
      "need the data-flow note for your security team when you have a moment.\n\n" +
      "Thank you for your patience.\n\nSincerely,",

    Government:
      "Dear Ms. Nair,\n\n" +
      "I am writing to request confirmation that the March 15 go-live date remains agreed upon, " +
      "and that the SSO integration with Azure Active Directory has been resolved in accordance with " +
      "the requirements stated. It is further requested that the data-flow documentation be provided " +
      "to your security team at the earliest opportunity, in line with the agreed process.\n\n" +
      "Yours sincerely,",

    Formal:
      "Dear Ms. Nair,\n\n" +
      "I am writing to seek your confirmation on two outstanding matters. First, please advise " +
      "whether the March 15 go-live date remains in force given the revised SSO scope. Second, " +
      "I would be grateful if you could arrange for the data-flow note to be forwarded to your " +
      "security team at your earliest convenience.\n\n" +
      "Yours sincerely,",

    Diplomatic:
      "Hi Priya,\n\n" +
      "I appreciate you taking the time to review the proposal in such detail — your points on SSO " +
      "and role-based access are well taken and have been noted carefully. I wanted to gently " +
      "touch base on whether the March 15 timeline still feels workable from your side, and to " +
      "follow up on the data-flow note for your security team whenever that is convenient.\n\n" +
      "With kind regards,",

    Direct:
      "March 15 — still on? SSO resolved? Send the data-flow note.",

    Empathetic:
      "Hi Priya,\n\n" +
      "I know you're juggling a lot with these requirements, and I really appreciate the thorough " +
      "review. I just wanted to check in — does March 15 still feel manageable with the SSO work " +
      "moved forward? And whenever you get a chance, the data-flow note for your security team " +
      "would be a great help.\n\n" +
      "With warm regards,",

    Casual:
      "Hey Priya,\n\n" +
      "Quick check-in — are we still good for March 15 now that SSO is in scope from day one? " +
      "Also, don't forget the data-flow note for your security folks whenever you get a chance!\n\n" +
      "Cheers,",
  },
};

// ---------------------------------------------------------------------------
// UI
// ---------------------------------------------------------------------------

const $ = (id) => document.getElementById(id);
const state = { action: "summarize", tone: "Professional", mode: "demo" };

init();

function init() {
  document.querySelectorAll(".tab").forEach((t) =>
    t.addEventListener("click", () => setAction(t.dataset.action))
  );

  const tonesEl = $("tones");
  TONES.forEach((t, i) => {
    const b = document.createElement("button");
    b.className = "tone" + (i === 0 ? " is-on" : "");
    b.type = "button";
    b.textContent = t;
    b.addEventListener("click", () => {
      state.tone = t;
      document.querySelectorAll(".tone").forEach((x) => x.classList.toggle("is-on", x === b));
    });
    tonesEl.appendChild(b);
  });

  $("runBtn").addEventListener("click", run);
  $("copyBtn").addEventListener("click", copyResult);
  $("runtimeChip").addEventListener("click", () => {
    setMode(state.mode === "demo" ? "live" : "demo");
  });

  document.querySelectorAll(".seg__opt").forEach((o) =>
    o.addEventListener("click", () => setMode(o.dataset.mode))
  );

  $("liveProvider").addEventListener("change", (e) => {
    const provider = e.target.value;
    $("liveBase").hidden = provider !== "custom";
    const placeholders = {
      anthropic: `Model (default: ${PROVIDER_DEFAULTS.anthropic.model})`,
      openai:    `Model (default: ${PROVIDER_DEFAULTS.openai.model})`,
      google:    `Model (default: ${PROVIDER_DEFAULTS.google.model})`,
      custom:    "Model name",
    };
    $("liveModel").placeholder = placeholders[provider] || "Model (optional)";
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
    action === "summarize" ? "Summarize thread" : action === "reply" ? "Draft reply" : "Rewrite tone";
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
      await wait(400);
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
  if (state.action === "tone") return CANNED.tone[state.tone] || CANNED.tone.Professional;
  return CANNED[state.action];
}

async function liveResult() {
  const key = $("liveKey").value.trim();
  if (!key) throw new Error("Paste an API key, or switch back to Demo mode.");

  const provider = $("liveProvider").value;
  const customBase = $("liveBase").value.trim();

  const payload = {
    body: $("emailBody").value,
    subject: $("subjLine").textContent,
    from: $("fromLine").textContent,
  };
  if (state.action === "reply") payload.intent = $("intent").value.trim();
  if (state.action === "tone") payload.tone = state.tone;

  const messages = buildMessages(state.action, payload);

  const resolvedProvider = provider === "google" ? "google" : provider === "custom" ? "openai" : provider;

  return await runChat(
    {
      provider: resolvedProvider,
      apiKey: key,
      model: $("liveModel").value.trim() || undefined,
      baseUrl: customBase || undefined,
      browser: true,
    },
    messages
  );
}

function setBusy(busy) {
  $("runBtn").disabled = busy;
  $("runText").textContent = busy
    ? "Working…"
    : state.action === "summarize"
    ? "Summarize thread"
    : state.action === "reply"
    ? "Draft reply"
    : "Rewrite tone";
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
  const text = $("resultBody").textContent;
  try {
    await navigator.clipboard.writeText(text);
    const b = $("copyBtn");
    b.textContent = "Copied";
    setTimeout(() => (b.textContent = "Copy"), 1200);
  } catch {
    showError("Couldn't copy — try selecting and copying manually.");
  }
}

const wait = (ms) => new Promise((r) => setTimeout(r, ms));
