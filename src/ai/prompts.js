// prompts.js
// Builds provider-agnostic chat messages for each action.
// Used by both the local server and the browser demo, so it stays free of
// any Node- or browser-specific code.

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

// Per-tone style guidance injected into the rewrite system prompt.
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

// Shared formatting contract sent to the model so it always produces
// Markdown that our converter can render cleanly in Outlook.
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
  if (from)    lines.push(`From: ${from}`);
  lines.push("", body || "");
  return lines.join("\n");
}

/**
 * @param {"summarize"|"reply"|"tone"} action
 * @param {{
 *   subject?:string, from?:string, body:string,
 *   intent?:string, tone?:string, hasThread?:boolean
 * }} payload
 * @returns {{role:string, content:string}[]}
 */
export function buildMessages(action, payload) {
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
          ].filter(Boolean).join("\n"),
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

export { TONES };
