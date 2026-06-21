// providers.js
// One runChat() entry point over three backends:
//   - ollama    (local, default — nothing leaves the machine)
//   - openai    (OpenAI or any OpenAI-compatible endpoint via baseUrl)
//   - anthropic (Claude)
// Uses the global fetch available in Node 18+ and in the browser, so the same
// module works server-side and in the demo.

// Server-side fallback defaults (used only when the client sends no model/baseUrl).
// Model IDs verified from official docs, June 2026.
const DEFAULTS = {
  ollama:    { baseUrl: "http://localhost:11434",                                      model: "llama3.2" },
  openai:    { baseUrl: "https://api.openai.com/v1",                                  model: "gpt-5.5" },
  anthropic: { baseUrl: "https://api.anthropic.com",                                  model: "claude-sonnet-4-6" },
  // Google Gemini uses the OpenAI-compatible endpoint.
  // The server maps UI provider "google" → serverProvider "openai" + this base URL.
  google:    { baseUrl: "https://generativelanguage.googleapis.com/v1beta/openai",    model: "gemini-2.5-flash" },
};

/**
 * @param {{provider:string, model?:string, baseUrl?:string, apiKey?:string, maxTokens?:number, browser?:boolean}} config
 * @param {{role:string, content:string}[]} messages
 * @returns {Promise<string>}
 */
export async function runChat(config, messages) {
  const provider = config.provider;
  const d = DEFAULTS[provider];
  if (!d) throw new Error(`Unsupported provider: ${provider}`);

  const model = config.model || d.model;
  const baseUrl = (config.baseUrl || d.baseUrl).replace(/\/$/, "");
  const maxTokens = config.maxTokens || 1024;

  if (provider === "ollama") {
    let res;
    try {
      res = await fetch(`${baseUrl}/api/chat`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ model, messages, stream: false }),
      });
    } catch {
      throw new Error(
        `Couldn't reach a local model at ${baseUrl}. Is Ollama running? ` +
          `Start it and run "ollama pull ${model}", or switch to Cloud.`
      );
    }
    if (!res.ok) throw new Error(`Ollama error ${res.status}: ${await safeText(res)}`);
    const data = await res.json();
    return data?.message?.content?.trim() || "";
  }

  if (provider === "openai") {
    const res = await fetch(`${baseUrl}/chat/completions`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${config.apiKey}`,
      },
      body: JSON.stringify({ model, messages, max_tokens: maxTokens }),
    });
    if (!res.ok) throw new Error(`OpenAI error ${res.status}: ${await safeText(res)}`);
    const data = await res.json();
    return data?.choices?.[0]?.message?.content?.trim() || "";
  }

  if (provider === "anthropic") {
    // Anthropic takes the system prompt as a top-level field, not a message.
    const system = messages.filter((m) => m.role === "system").map((m) => m.content).join("\n\n");
    const turns = messages.filter((m) => m.role !== "system");
    const headers = {
      "Content-Type": "application/json",
      "x-api-key": config.apiKey,
      "anthropic-version": "2023-06-01",
    };
    // Allow direct browser calls only in the demo, where the user supplies their own key.
    if (config.browser) headers["anthropic-dangerous-direct-browser-access"] = "true";
    const res = await fetch(`${baseUrl}/v1/messages`, {
      method: "POST",
      headers,
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

export { DEFAULTS };
