// server.js
// A small local server that does two jobs:
//   1. Serves the add-in's pages over HTTPS (Office requires HTTPS).
//   2. Runs the AI calls server-side, so a cloud API key never touches the
//      task pane / browser. By default it talks to a local Ollama model, so
//      nothing leaves the machine at all.
//
// Cloud settings can come from two places (client overrides server .env):
//   a. The request body (cloudSettings), set via the task pane settings drawer.
//   b. The .env file on the server (OPENAI_API_KEY, ANTHROPIC_API_KEY, etc.).

import express from "express";
import https from "node:https";
import http from "node:http";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { buildMessages } from "../src/ai/prompts.js";
import { runChat } from "../src/ai/providers.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = path.join(__dirname, "..");
const PORT = process.env.PORT || 3000;

// Load .env if present (no dotenv dependency).
try {
  const envFile = fs.readFileSync(path.join(root, ".env"), "utf8");
  for (const line of envFile.split("\n")) {
    const m = line.match(/^\s*([A-Z0-9_]+)\s*=\s*(.*?)\s*$/);
    if (m && !process.env[m[1]]) process.env[m[1]] = m[2].replace(/^["']|["']$/g, "");
  }
} catch {
  // no .env — fully local is fine
}

// ── Provider config ───────────────────────────────────────────────────────────
// `override` comes from the task pane's cloudSettings and takes precedence
// over .env values so users can configure everything from the UI.

function resolveConfig(requested, override = {}) {
  if (requested === "cloud") {
    // UI providers: "openai" | "anthropic" | "google" | "custom"
    // "google" and "custom" both use the OpenAI-compatible code path.
    const uiProvider  = override.cloudProvider || process.env.CLOUD_PROVIDER || "openai";
    const serverProvider = uiProvider === "anthropic" ? "anthropic" : "openai";

    // API key: client override wins, then fall back to .env.
    const apiKey =
      override.apiKey ||
      (serverProvider === "anthropic"
        ? process.env.ANTHROPIC_API_KEY
        : process.env.OPENAI_API_KEY) ||
      "";

    // Base URL: client override, then Google's known endpoint, then .env.
    let baseUrl = override.cloudBaseUrl || "";
    if (!baseUrl && uiProvider === "google") {
      baseUrl = "https://generativelanguage.googleapis.com/v1beta/openai";
    }
    if (!baseUrl) baseUrl = process.env.CLOUD_BASE_URL || "";

    return {
      provider: serverProvider,
      model:    override.cloudModel || process.env.CLOUD_MODEL || undefined,
      baseUrl:  baseUrl || undefined,
      apiKey:   apiKey  || undefined,
    };
  }

  // Local / Ollama
  return {
    provider: "ollama",
    model:   override.ollamaModel   || process.env.OLLAMA_MODEL   || undefined,
    baseUrl: override.ollamaBaseUrl || process.env.OLLAMA_BASE_URL || undefined,
  };
}

// ── Express app ───────────────────────────────────────────────────────────────

const app = express();
app.use(express.json({ limit: "1mb" }));

// Serve the add-in pages at the URLs the manifest points to.
app.use(express.static(path.join(root, "src", "taskpane")));
app.use(express.static(path.join(root, "src", "commands")));
app.use("/assets", express.static(path.join(root, "assets")));
app.use("/ai", express.static(path.join(root, "src", "ai")));

// Serve Office.js locally so the add-in works without internet access.
// The npm package (@microsoft/office-js) ships the same dist files that the
// Microsoft CDN hosts; office.js is the umbrella loader for all platforms.
app.use("/vendor/office-js", express.static(
  path.join(root, "node_modules", "@microsoft", "office-js", "dist")
));

// Health check — tells the UI whether server-side keys are already set,
// so the drawer can show a sensible hint.
app.get("/api/health", (_req, res) => {
  const hasOpenAI    = Boolean(process.env.OPENAI_API_KEY);
  const hasAnthropic = Boolean(process.env.ANTHROPIC_API_KEY);
  res.json({
    ok: true,
    localModel: process.env.OLLAMA_MODEL || "llama3.2",
    cloudConfigured: hasOpenAI || hasAnthropic,
    serverProviders: { openai: hasOpenAI, anthropic: hasAnthropic },
    cloudProvider: process.env.CLOUD_PROVIDER || "openai",
  });
});

// Fetches the live model list from a provider and returns chat-capable IDs.
app.post("/api/models", async (req, res) => {
  try {
    const { cloudProvider, apiKey, cloudBaseUrl } = req.body || {};
    if (!apiKey) return res.status(400).json({ error: "API key required to fetch models." });

    const uiProvider    = cloudProvider || "openai";
    const isAnthropic   = uiProvider === "anthropic";
    const serverProvider = isAnthropic ? "anthropic" : "openai";

    let baseUrl = cloudBaseUrl || "";
    if (!baseUrl && uiProvider === "google") {
      baseUrl = "https://generativelanguage.googleapis.com/v1beta/openai";
    }
    if (!baseUrl) baseUrl = DEFAULTS[serverProvider].baseUrl;
    baseUrl = baseUrl.replace(/\/$/, "");

    let models = [];

    if (isAnthropic) {
      const resp = await fetch(`${baseUrl}/v1/models`, {
        headers: {
          "x-api-key":          apiKey,
          "anthropic-version":  "2023-06-01",
        },
      });
      if (!resp.ok) throw new Error(`Anthropic ${resp.status}: ${await safeText(resp)}`);
      const data = await resp.json();
      models = (data.data || []).map((m) => m.id).filter(Boolean);
    } else {
      // OpenAI-compatible endpoint (OpenAI, Google Gemini, custom)
      const resp = await fetch(`${baseUrl}/models`, {
        headers: { Authorization: `Bearer ${apiKey}` },
      });
      if (!resp.ok) throw new Error(`${resp.status}: ${await safeText(resp)}`);
      const data = await resp.json();
      models = (data.data || []).map((m) => m.id).filter(Boolean);
    }

    // Filter to chat-capable models only and sort.
    models = filterChatModels(models, uiProvider);
    res.json({ models });
  } catch (err) {
    console.error("[api/models]", err.message);
    res.status(500).json({ error: err.message });
  }
});

/** Keeps only models that can be used for chat completions. */
function filterChatModels(ids, uiProvider) {
  let filtered = ids;

  if (uiProvider === "anthropic") {
    // Anthropic only returns text models, no filtering needed.
    filtered = ids;
  } else if (uiProvider === "google") {
    // Keep only gemini models from the OpenAI-compat endpoint.
    filtered = ids.filter((id) => /^gemini/i.test(id));
  } else {
    // OpenAI and custom: keep GPT / o-series, exclude non-chat.
    const exclude = /embed|whisper|tts|dall-e|moderation|realtime|audio|search-preview|babbage|davinci|curie|ada/i;
    const include = /^(gpt-|o[0-9]|chatgpt)/i;
    filtered = ids.filter((id) => include.test(id) && !exclude.test(id));
  }

  return filtered.sort();
}

async function safeText(res) {
  try { return (await res.text()).slice(0, 300); } catch { return "(no body)"; }
}

app.post("/api/run", async (req, res) => {
  try {
    const { action, payload, provider, cloudSettings } = req.body || {};
    if (!action || !payload) {
      return res.status(400).json({ error: "Missing 'action' or 'payload'." });
    }

    const config = resolveConfig(provider, cloudSettings || {});

    if (provider === "cloud" && !config.apiKey) {
      return res.status(400).json({
        error:
          "No API key found. Open the settings drawer (click the Cloud chip), " +
          "select your provider, and paste your API key. " +
          "Alternatively set OPENAI_API_KEY or ANTHROPIC_API_KEY in the server .env file.",
      });
    }

    const messages = buildMessages(action, payload);
    const result   = await runChat(config, messages);
    res.json({ result, ran: { provider: config.provider, model: config.model || "(default)" } });
  } catch (err) {
    console.error("[api/run]", err.message);
    res.status(500).json({ error: err.message || "Something went wrong running the model." });
  }
});

// ── HTTPS / HTTP startup ──────────────────────────────────────────────────────

async function start() {
  let httpsOptions = null;
  try {
    const devCerts = await import("office-addin-dev-certs");
    httpsOptions   = await devCerts.getHttpsServerOptions();
  } catch {
    // dev certs not installed yet — run `npm run setup` to fix
  }

  if (httpsOptions) {
    https.createServer(httpsOptions, app).listen(PORT, () => {
      console.log(`AI Mail Assistant running at https://localhost:${PORT}`);
      console.log("Sideload manifest.xml in Outlook. Default: local Ollama.");
    });
  } else {
    http.createServer(app).listen(PORT, () => {
      console.log(`AI Mail Assistant running at http://localhost:${PORT} (HTTP fallback).`);
      console.log("Run `npm run setup` to install HTTPS dev certs — Outlook requires HTTPS.");
    });
  }
}

start();
