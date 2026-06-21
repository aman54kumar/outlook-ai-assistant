# AI Mail Assistant for Outlook

A privacy-first AI assistant that lives inside Outlook. It can **summarize a thread**, **draft a reply**, and **adjust the tone** of a message across **11 distinct tones** — running on a **local model (Ollama) by default**, so sensitive email never leaves the machine. A single toggle switches to a cloud provider (OpenAI, Anthropic, or Google Gemini) when you want more power, and you can configure the provider, API key, and model **right inside the task pane** — no file editing required.

Built with **Office.js** and **runs fully offline** with a local model — no external CDNs, no telemetry, no accounts, no data collection.

> **Live demo (no install):** _add your GitHub Pages URL here_ — try the three features on a sample email right in the browser.

<!-- Add a screenshot or short GIF of the task pane here once you've sideloaded it:
     ![AI Mail Assistant task pane](docs/screenshot.png) -->

---

## Why it's built this way

Most AI email tools send your messages to a vendor's servers. That's a non-starter in finance, legal, healthcare, and government inboxes. This add-in is **local-first**: by default every request goes to a model running on your own machine through a small local server, and when you do use a cloud provider, the **API key travels over localhost HTTPS only and is never stored server-side**. The privacy posture is always visible in the pane via the runtime chip — `Local` or `Cloud`.

With a local Ollama model the add-in runs **completely offline** — both Office.js and the Markdown renderer are served locally, so there are no external CDN requests. The only outbound traffic is the AI API call itself, and only when the Cloud provider is selected.

## Features

| Feature | Where | What it does |
| --- | --- | --- |
| **Summarize thread** | Reading | A 2–3 sentence summary plus action items and open questions. |
| **Draft reply** | Composing | A ready-to-send reply, optionally guided by a one-line intent. Reads the **full thread history** for context and sets it as the complete draft body. |
| **Adjust tone** | Composing | Rewrites your draft across **11 tones** — Professional, Friendly, Concise, Assertive, Apologetic, Government, Formal, Diplomatic, Direct, Empathetic, or Casual — with meaning and facts preserved. |

Results come back **richly formatted**: the model is prompted to use Markdown (bold, lists, headings, blockquotes), which is rendered in the pane and inserted into Outlook with inline styles so formatting survives. **Copy** puts both rich HTML and plain Markdown on the clipboard. Email bodies are read as HTML with signatures and quoted noise stripped before being sent to the AI (the reply action keeps the thread intact for context).

## How it runs

```
Outlook task pane ──▶ local server (HTTPS) ──▶ Ollama        (local, default — nothing leaves the device)
   (Office.js)         proxies requests &   ├──▶ OpenAI / OpenAI-compatible
                       applies pane settings├──▶ Anthropic
                                            └──▶ Google Gemini   (cloud toggle)
```

Provider, API key, model, and base URL are configured in the pane's settings drawer and sent with each request; the server applies them with the priority **request body → `.env` file → built-in default**. The same prompt and provider modules power both the add-in and the browser demo, so there's one place to change behaviour.

## Quick start (local, fully private)

**Prerequisites:** Node.js 18+, Outlook (desktop or web), and [Ollama](https://ollama.com) for the local model.

```bash
# 1. Install dependencies
npm install

# 2. Pull a local model (once)
ollama pull llama3.2

# 3. Trust the local HTTPS dev certificate (Office requires HTTPS)
npm run setup

# 4. Start the add-in server
npm start            # serves https://localhost:3000

# 5. Sideload into Outlook
npm run sideload     # or sideload manifest.xml manually (see below)
```

Open or compose an email, click **AI Assistant** on the ribbon, and the pane opens. That's it — no key, no cloud.

### Turning on the cloud toggle (optional)

Flip the pane's runtime chip to **Cloud** and open the settings drawer. From there you can:

- Pick a provider — **OpenAI**, **Anthropic**, **Google**, or **Custom** (any OpenAI-compatible endpoint such as Groq, Together.ai, or Mistral).
- Paste an **API key** (with show/hide), which is kept **per provider** so switching providers never loses a key.
- Choose a **model** from a dropdown that is **populated live from the provider's API** once a key is entered, falling back to sensible defaults.
- Inspect or override the **base URL**, which is pre-filled per provider with a one-click "Reset to default".

All of this is stored in the browser's `localStorage` and persists across reloads. For Google, get a free-tier key at [aistudio.google.com](https://aistudio.google.com).

Prefer to keep keys off the client entirely? You can still set them server-side:

```bash
cp .env.example .env
# set CLOUD_PROVIDER=anthropic (or openai/google) and the matching API key
```

Restart the server. The `.env` value is used as a fallback when the pane doesn't supply its own. The current model defaults are:

| Provider | Default model |
|---|---|
| OpenAI | `gpt-5.5` |
| Anthropic | `claude-sonnet-4-6` |
| Google | `gemini-2.5-flash` |
| Local | `llama3.2` |

### Sideloading manually

If you'd rather not use the script, follow Microsoft's guide for your platform and point it at `manifest.xml`:
[Sideload Outlook add-ins for testing](https://learn.microsoft.com/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing).

## Try it without Outlook

The `/demo` folder is a static page that mirrors the pane. In **Demo mode** it returns realistic pre-written results instantly — no setup. Switch to **Live** to call a provider with your own key (kept in the browser tab only). Host it free on GitHub Pages by enabling Pages from the repository root.

## Project structure

```
outlook-ai-assistant/
├── manifest.xml              # Outlook add-in manifest (read + compose ribbon buttons)
├── server/
│   └── server.js             # HTTPS host + /api/run, /api/models, /api/health
├── src/
│   ├── taskpane/             # the in-Outlook UI: settings drawer, HTML body
│   │                         #   reading + signature stripping, Markdown rendering
│   ├── commands/             # ribbon function file
│   └── ai/
│       ├── prompts.js        # prompt builders + TONE_GUIDES (11 tones)
│       └── providers.js      # Ollama / OpenAI-compatible / Anthropic / Google behind one call
├── demo/                     # zero-setup browser demo (GitHub Pages)
└── assets/                   # icons
```

Office.js is vendored locally via the `@microsoft/office-js` npm package and served from `/vendor/office-js`, so no CDN is needed at runtime.

## Tech

JavaScript (ES modules), Office.js (served locally), and Express. A provider layer sits over Ollama, OpenAI-compatible APIs, Anthropic, and Google Gemini (via its OpenAI-compatible endpoint). Markdown rendering is a built-in, dependency-free converter — no `marked.js` CDN. No build step; the task pane uses native ES modules.

The server exposes three endpoints: `POST /api/run` (proxies a generation request and applies pane settings), `POST /api/models` (lists chat-capable models live from the selected provider), and `GET /api/health` (reports reachability and which server-side `.env` keys exist).

## Notes

This is an independent, open-source project built from scratch as a reference implementation. It carries no third-party code or data. Configure your own models and keys before use.

## License

MIT — see [LICENSE](LICENSE).
