# Changelog — AI Mail Assistant

All changes made in this session (Jun 21, 2026).

---

## 1. Cloud Provider Configuration UI

**Problem:** The Cloud toggle existed in the drawer but there was no way to enter API keys, choose a provider, or configure a model from within the add-in. Users had to manually edit the server's `.env` file.

**Changes:**
- Expanded the settings drawer with a full **Cloud settings form** when Cloud is selected, and a **Local (Ollama) settings form** when Local is selected. The two sections are mutually exclusive — only the relevant one is visible at any time.
- Cloud form includes:
  - **Provider selector** (OpenAI / Anthropic / Google / Custom) as a segmented control
  - **API Key** field (password type) with a show/hide toggle button
  - **Model** field pre-filled with the provider's recommended default, with a live-populated dropdown
  - **Base URL** field, pre-filled and locked to the provider default, with a **"Reset to default"** link that appears only when the URL has been changed
- Local form includes:
  - **Model** override (optional, defaults to `llama3.2`)
  - **Ollama URL** override (optional, defaults to `http://localhost:11434`)
- All settings are persisted to **`localStorage`** so they survive pane reloads.
- The Local/Cloud toggle state is also persisted — selecting Cloud and closing the pane no longer resets back to Local.

---

## 2. Per-Provider Settings

**Problem:** There was a single API key field. Switching between OpenAI and Anthropic would clear the previously entered key.

**Changes:**
- Each cloud provider now has its own independent slot in `localStorage` for:
  - `apiKey_{provider}` — API key
  - `model_{provider}` — selected model
  - `baseUrl_{provider}` — base URL override
- Switching providers restores that provider's saved values. Keys are never lost when switching.
- **Migration:** Old single-key format is automatically migrated to the appropriate per-provider slot on first load.

---

## 3. Pre-filled and Locked Base URLs

**Problem:** Base URLs were hidden and undocumented. Users had no way to inspect or override them from the UI.

**Changes:**
- Base URL is now **always visible** for all providers, pre-filled with the correct default:
  | Provider | Default Base URL |
  |---|---|
  | OpenAI | `https://api.openai.com/v1` |
  | Anthropic | `https://api.anthropic.com` |
  | Google | `https://generativelanguage.googleapis.com/v1beta/openai` |
  | Custom | *(user fills in)* |
- The field is editable. A **"Reset to default"** link appears inline next to the label whenever the value differs from the provider default.
- This makes future API URL changes easy to handle with a single edit.

---

## 4. Google Gemini Support

**Problem:** Google Gemini was not supported as a provider.

**Changes:**
- Added **Google** as a first-class provider in the UI alongside OpenAI, Anthropic, and Custom.
- Uses Google's OpenAI-compatible endpoint (`generativelanguage.googleapis.com/v1beta/openai`) — no new provider code path needed in the backend.
- A contextual note is shown when Google is selected: "Get a key at **aistudio.google.com** (free tier available)."
- Keys obtained from [aistudio.google.com](https://aistudio.google.com) work directly with the `Bearer` auth that the OpenAI-compatible path uses.

---

## 5. Dynamic Model List (Live from Provider API)

**Problem:** Model lists were hardcoded in the source code and went stale quickly.

**Changes:**
- Added `POST /api/models` endpoint on the server. It proxies the provider's official models endpoint using the user's API key and returns only chat-capable model IDs.
- The task pane calls this endpoint in two situations:
  1. **Provider switch** — fetches immediately if a key is already saved for that provider.
  2. **Key input** — fetches 900 ms after the user stops typing (debounced).
- Results are filtered to chat models only:
  - **OpenAI/Custom:** keeps `gpt-*` and `o[0-9]-*` families; excludes embedding, TTS, Whisper, DALL-E, moderation models.
  - **Anthropic:** no filtering needed (endpoint only returns text models).
  - **Google:** keeps models whose ID starts with `gemini`.
- The dropdown shows a **"Loading models…"** status while fetching, **"N models loaded"** on success (auto-clears after 3 s), and **"Could not load — using defaults"** on error, reverting to the static fallback list.
- The static fallback list (hardcoded suggestions) always seeds the dropdown first so it is never empty.

---

## 6. Email Body Processing — HTML Reading and Signature Stripping

**Problem:** The add-in read email bodies as plain text, which lost structure. Signatures and quoted/replied content were sent to the AI unnecessarily.

**Changes:**
- Email body is now read as **`CoercionType.Html`** instead of plain text, giving the AI cleaner, structure-aware input.
- `stripEmailNoise(html)` removes before sending to the AI:
  - Known signature containers by ID/class: `#Signature`, `.gmail_signature`, `.MsoSignature`, `.apple-mail-signature`, etc.
  - Outlook and Gmail quoted/forwarded sections: `.gmail_quote`, `[id^='divRplyFwdMsg']`, `.moz-cite-prefix`, etc.
  - `<hr>` elements that are followed by email-header-style text (`From:`, `Sent:`, `To:`) — i.e., genuine reply/forward dividers only, not content separators.
  - RFC 3676 `-- ` text delimiter (the standard plain-text signature separator).
- `htmlToText(html)` converts the cleaned HTML to plain text with proper line breaks for block elements and `•` bullets for list items.
- **Fallback:** If `CoercionType.Html` fails (older Outlook / plain-text message), the add-in automatically retries with `CoercionType.Text`.
- **Safety net:** If signature stripping removes everything, the full unstripped body is used instead of sending an empty string to the AI.

---

## 7. Formatted AI Output (Markdown → HTML)

**Problem:** AI responses were plain text displayed in a `pre-wrap` box. Inserting into emails stripped all formatting.

**Changes:**
- All three prompts (Summarize, Reply, Tone) now include explicit **Markdown formatting instructions** asking the model to use:
  - Blank lines between paragraphs
  - `**bold**` for key names, dates, decisions
  - `- bullet` lists for action items
  - Numbered lists when order matters
  - `###` headings only when content clearly has sections
- `markdownToHtml(md)` converts the AI's Markdown output to safe HTML:
  - Escapes all input HTML entities first (XSS-safe)
  - Processes headings, bullet lists, numbered lists, bold, italic
  - Wraps paragraphs in `<p>` tags; single newlines within a paragraph become `<br>`
- The result panel renders HTML via `innerHTML` with proper CSS styles for `<p>`, `<ul>`, `<ol>`, `<strong>`, `<em>`, `<h3>`.
- **Copy** copies the raw Markdown (universally pasteable anywhere).
- **Insert / Replace** uses `CoercionType.Html` so formatting is preserved in the email.

---

## 8. Insert vs. Replace Draft

**Problem:** The Insert button always used `setSelectedDataAsync` regardless of action, which was wrong for the Tone tab (a full rewrite should replace the entire body, not insert at cursor).

**Changes:**
- **Tone tab** → button label is **"Replace draft"**; uses `body.setAsync(html, { coercionType: Html })` to replace the entire compose body.
- **Summarize / Reply tabs** → button label is **"Insert"**; uses `body.setSelectedDataAsync(html, { coercionType: Html })` to insert at the cursor position.
- Button label updates automatically when the tab changes.

---

## 9. Server-Side Settings Override

**Problem:** The server only read cloud provider settings from `.env`. Client-provided settings were ignored, making the UI settings drawer useless.

**Changes:**
- `POST /api/run` now accepts a `cloudSettings` object in the request body.
- `resolveConfig()` on the server applies client-provided values with this priority: **request body → `.env` file → hardcoded default**.
- Fields that can be overridden per request:
  - `cloudProvider` — `"openai"` | `"anthropic"` | `"google"` | `"custom"`
  - `apiKey` — the user's API key from the drawer
  - `cloudModel` — model name
  - `cloudBaseUrl` — base URL (custom endpoints, Groq, Together.ai, Mistral, etc.)
  - `ollamaModel` — local model override
  - `ollamaBaseUrl` — local Ollama URL override
- Keys travel over **localhost HTTPS only** and are never stored server-side.

---

## 10. `/api/health` Improvements

- Returns `serverProviders: { openai: bool, anthropic: bool }` so the UI can show whether server-side `.env` keys exist as a fallback.

---

## 11. Bug Fixes

| Bug | Fix |
|---|---|
| Cloud toggle reset to Local on every pane reload | `state.provider` is now saved in `localStorage` and restored on load |
| `[hidden]` attribute ignored on flex elements | Added `[hidden] { display: none !important; }` to CSS — ID/class `display:flex` rules were overriding the browser's default `[hidden]` style |
| Anthropic 404 error with `claude-sonnet-4-5` | Invalid shorthand model IDs removed; replaced with correct API IDs (`claude-sonnet-4-6`, `claude-3-5-sonnet-latest`, etc.) |
| `<hr>` over-stripping removed body content | Horizontal rules are now only stripped when followed by email-header text, not unconditionally |
| Server running stale code | Server must be restarted after code changes; `npm start` in the project root |
| `.env` value regex captured trailing comments | Regex updated from `(.*)\s*$` to `(.*?)\s*$` to avoid capturing trailing whitespace |

---

## 12. Updated Model Defaults (June 2026)

| Provider | Default | Source |
|---|---|---|
| OpenAI | `gpt-5.5` | [OpenAI models](https://developers.openai.com/api/docs/models/all) |
| Anthropic | `claude-sonnet-4-6` | [Anthropic models overview](https://platform.claude.com/docs/en/about-claude/models/overview) |
| Google | `gemini-2.5-flash` | [Gemini API models](https://ai.google.dev/gemini-api/docs/models) |
| Local | `llama3.2` | Ollama default |

---

## 13. Reply Tab — Set Full Draft Body Instead of Cursor Insert

**Problem:** The Reply tab's "Insert" button used `setSelectedDataAsync()`, which inserts the generated reply at the cursor position inside the compose body. This was inappropriate — a fully drafted reply should replace the compose body entirely, not splice into it.

**Changes:**
- The Reply tab's action button is now labelled **"Set as draft"** (was "Insert").
- `insertResult()` now uses `body.setAsync(html, { coercionType: Html })` for the reply action, replacing the entire compose body with the AI-generated draft — the same behaviour as the Tone tab's "Replace draft".
- Summarize retains cursor-insert (`setSelectedDataAsync`) since inserting a summary at the cursor position is intentional.

---

## 14. Full Thread Context for Reply Generation

**Problem:** When generating a reply, `stripEmailNoise()` removed all quoted/forwarded thread content before sending the body to the AI. The AI could only see the top-level message and had no awareness of the conversation history.

**Changes:**
- Added `getBodyWithThread()` — reads the email body, strips only signature blocks and the RFC 3676 `-- ` delimiter, but **preserves quoted reply chain content** so the full conversation history reaches the AI.
- Added `stripSignaturesOnly()` — a lighter version of `stripEmailNoise()` that only targets signature containers and the signature delimiter, leaving thread content intact.
- `run()` now routes by action: `getBodyWithThread()` for reply, `getBody()` (full noise stripping) for summarize and tone.
- A `hasThread: true` flag is sent in the API payload for reply actions, which the server uses to inject a thread-awareness instruction into the system prompt.
- The reply system prompt now explicitly instructs the model to read all messages in the thread before composing a response.

---

## 15. Rich Text Formatting in Outlook Compose Body

**Problem:** Even though `CoercionType.Html` was used for insertion, the generated HTML lacked inline styles. Outlook's Word-based compose engine strips class-based CSS when content is inserted via Office.js, so bold, bullets, and paragraph spacing were lost. Additionally, the "Copy" button used `writeText()` — pasting into Outlook from the clipboard gave plain Markdown, not formatted text.

**Changes:**
- Added `markdownToHtmlForOutlook(md)` — converts Markdown to HTML then walks the DOM with `DOMParser` to attach `style=` attributes to every element. Outlook preserves inline styles:
  - `<p>` → `margin: 0 0 10px 0; line-height: 1.6`
  - `<strong>`, `<b>` → `font-weight: bold`
  - `<em>`, `<i>` → `font-style: italic`
  - `<u>` → `text-decoration: underline`
  - `<ul>` → `margin / list-style-type: disc`
  - `<ol>` → `margin / list-style-type: decimal`
  - `<li>` → `margin: 4px 0; line-height: 1.6`
  - `<h1>`–`<h6>` → appropriate `font-size` and `font-weight`
  - `<blockquote>` → `border-left: 3px solid #ccc; color: #555; font-style: italic`
  - `<a>` → `color: #0563C1; text-decoration: underline`
- All insert paths (Reply "Set as draft", Tone "Replace draft", Summarize "Insert") now call `markdownToHtmlForOutlook()` instead of `markdownToHtml()`.
- `copyResult()` now writes a `ClipboardItem` with both `text/html` (inline-styled HTML) and `text/plain` (raw Markdown) MIME types — pasting into Outlook or Word gives rich formatted text; pasting into a plain-text field still works. Falls back to `writeText()` if `ClipboardItem` is unsupported in that context.
- The result panel display continues to use class-based CSS (`markdownToHtml()` only, no inline styles) so the panel appearance is controlled by `taskpane.css`.

---

## 16. Expanded Tone Set — 11 Tones with Per-Tone Prompts

**Problem:** Only 5 generic tones were available (Professional, Friendly, Concise, Assertive, Apologetic), with a single generic system prompt for all of them. The AI had no specific guidance about what each tone actually means, leading to inconsistent or shallow rewrites.

**Changes:**
- Added 6 new tones, bringing the total to **11**:
  | Tone | Character |
  |---|---|
  | **Government** | Formal official correspondence — no contractions, passive voice acceptable, honorifics, proper salutations (Dear Sir/Madam), Yours faithfully/sincerely |
  | **Formal** | Elevated vocabulary, traditional letter structure, no colloquialisms |
  | **Diplomatic** | Tactful, balanced; acknowledges all perspectives before stating own position; avoids confrontational language |
  | **Direct** | No pleasantries; lead with the ask; short declarative sentences; minimal greeting/closing |
  | **Empathetic** | Acknowledges recipient's feelings/situation first; warm supportive language throughout |
  | **Casual** | Conversational, contractions encouraged, light tone, first-name basis |
- Added `TONE_GUIDES` map in `prompts.js` — each tone has specific behavioural instructions (language register, greeting style, closing style, what to avoid) that are injected verbatim into the system prompt.
- The tone system prompt is now: `"You rewrite email drafts to sound [tone desc]. [tone-specific guidance]. Preserve meaning exactly. Adjust greeting and closing to match the tone."` — giving the model precise behavioural anchors instead of just a label.
- Tone chip buttons in the UI are slightly smaller (12px / 5px–10px padding) so all 11 fit in 2–3 rows without overflowing.

---

## 17. Improved AI Prompts

**Problem:** Prompts were functional but lacked specificity. The reply prompt did not instruct the model to avoid restating what was already said. The format instructions did not mention blockquotes, italic, or the no-preamble rule clearly.

**Changes:**

**Reply prompt:**
- Added conditional thread-awareness instruction when `hasThread` is set: instructs the model to read all messages before replying.
- Added: "Keep the reply focused — do not restate what the original message already said."
- Greeting guidance now covers both informal (`Hi [Name],`) and formal (`Dear [Name],`) registers with explicit contextual note.
- User message changed from `--- Original message ---` to `--- Email thread ---` to better reflect that full thread content may be present.

**Tone prompt:**
- System prompt now opens with the specific tone name and a prose description of that tone's character (from `TONE_GUIDES`).
- Added: "Adjust the greeting and closing to match the tone."

**Format instructions (all actions):**
- Added `*italic*` and `> blockquote` to the supported syntax list.
- Strengthened the no-preamble rule: "Do not add a preamble like 'Here is the draft:' — start directly with the greeting or content."

---

## 18. Fully Offline — No External CDN Dependencies

**Problem:** The add-in loaded two libraries from external CDNs at runtime, requiring internet access even when using a local Ollama model:
1. `https://appsforoffice.microsoft.com/lib/1/hosted/office.js` — Microsoft's Office.js CDN
2. `https://cdn.jsdelivr.net/npm/marked/marked.min.js` — jsDelivr CDN for Markdown parsing

**Changes:**

### Office.js — served locally via npm
- Installed `@microsoft/office-js` as an npm dependency.
- Added a static route in `server.js`:
  ```js
  app.use("/vendor/office-js", express.static(
    path.join(root, "node_modules", "@microsoft", "office-js", "dist")
  ));
  ```
- Updated `taskpane.html`: `<script src="/vendor/office-js/office.js"></script>`
- The `office.js` umbrella loader at the dist root is the same file Microsoft's CDN serves. It is served over the existing local HTTPS server at `localhost:3000`.
- **Note:** `@microsoft/office-js` was deprecated by Microsoft in favour of CDN-only usage. The installed version (1.1.110) covers the full Mailbox 1.3 API surface used by this add-in and will continue to work indefinitely. Run `npm update @microsoft/office-js` if new Office.js API features are needed.

### marked.js — replaced with a built-in converter
- Removed the CDN `<script>` tag for marked.js entirely.
- Rewrote `markdownToHtml()` as a self-contained parser with no external dependency. Supported syntax:
  - Headings `#` through `######`
  - Bold `**text**` / `__text__`, italic `*text*` / `_text_`, bold-italic `***text***`
  - Strikethrough `~~text~~`
  - Inline code `` `code` ``
  - Links `[text](url)` → `<a href>`
  - Bullet lists (`-`, `*`, `+`) with 2-level indent nesting
  - Ordered lists (`1.`, `2.`, …) with 2-level indent nesting
  - Blockquotes `> text`
  - Horizontal rules `---` / `***` / `___`
  - Blank-line-separated `<p>` paragraphs with `<br>` for in-paragraph line breaks
- All HTML entities are escaped before Markdown patterns are applied (XSS-safe).
- The old `_legacyMarkdownToHtml()` fallback function has been removed (no longer needed).

**Result:** The add-in now runs completely offline when using a local Ollama model. The only internet traffic is AI API calls when the Cloud provider is selected.
