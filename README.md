# Your Text, Your Style – Text → PowerPoint (Template Clone)

Turn bulk text, markdown, or prose into a fully formatted PowerPoint that **inherits your uploaded template's** look & feel — colors, layouts, and images — no AI image generation.

## ✨ Features
- Paste long text or markdown.
- Optional one-line guidance (e.g., "turn into investor pitch deck").
- Bring your own LLM API key: **OpenAI**, **Anthropic**, or **Gemini** (keys are never stored).
- Upload a `.pptx`/`.potx` template or presentation — we start from it, so theme & styles carry over.
- Reuse images present in your template; **no new images** are generated.
- Auto-select layouts and number of slides based on content (LLM-driven).
- Optional: auto speaker notes via LLM.
- Download a generated `.pptx`.

## 🧱 Tech Stack
- **Flask** backend
- **python-pptx** for PPTX manipulation
- **HTMX + Tailwind** frontend (no JS build step)
- **requests** for provider-agnostic LLM calls

## ⚙️ Local Setup
```bash
git clone https://github.com/yourname/text-to-pptx-template-cloner.git
cd text-to-pptx-template-cloner
python -m venv .venv && source .venv/bin/activate  # on Windows: .venv\Scripts\activate
pip install -r requirements.txt
python app.py
# Visit http://localhost:8000
```

## 🚀 Deploy (Render)
1. Push to a public GitHub repo.
2. On Render.com, create a **Web Service**:
   - Runtime: Python 3.11+
   - Build command: `pip install -r requirements.txt`
   - Start command: `gunicorn -w 2 -b 0.0.0.0:8000 app:app`
3. Deploy. No environment variables needed. (Keys are entered by users at runtime.)

## 🧠 How it works (short write-up)
**Parsing & Slide Mapping.** The app prompts the chosen LLM to output a JSON outline:
```json
{ "slides": [ { "title": "T", "bullets": ["..."], "notes": "" }, ... ] }
```
It asks for concise bullets and chooses **6–15 slides** based on content & guidance. If an LLM is unavailable, a fallback splitter chunks the text by sections and creates simple slides.

**Applying Visual Style.** We **start from your uploaded `.pptx/.potx`** as the base presentation, so theme colors, fonts, and masters are preserved. New slides use existing layouts (title, title+content, etc). Any images found inside the template (from `ppt/media`) are **reused** by filling picture placeholders; if none exist, a small image is placed on the slide. We never generate new images.

## 🔐 Privacy & Security
- **No API keys are stored or logged.** They're only used to call your selected provider.
- Uploaded files are processed in-memory and discarded after response.
- Basic file-size cap (25 MB) to keep processing reasonable.

## 🧪 API Providers
Supply your own key and a model name (examples):
- OpenAI: `gpt-4o-mini`, `gpt-4o`, `o4-mini`, etc.
- Anthropic: `claude-3-5-sonnet-20240620`, etc.
- Gemini: `gemini-1.5-pro`

> Note: Endpoints are kept simple and may require updates to match provider changes.

## 🧩 Optional Enhancements
- Previews before download (render thumbnail of slides).
- Better tone presets (Sales, Research, Investor, Technical).
- Robust error handling, retries, and larger file support.
- Support for section breaks and multi-column layouts.

## 📝 Notes
- Exact 1:1 layout cloning isn't guaranteed, but starting from your file preserves most styles.
- Slide cap at 25 to avoid runaway outputs.
- This project is MIT licensed.
