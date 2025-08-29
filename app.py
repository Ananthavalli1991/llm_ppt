
import io
import os
import re
import tempfile
from typing import List, Optional

from fastapi import FastAPI, File, Form, UploadFile, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import PP_PLACEHOLDER
import requests

app = FastAPI(title="Presentify", version="1.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

def call_llm(provider: str, api_key: str, text: str, guidance: str) -> Optional[str]:
    provider = (provider or "").lower().strip()
    prompt = f"""
You are a slide architect. Rewrite the user's content into a slide outline in JSON.

Return JSON with this shape (and nothing else):
{{
  "slides": [
    {{"title": "Slide Title", "bullets": ["point 1", "point 2"], "notes": "optional speaker notes"}}
  ]
}}

Guidance: {guidance or "general"}

User content:
{text}
""".strip()

    try:
        if provider == "openai":
            url = "https://api.openai.com/v1/chat/completions"
            headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
            data = {
                "model": "gpt-4o-mini",
                "messages": [
                    {"role": "system", "content": "You convert content into a concise slide outline JSON."},
                    {"role": "user", "content": prompt},
                ],
                "temperature": 0.2,
            }
            r = requests.post(url, headers=headers, json=data, timeout=45)
            r.raise_for_status()
            return r.json()["choices"][0]["message"]["content"]
        elif provider == "anthropic":
            url = "https://api.anthropic.com/v1/messages"
            headers = {
                "x-api-key": api_key,
                "anthropic-version": "2023-06-01",
                "content-type": "application/json",
            }
            data = {
                "model": "claude-3-haiku-20240307",
                "max_tokens": 1200,
                "messages": [{"role": "user", "content": prompt}],
            }
            r = requests.post(url, headers=headers, json=data, timeout=45)
            r.raise_for_status()
            return "".join([b["text"] for b in r.json()["content"] if b.get("type")=="text"])
        elif provider == "gemini":
            url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key={api_key}"
            data = {"contents": [{"parts": [{"text": prompt}]}]}
            r = requests.post(url, json=data, timeout=45)
            r.raise_for_status()
            js = r.json()
            return js["candidates"][0]["content"]["parts"][0]["text"]
        else:
            return None
    except Exception:
        return None

def split_into_sections(text: str) -> List[dict]:
    text = text.strip()
    if not text:
        return []
    lines = text.splitlines()
    slides = []
    current_title = None
    current_points = []

    def flush():
        nonlocal current_title, current_points
        if current_title or current_points:
            slides.append({
                "title": current_title or (current_points[0][:60] if current_points else "Slide"),
                "bullets": [b for b in current_points if b.strip()][:8]
            })
        current_title, current_points = None, []

    import re as _re
    for ln in lines:
        if _re.match(r'^\s*#{1,3}\s+', ln):
            flush()
            current_title = _re.sub(r'^\s*#{1,3}\s+', '', ln).strip()
        elif _re.match(r'^\s*[-*+]\s+', ln):
            current_points.append(_re.sub(r'^\s*[-*+]\s+', '', ln).strip())
        elif ln.strip():
            parts = _re.split(r'(?<=[.!?])\s+', ln.strip())
            for p in parts:
                if p:
                    current_points.append(p.strip())
    flush()
    if len(slides) > 30:
        slides = slides[:30]
    if len(slides) == 1 and len(slides[0]["bullets"]) > 10 and not slides[0]["title"]:
        bullets = slides[0]["bullets"]
        slides = []
        for i in range(0, len(bullets), 6):
            slides.append({"title": f"Section {len(slides)+1}", "bullets": bullets[i:i+6]})
    return slides

def parse_outline(text: str, guidance: str, provider: str, api_key: str) -> List[dict]:
    if api_key and provider:
        llm = call_llm(provider, api_key, text, guidance)
        if llm:
            try:
                import json
                js = json.loads(llm)
                if isinstance(js, dict) and "slides" in js and isinstance(js["slides"], list):
                    normalized = []
                    for s in js["slides"]:
                        title = (s.get("title") or "").strip() or "Slide"
                        bullets = [str(b).strip() for b in (s.get("bullets") or []) if str(b).strip()][:8]
                        notes = str(s.get("notes") or "").strip()
                        normalized.append({"title": title, "bullets": bullets, "notes": notes})
                    if normalized:
                        return normalized
            except Exception:
                pass
    slides = split_into_sections(text)
    for s in slides:
        s["notes"] = ""
    return slides

from pptx.dml.color import RGBColor
from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER

def extract_template_images(prs: Presentation, tmpdir: str):
    paths = []
    pkg = prs.part.package
    for idx, ipart in enumerate(getattr(pkg, "image_parts", [])):
        ext = os.path.splitext(ipart.partname)[1] or ".png"
        pth = os.path.join(tmpdir, f"tpl_img_{idx}{ext}")
        with open(pth, "wb") as f:
            f.write(ipart.blob)
        paths.append(pth)
    return paths

def apply_branding_defaults(prs: Presentation):
    try:
        theme = prs.presentation_part.theme.part.themeElements
        accent = theme.clrScheme.accent1.sysClr.lastClr if hasattr(theme.clrScheme.accent1, "sysClr") else None
        if not accent and hasattr(theme.clrScheme.accent1, "srgbClr"):
            accent = theme.clrScheme.accent1.srgbClr.val
        if accent:
            rgb = bytes.fromhex(accent)
            return RGBColor(rgb[0], rgb[1], rgb[2])
    except Exception:
        pass
    return RGBColor(30, 41, 59)

def pick_layout(prs: Presentation):
    candidates = [l for l in prs.slide_layouts]
    for l in candidates:
        name = (getattr(l, 'name', '') or '').lower()
        if 'title and content' in name:
            return l
    for l in candidates:
        name = (getattr(l, 'name', '') or '').lower()
        if 'title only' in name or 'title' in name:
            return l
    return prs.slide_layouts[0]

def build_pptx(template_bytes: bytes, slides: list) -> bytes:
    prs = Presentation(io.BytesIO(template_bytes))
    brand_color = apply_branding_defaults(prs)

    import tempfile as _tf
    with _tf.TemporaryDirectory() as td:
        reuse_imgs = extract_template_images(prs, td)
        layout = pick_layout(prs)
        img_index = 0

        for s in slides:
            slide = prs.slides.add_slide(layout)

            if slide.shapes.title:
                slide.shapes.title.text = s["title"][:120]

            body = None
            for ph in slide.placeholders:
                if ph.placeholder_format.type in (PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.CONTENT):
                    body = ph
                    break

            if body:
                body.text = ""
                tf = body.text_frame
                tf.clear()
                if s["bullets"]:
                    p = tf.paragraphs[0]
                    p.text = s["bullets"][0]
                    p.level = 0
                    p.font.size = Pt(18)
                    p.font.color.rgb = brand_color
                    for b in s["bullets"][1:]:
                        p = tf.add_paragraph()
                        p.text = b
                        p.level = 0
                        p.font.size = Pt(18)
                        p.font.color.rgb = brand_color
            else:
                left, top, width, height = Inches(1), Inches(2), Inches(8), Inches(4)
                box = slide.shapes.add_textbox(left, top, width, height)
                tf = box.text_frame
                tf.clear()
                for i, b in enumerate(s["bullets"]):
                    p = tf.add_paragraph() if i else tf.paragraphs[0]
                    p.text = f"• {b}"
                    p.font.size = Pt(18)
                    p.font.color.rgb = brand_color

            if reuse_imgs:
                try:
                    img_path = reuse_imgs[img_index % len(reuse_imgs)]
                    img_index += 1
                    slide.shapes.add_picture(img_path, Inches(9), Inches(1.5), height=Inches(3))
                except Exception:
                    pass

            if s.get("notes"):
                slide.notes_slide.notes_text_frame.text = s["notes"][:1000]

        out = io.BytesIO()
        prs.save(out)
        out.seek(0)
        return out.getvalue()

class Health(BaseModel):
    ok: bool

@app.get("/health", response_model=Health)
def health():
    return {"ok": True}

@app.post("/api/generate")
async def generate(
    text: str = Form(...),
    guidance: str = Form(""),
    provider: str = Form(""),
    api_key: str = Form(""),
    template: UploadFile = File(...),
):
    if not template.filename.lower().endswith((".pptx", ".potx")):
        raise HTTPException(status_code=400, detail="Please upload a .pptx or .potx file.")
    tpl_bytes = await template.read()
    slides = parse_outline(text, guidance, provider, api_key)
    if not slides:
        raise HTTPException(status_code=400, detail="Could not parse any slides from the input.")
    try:
        pptx_bytes = build_pptx(tpl_bytes, slides)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to generate PowerPoint: {e}")
    headers = {"Content-Disposition": 'attachment; filename="presentify_output.pptx"'}
    return StreamingResponse(io.BytesIO(pptx_bytes),
                             media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                             headers=headers)

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=int(os.getenv("PORT", "8000")))
