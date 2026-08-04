# AI PowerPoint Generator — Python POC

Generate a valid, editable `.pptx` from a natural-language prompt. The LLM
drafts a structured deck outline; the backend assembles raw OOXML and zips it
into a `.pptx` — no `python-pptx` or other PowerPoint library is used.

## Layout

```
ppt_gen/
├── pptx_builder.py   # Pure OOXML + zipfile. Frontend-agnostic. Reusable.
├── ai_client.py      # OpenRouter client (OpenAI SDK). Structured JSON output.
├── generator.py      # Orchestration: prompt -> outline -> .pptx bytes.
├── app.py            # Streamlit demo UI. Thin — only collects input + download.
├── pyproject.toml    # uv project + dependencies
├── README.md
└── POC.md            # Model choice, OOXML file map, architecture, Phase-2 hookup.
```

The frontend (`app.py`) only imports `generator.generate_pptx(...)`. Every
piece of business logic sits in modules that the main application can import
directly without any Streamlit dependency.

## Setup

```bash
cd ppt_gen
uv sync
```

## API key

Get a free key at <https://openrouter.ai/keys>, then:

```bash
export OPENROUTER_API_KEY=sk-or-...
```

## Run

```bash
uv run streamlit run app.py
```

Open the URL Streamlit prints, enter a prompt, pick a slide count, and click
**Generate .pptx**.

## Use the backend directly (no Streamlit)

```bash
uv run python -c "
from generator import generate_pptx
pptx_bytes, filename = generate_pptx(
    prompt='A 6-slide overview of quantum computing for a business audience',
    slide_count=6,
)
open(filename, 'wb').write(pptx_bytes)
"
```

## Model

Default: **`google/gemini-2.0-flash-exp:free`** via OpenRouter. See
[`POC.md`](./POC.md) for the rationale and alternatives. Override per-call by
passing `model=...` to `generate_pptx`.

## Validation

The generated file opens in PowerPoint, Keynote, Google Slides, and
LibreOffice with fully editable text. To verify programmatically:

```bash
uv run python -c "from generator import generate_pptx; \
b,_ = generate_pptx('AI in healthcare', slide_count=4); \
open('out.pptx','wb').write(b); print('ok', len(b))"
soffice --headless --convert-to pdf out.pptx   # optional PDF sanity check
```
