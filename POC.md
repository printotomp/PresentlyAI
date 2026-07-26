# AI PowerPoint Generator — Python POC

## Objective

Generate a valid, editable `.pptx` file from a user prompt by having an LLM
draft a structured deck outline and programmatically assembling the OOXML
files. Implementation is Python, with a thin Streamlit frontend used only for
demonstration; all business logic lives in importable modules with no
frontend dependency.

## Selected OpenRouter Model

**`google/gemini-2.0-flash-exp:free`**

Why:
- **Free tier** on OpenRouter — zero cost for the POC's request volume.
- **Fast** (flash class) — deck outlines return in a few seconds.
- **Reliable JSON output** with `response_format={"type": "json_object"}`,
  validated against a Pydantic schema.
- **Multimodal**, so it can be trivially extended in Phase 2 when image
  references or vision inputs come into play.

Called via the OpenAI Python SDK pointed at
`https://openrouter.ai/api/v1`, so switching models is a one-line change in
`ai_client.py` (`DEFAULT_MODEL`). Solid free alternatives:
`meta-llama/llama-3.3-70b-instruct:free`,
`deepseek/deepseek-chat-v3.1:free`.

## OOXML File Structure

A `.pptx` is a ZIP archive. The minimum parts required for PowerPoint to
open the file as valid and editable:

```
[Content_Types].xml                            # MIME registry for every part
_rels/.rels                                    # Package-level rels -> ppt/presentation.xml
ppt/presentation.xml                           # Deck: slide list, slide size, master ref
ppt/_rels/presentation.xml.rels                # Rels -> slideMaster, each slide, theme
ppt/theme/theme1.xml                           # Required theme (colors, fonts, effects)
ppt/slideMasters/slideMaster1.xml              # Required master
ppt/slideMasters/_rels/slideMaster1.xml.rels   # Rels -> layout, theme
ppt/slideLayouts/slideLayout1.xml              # Required blank layout
ppt/slideLayouts/_rels/slideLayout1.xml.rels   # Rels -> master
ppt/slides/slideN.xml                          # One per slide
ppt/slides/_rels/slideN.xml.rels               # Rels -> layout
```

Key OOXML notes:
- Every part referenced from a `.rels` file must also appear in
  `[Content_Types].xml` (via `Default` extension or `Override` PartName).
- Positions and sizes use **English Metric Units** (914,400 EMU = 1 inch).
  The POC uses 16:9 slides at `12,192,000 x 6,858,000` EMU.
- Text lives inside `<p:sp>` shapes → `<p:txBody>` → `<a:p>` paragraphs →
  `<a:r>` runs → `<a:t>` text. Bullets are paragraphs with `<a:buChar>`.
- Theme, master, and at least one layout are **required** — PowerPoint
  refuses to open a deck without them.

## Architecture

```
[ Streamlit UI (app.py) ]                 [ Python backend ]
        │                                  generator.generate_pptx(...)
        │  prompt, slide_count                       │
        │──────────────────────────────────────────▶│  1. Validate input
        │                                            │  2. ai_client.generate_deck_outline
        │                                            │     -> OpenRouter (google/gemini-2.0-flash-exp:free)
        │                                            │     -> DeckOutline (Pydantic)
        │                                            │  3. pptx_builder.build_pptx(deck)
        │                                            │     -> render OOXML parts
        │                                            │     -> zipfile.ZipFile packages .pptx
        │  pptx_bytes, filename                      │
        │◀───────────────────────────────────────────│
        │
        │  4. st.download_button hands the file to the user
```

Modules:
- **`pptx_builder.py`** — pure OOXML + `zipfile`. No PowerPoint libraries.
  Public API: `build_pptx(DeckInput) -> bytes`. Frontend-agnostic.
- **`ai_client.py`** — OpenRouter client using the OpenAI SDK. Enforces a
  Pydantic `DeckOutline` schema. Frontend-agnostic.
- **`generator.py`** — orchestrates `prompt -> outline -> .pptx bytes`.
  The single entry point the main app will import: `generate_pptx(...)`.
- **`app.py`** — Streamlit demo. Only calls `generate_pptx` and offers a
  download button. Zero business logic.

Because the frontend imports only `generator.generate_pptx`, the main
application can drop Streamlit entirely and reuse the same three backend
modules unchanged.

## Image Handling

Per the brief, no images are generated. Each slide with an
`image_description` renders a dashed-border rectangle containing
`[Image] <description>` at the same coordinates a real image would occupy.

**Phase 2 hookup** — inside `pptx_builder._image_placeholder`:
1. Fetch the image binary from the non-stock image API using
   `image_description`.
2. Add the bytes to `ppt/media/imageN.<ext>` in the ZIP.
3. Register the extension in `[Content_Types].xml`
   (`<Default Extension="jpg" ContentType="image/jpeg"/>`).
4. Add a `Relationship` (Type `.../image`) in the slide's `_rels`.
5. Emit a `<p:pic>` element referencing that `rId` at the current
   `IMAGE_X / IMAGE_Y / IMAGE_W / IMAGE_H` geometry.

No layout code above `_image_placeholder` needs to change.

## Trade-offs / Notes

- **Token usage** is bounded: the LLM only outputs structured JSON (titles,
  short bullets, one-liner image descriptions). No image data flows through
  the model.
- **Single blank layout** keeps the OOXML minimal. Additional layouts
  (title/section/two-column) can be added by dropping in more
  `slideLayoutN.xml` files and updating the master's `sldLayoutIdLst`.
- **Deterministic assembly** — given the same deck JSON, the builder
  produces identical bytes (useful for caching/diffing).
- **Standard library only** for the OOXML/zip step. Non-stdlib dependencies
  are limited to the LLM client (`openai`), schema validation (`pydantic`),
  and the demo UI (`streamlit`).
