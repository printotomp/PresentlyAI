"""
Streamlit demo frontend.

Deliberately thin — the only thing this file does is:
    1. Collect a prompt and slide count.
    2. Call `generator.generate_pptx(...)`.
    3. Expose the resulting .pptx via a download button.

All business logic lives in `generator.py`, `ai_client.py`, and `pptx_builder.py`,
so those modules can be imported directly by the main application later.

Run:
    export OPENROUTER_API_KEY=sk-or-...
    uv run streamlit run app.py
"""

from __future__ import annotations

import streamlit as st

from generator import generate_pptx


st.set_page_config(page_title="AI PowerPoint Generator — POC", page_icon="📊")

st.title("AI PowerPoint Generator")
st.caption(
    "Describe a topic. An LLM drafts the deck outline; the backend assembles a "
    "valid, editable .pptx from raw OOXML."
)

with st.form("generate"):
    prompt = st.text_area(
        "Description",
        placeholder="e.g. A 6-slide overview of quantum computing for a general business audience",
        height=150,
    )
    slide_count = st.slider("Number of slides", min_value=1, max_value=15, value=6)
    submitted = st.form_submit_button("Generate .pptx")

if submitted:
    if not prompt.strip():
        st.error("Please enter a description.")
    else:
        with st.spinner("Calling the LLM and assembling OOXML…"):
            try:
                pptx_bytes, filename = generate_pptx(prompt=prompt, slide_count=slide_count)
            except Exception as e:  # noqa: BLE001 — surface any failure to the user
                st.error(f"Generation failed: {e}")
            else:
                st.success(f"Deck ready — {len(pptx_bytes):,} bytes")
                st.download_button(
                    label="Download .pptx",
                    data=pptx_bytes,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                )

with st.expander("How this works"):
    st.markdown(
        """
- **LLM:** `google/gemini-2.0-flash-exp:free` via **OpenRouter** (OpenAI-compatible API).
- The model returns a strict JSON outline (title, slides, bullets, image descriptions).
- The backend writes raw OOXML parts (`[Content_Types].xml`, `_rels/.rels`,
  `presentation.xml`, theme, master, layout, per-slide XML) and zips them
  into a `.pptx` using only the Python standard library.
- **No `python-pptx`** or other PowerPoint libraries are used for assembly.
- Images are not generated — each slide reserves a dashed placeholder holding
  the description, ready for Phase 2 (fetch from a non-stock image API).
        """
    )
