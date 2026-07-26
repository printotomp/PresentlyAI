"""
Orchestration layer — the ONLY module the frontend needs to import.

Frontend-agnostic: takes a prompt + slide count, returns .pptx bytes and a
filename. No Streamlit, Flask, or HTTP dependency here.

Usage
-----
    from generator import generate_pptx

    pptx_bytes, filename = generate_pptx(
        prompt="A 6-slide overview of quantum computing",
        slide_count=6,
    )
    with open(filename, "wb") as f:
        f.write(pptx_bytes)
"""

from __future__ import annotations

import re
from typing import Optional, Tuple

from ai_client import generate_deck_outline
from pptx_builder import DeckInput, SlideInput, build_pptx


def _safe_filename(title: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9\-_ ]+", "", title).strip()
    cleaned = re.sub(r"\s+", "_", cleaned)
    return (cleaned[:60] or "presentation") + ".pptx"


def generate_pptx(
    prompt: str,
    slide_count: int = 6,
    model: Optional[str] = None,
    api_key: Optional[str] = None,
) -> Tuple[bytes, str]:
    """
    End-to-end pipeline: prompt -> LLM outline -> OOXML .pptx bytes.

    Returns:
        (pptx_bytes, suggested_filename)
    """
    if not prompt or not prompt.strip():
        raise ValueError("Prompt is empty.")
    slide_count = max(1, min(20, int(slide_count)))

    outline_kwargs = {"prompt": prompt, "slide_count": slide_count, "api_key": api_key}
    if model:
        outline_kwargs["model"] = model
    outline = generate_deck_outline(**outline_kwargs)

    deck = DeckInput(
        title=outline.title,
        slides=[
            SlideInput(
                title=s.title,
                bullets=list(s.bullets),
                image_description=s.image_description,
            )
            for s in outline.slides[:slide_count]
        ],
    )

    pptx_bytes = build_pptx(deck)
    return pptx_bytes, _safe_filename(deck.title)
