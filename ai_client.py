"""
OpenRouter client for generating structured deck outlines.

Uses the OpenAI Python SDK pointed at OpenRouter's OpenAI-compatible endpoint.
The API key is read from the OPENROUTER_API_KEY environment variable.

Selected model: `google/gemini-2.0-flash-exp:free`
    - Free tier on OpenRouter.
    - Fast (flash tier), returns structured JSON reliably.
    - Strong instruction-following for schema-constrained output.

To swap models, change DEFAULT_MODEL below. Other solid free options:
    - `meta-llama/llama-3.3-70b-instruct:free`
    - `deepseek/deepseek-chat-v3.1:free`
"""

from __future__ import annotations

import json
import os
from typing import List, Optional

from openai import OpenAI
from pydantic import BaseModel, Field


DEFAULT_MODEL = "deepseek/deepseek-chat-v3.1"
OPENROUTER_BASE_URL = "https://openrouter.ai/api/v1"


class SlideOutline(BaseModel):
    title: str = Field(..., description="Slide title, max 8 words.")
    bullets: List[str] = Field(..., description="3-5 short bullet points, each under 18 words.")
    image_description: Optional[str] = Field(
        None, description="Short vivid image description (under 20 words), or null if no image fits."
    )


class DeckOutline(BaseModel):
    title: str = Field(..., description="Deck title, under 10 words.")
    slides: List[SlideOutline]


def _build_system_prompt(slide_count: int) -> str:
    return (
        "You are an expert presentation designer. Given a topic, produce a clean, "
        "well-structured slide deck outline as JSON.\n\n"
        "Output shape (STRICT):\n"
        "{\n"
        '  "title": "<deck title, under 10 words>",\n'
        '  "slides": [\n'
        '    { "title": "<slide title, max 8 words>", '
        '"bullets": ["...", "..."], '
        '"image_description": "<short vivid image description, or null>" }\n'
        "  ]\n"
        "}\n\n"
        "Rules:\n"
        '- Return a single JSON object with exactly the keys "title" and "slides". Never a bare array.\n'
        f'- Produce exactly {slide_count} content slides in the "slides" array '
        "(do not include a separate title slide there).\n"
        "- Each slide has 3-5 short bullet points, each under 18 words. No markdown syntax.\n"
        '- Every slide should include an "image_description" string (under 20 words). '
        "Use null only if truly no image fits.\n"
        "- Respond with the JSON object only. No prose, no code fences."
    )


def generate_deck_outline(
    prompt: str,
    slide_count: int = 6,
    model: str = DEFAULT_MODEL,
    api_key: Optional[str] = None,
) -> DeckOutline:
    """
    Call OpenRouter and parse a structured DeckOutline.

    Raises:
        RuntimeError: if OPENROUTER_API_KEY is missing.
        ValueError: if the model response is not valid JSON matching DeckOutline.
    """
    key = api_key or os.environ.get("OPENROUTER_API_KEY")
    if not key:
        raise RuntimeError(
            "Missing OPENROUTER_API_KEY. Set it in your environment or a .env file."
        )

    client = OpenAI(base_url=OPENROUTER_BASE_URL, api_key=key)

    response = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": _build_system_prompt(slide_count)},
            {"role": "user", "content": prompt},
        ],
        response_format={"type": "json_object"},
        temperature=0.7,
    )

    raw = response.choices[0].message.content or ""
    try:
        data = json.loads(raw)
    except json.JSONDecodeError as e:
        # Some models wrap JSON in code fences despite the instruction.
        stripped = raw.strip().strip("`")
        if stripped.startswith("json"):
            stripped = stripped[4:].strip()
        try:
            data = json.loads(stripped)
        except json.JSONDecodeError:
            raise ValueError(f"Model did not return valid JSON: {raw[:400]}") from e

    return DeckOutline.model_validate(data)
