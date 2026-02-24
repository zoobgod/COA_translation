"""
OpenAI-based translation module for pharmaceutical COA documents.

Translates English text to Russian with pharmaceutical-specific context
and glossary enforcement.
"""

import logging
from typing import Optional

from openai import OpenAI

from modules.glossary import get_glossary_prompt_section

logger = logging.getLogger(__name__)

# Maximum characters per translation chunk (to stay within token limits)
MAX_CHUNK_SIZE = 6000

SYSTEM_PROMPT = """You are a professional pharmaceutical translator specializing in translating \
Certificate of Analysis (COA) documents and related pharmaceutical documentation from English to Russian.

You must follow these rules strictly:
1. Translate ALL text from English to Russian.
2. Preserve the exact document structure, formatting, and layout of the original.
3. Keep numerical values, chemical formulas, CAS numbers, and catalog numbers unchanged.
4. Keep Latin scientific names (e.g., species names) in their original Latin form.
5. Use the pharmaceutical glossary provided below for standard terminology. These translations \
are mandatory — always prefer the glossary term over a generic translation.
6. Maintain standard Russian pharmaceutical and regulatory terminology consistent with the \
Russian Pharmacopoeia (Государственная Фармакопея) conventions.
7. Preserve any table structure using the same delimiters (e.g., | for columns).
8. Do not add explanations, comments, or notes — output ONLY the translated text.
9. If a term has no glossary entry, translate it using standard pharmaceutical Russian terminology.
10. Keep abbreviations that are internationally recognized (pH, HPLC, GC, etc.) but add \
their Russian equivalents where they appear in the glossary.

PHARMACEUTICAL GLOSSARY (English → Russian):
{glossary}
"""


def _build_system_prompt() -> str:
    """Build the system prompt with the pharmaceutical glossary."""
    glossary_text = get_glossary_prompt_section()
    return SYSTEM_PROMPT.format(glossary=glossary_text)


def _chunk_text(text: str, max_size: int = MAX_CHUNK_SIZE) -> list[str]:
    """
    Split text into chunks for translation, respecting paragraph boundaries.
    """
    if len(text) <= max_size:
        return [text]

    chunks = []
    paragraphs = text.split("\n\n")
    current_chunk = ""

    for para in paragraphs:
        if len(current_chunk) + len(para) + 2 > max_size:
            if current_chunk:
                chunks.append(current_chunk.strip())
                current_chunk = ""
            # If a single paragraph is longer than max_size, split by lines
            if len(para) > max_size:
                lines = para.split("\n")
                for line in lines:
                    if len(current_chunk) + len(line) + 1 > max_size:
                        if current_chunk:
                            chunks.append(current_chunk.strip())
                        current_chunk = line + "\n"
                    else:
                        current_chunk += line + "\n"
            else:
                current_chunk = para + "\n\n"
        else:
            current_chunk += para + "\n\n"

    if current_chunk.strip():
        chunks.append(current_chunk.strip())

    return chunks


def translate_text(
    text: str,
    api_key: str,
    model: str = "gpt-4o",
    progress_callback: Optional[callable] = None,
) -> dict:
    """
    Translate pharmaceutical COA text from English to Russian using OpenAI.

    Args:
        text: The English text to translate
        api_key: OpenAI API key
        model: OpenAI model to use (default: gpt-4o)
        progress_callback: Optional callback(current_chunk, total_chunks) for progress

    Returns:
        dict with keys:
            - 'translated_text': the Russian translation
            - 'success': whether translation succeeded
            - 'error': error message if failed
            - 'model_used': which model was used
            - 'chunks_translated': number of chunks processed
    """
    if not text.strip():
        return {
            "translated_text": "",
            "success": False,
            "error": "No text provided for translation",
            "model_used": model,
            "chunks_translated": 0,
        }

    try:
        client = OpenAI(api_key=api_key)
        system_prompt = _build_system_prompt()
        chunks = _chunk_text(text)
        translated_parts = []

        for i, chunk in enumerate(chunks):
            if progress_callback:
                progress_callback(i + 1, len(chunks))

            user_message = (
                f"Translate the following pharmaceutical COA text from English to Russian. "
                f"Output ONLY the translation, nothing else.\n\n{chunk}"
            )

            response = client.chat.completions.create(
                model=model,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_message},
                ],
                temperature=0.1,  # Low temperature for consistent translations
            )

            translated = response.choices[0].message.content
            if translated:
                translated_parts.append(translated.strip())

        full_translation = "\n\n".join(translated_parts)

        return {
            "translated_text": full_translation,
            "success": True,
            "error": None,
            "model_used": model,
            "chunks_translated": len(chunks),
        }

    except Exception as e:
        logger.error(f"Translation failed: {e}")
        return {
            "translated_text": "",
            "success": False,
            "error": str(e),
            "model_used": model,
            "chunks_translated": 0,
        }
