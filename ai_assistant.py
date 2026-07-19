"""
ai_assistant.py
Backend logic for the "AI Assistant" sidebar page in LinkHarvest.

A Copilot-style helper: attach 1-3 images (a screenshot of a spreadsheet,
a table, an error message, a chart, a whiteboard sketch — anything visual)
plus a plain-English prompt, and it writes the logic you asked for —
SQL, DAX, Python, an Excel formula, VBA, or anything else the prompt
requests — reasoning directly from what's in the images.

Reuses the same OpenAI/Gemini provider plumbing (constants, saved .env
keys, default models) as pdf_extractor.py so the key you already saved
there works here automatically.

No UI code lives here. Keep this importable and testable on its own.
"""
import os
import time
import base64
import mimetypes

import pdf_extractor as pe  # reuse provider constants + saved .env API keys

try:
    from openai import OpenAI
    from openai import (
        AuthenticationError as OpenAIAuthenticationError,
        APITimeoutError as OpenAIAPITimeoutError,
        RateLimitError as OpenAIRateLimitError,
        APIError as OpenAIAPIError,
    )
    OPENAI_AVAILABLE = True
except ImportError:
    OPENAI_AVAILABLE = False

try:
    import google.generativeai as genai
    from google.api_core.exceptions import (
        GoogleAPIError,
        InvalidArgument as GoogleInvalidArgument,
        PermissionDenied as GooglePermissionDenied,
        ResourceExhausted as GoogleResourceExhausted,
        DeadlineExceeded as GoogleDeadlineExceeded,
    )
    from PIL import Image
    GEMINI_AVAILABLE = True
except ImportError:
    GEMINI_AVAILABLE = False


class AuthError(Exception):
    """Raised when the selected provider's API key is missing or invalid."""


MAX_IMAGES = 3
ALLOWED_EXTENSIONS = {".png", ".jpg", ".jpeg", ".webp", ".gif", ".bmp"}
MAX_RETRIES = 3
RETRY_BACKOFF_SECONDS = 2
API_TIMEOUT = 90

SYSTEM_PROMPT = (
    "You are an expert coding copilot embedded in a data-tools app, similar to GitHub "
    "Copilot but grounded in the images the user attaches (screenshots of spreadsheets, "
    "database schemas, error messages, whiteboard sketches, charts, etc.).\n\n"
    "Rules:\n"
    "1. Look carefully at every attached image before answering — column names, table "
    "structures, sample values, and any visible errors all matter.\n"
    "2. Write the logic the user asked for in whatever language/tool they specify or "
    "imply — SQL, DAX, Python (pandas), Excel formulas, VBA, or anything else. If they "
    "don't specify, pick the most obviously appropriate one from the images and say which "
    "you chose and why in one line.\n"
    "3. Base column/table/field names on exactly what's visible in the images. If a needed "
    "name isn't visible or is ambiguous, use a clearly-marked placeholder (e.g. "
    "<table_name>) and say so — never invent a specific name that isn't shown.\n"
    "4. Structure the answer as: a one-to-two sentence explanation of the approach, then "
    "the code in a fenced code block with the correct language tag, then (only if genuinely "
    "useful) a short note on edge cases or how to adapt it.\n"
    "5. Keep prose minimal — the code is the deliverable. Never pad with generic disclaimers."
)


def is_provider_available(provider: str) -> bool:
    return OPENAI_AVAILABLE if provider == pe.PROVIDER_OPENAI else GEMINI_AVAILABLE


def validate_images(image_paths: list) -> list:
    """Filters to existing, allowed-extension image files and caps at MAX_IMAGES."""
    valid = []
    for path in image_paths:
        if not path:
            continue
        ext = os.path.splitext(path)[1].lower()
        if os.path.isfile(path) and ext in ALLOWED_EXTENSIONS:
            valid.append(path)
    return valid[:MAX_IMAGES]


# --------------------------------------------------------------------------
# OpenAI (vision via image_url data URLs)
# --------------------------------------------------------------------------
def _encode_image_data_url(path: str) -> str:
    mime, _ = mimetypes.guess_type(path)
    mime = mime or "image/png"
    with open(path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode("ascii")
    return f"data:{mime};base64,{b64}"


def _call_openai_assist(client, model_name, prompt, image_paths):
    content = [{"type": "text", "text": prompt}]
    for path in image_paths:
        content.append({"type": "image_url", "image_url": {"url": _encode_image_data_url(path)}})

    last_error = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            response = client.chat.completions.create(
                model=model_name,
                messages=[
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user", "content": content},
                ],
                temperature=0.2,
                timeout=API_TIMEOUT,
            )
            return response.choices[0].message.content.strip()
        except OpenAIAuthenticationError as e:
            raise AuthError(f"Invalid or missing OpenAI API key: {e}")
        except (OpenAIAPITimeoutError, OpenAIRateLimitError) as e:
            last_error = e
            if attempt < MAX_RETRIES:
                time.sleep(RETRY_BACKOFF_SECONDS * (2 ** (attempt - 1)))
                continue
            raise RuntimeError(f"OpenAI timed out/rate-limited after {MAX_RETRIES} attempts: {e}")
        except OpenAIAPIError as e:
            raise RuntimeError(f"OpenAI API error (check the model name '{model_name}' supports images): {e}")
    raise RuntimeError(f"Request failed: {last_error}")


# --------------------------------------------------------------------------
# Gemini (vision via inline PIL images)
# --------------------------------------------------------------------------
def _call_gemini_assist(model, prompt, image_paths):
    parts = [prompt]
    for path in image_paths:
        parts.append(Image.open(path))

    last_error = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            response = model.generate_content(
                parts,
                generation_config=genai.types.GenerationConfig(temperature=0.2),
                request_options={"timeout": API_TIMEOUT},
            )
            return response.text.strip()
        except GooglePermissionDenied as e:
            raise AuthError(f"Invalid or missing Gemini API key: {e}")
        except GoogleInvalidArgument as e:
            msg = str(e).lower()
            if "api key" in msg or "api_key" in msg:
                raise AuthError(f"Invalid Gemini API key: {e}")
            raise RuntimeError(f"Gemini rejected the request: {e}")
        except (GoogleResourceExhausted, GoogleDeadlineExceeded) as e:
            last_error = e
            if attempt < MAX_RETRIES:
                time.sleep(RETRY_BACKOFF_SECONDS * (2 ** (attempt - 1)))
                continue
            raise RuntimeError(f"Gemini timed out/rate-limited after {MAX_RETRIES} attempts: {e}")
        except GoogleAPIError as e:
            raise RuntimeError(f"Gemini API error: {e}")
    raise RuntimeError(f"Request failed: {last_error}")


# --------------------------------------------------------------------------
# Public entry point
# --------------------------------------------------------------------------
def ask_assistant(prompt: str, image_paths: list, provider: str, api_key: str, model_name: str = None) -> str:
    """
    prompt: the user's instruction, e.g. "Write the DAX measure for this".
    image_paths: 0-3 paths to screenshots/images providing context.
    Returns the assistant's markdown-formatted response (explanation + code block).
    Raises AuthError on a missing/invalid key, RuntimeError on other failures.
    """
    if provider not in (pe.PROVIDER_OPENAI, pe.PROVIDER_GEMINI):
        raise ValueError(f"Unknown provider: {provider}")
    if not is_provider_available(provider):
        pkg = "openai" if provider == pe.PROVIDER_OPENAI else "google-generativeai"
        raise RuntimeError(f"The '{pkg}' package isn't installed. Run: pip install {pkg}")
    if not api_key:
        raise AuthError(f"No {provider.title()} API key configured.")
    if not prompt or not prompt.strip():
        raise ValueError("Please describe what you want written.")

    images = validate_images(image_paths)

    model_name = (model_name or "").strip()
    if not model_name:
        model_name = pe.DEFAULT_OPENAI_MODEL if provider == pe.PROVIDER_OPENAI else pe.DEFAULT_GEMINI_MODEL

    if provider == pe.PROVIDER_OPENAI:
        client = OpenAI(api_key=api_key)
        return _call_openai_assist(client, model_name, prompt.strip(), images)
    else:
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel(model_name, system_instruction=SYSTEM_PROMPT)
        return _call_gemini_assist(model, prompt.strip(), images)
