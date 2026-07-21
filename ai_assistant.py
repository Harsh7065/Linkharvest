"""
ai_assistant.py
Backend logic for the "AI Assistant" sidebar page in LinkHarvest.

A Copilot-style helper: attach 1-3 images (a screenshot of a spreadsheet,
a table, an error message, a chart, a whiteboard sketch — anything visual)
plus, optionally, an actual Excel/CSV file, and a plain-English prompt.
It writes the logic you asked for — SQL, DAX, Python, an Excel formula,
VBA, or anything else the prompt requests — reasoning from what's in the
images and/or the real column names, dtypes, and a small sample of rows
read directly from the attached spreadsheet.

Reuses the same OpenAI/Gemini provider plumbing (constants, saved .env
keys, default models) as pdf_extractor.py so the key you already saved
there works here automatically.

No UI code lives here. Keep this importable and testable on its own.
"""
import os
import time
import base64
import mimetypes

import pandas as pd

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
SPREADSHEET_EXTENSIONS = {".xlsx", ".xls", ".csv"}
MAX_SPREADSHEET_SAMPLE_ROWS = 8
MAX_SPREADSHEET_SHEETS_PROFILED = 5
MAX_RETRIES = 3
RETRY_BACKOFF_SECONDS = 2
API_TIMEOUT = 90

SYSTEM_PROMPT = (
    "You are an expert coding copilot embedded in a data-tools app, similar to GitHub "
    "Copilot but grounded in the images and/or spreadsheet the user attaches (screenshots "
    "of spreadsheets, database schemas, error messages, whiteboard sketches, charts, an "
    "actual Excel/CSV file's real columns and sample rows, etc.).\n\n"
    "Rules:\n"
    "1. Look carefully at every attached image, and at any spreadsheet context provided as "
    "text, before answering — column names, dtypes, table structures, sample values, and "
    "any visible errors all matter. If both an image and a spreadsheet are attached, treat "
    "the spreadsheet's real column names/dtypes as ground truth over anything ambiguous in "
    "the image.\n"
    "2. Write the logic the user asked for in whatever language/tool they specify or "
    "imply — SQL, DAX, Python (pandas), Excel formulas, VBA, or anything else. If they "
    "don't specify, pick the most obviously appropriate one from what's attached and say "
    "which you chose and why in one line.\n"
    "3. Base column/table/field names on exactly what's visible in the images or listed in "
    "the spreadsheet context. If a needed name isn't visible or is ambiguous, use a "
    "clearly-marked placeholder (e.g. <table_name>) and say so — never invent a specific "
    "name that isn't shown.\n"
    "4. Structure the answer as: a one-to-two sentence explanation of the approach, then "
    "the code in a fenced code block with the correct language tag, then (only if genuinely "
    "useful) a short note on edge cases or how to adapt it.\n"
    "5. Keep prose minimal — the code is the deliverable. Never pad with generic disclaimers."
)


def validate_spreadsheet(file_path: str):
    """Returns file_path if it's a readable, allowed-extension spreadsheet, else None."""
    if not file_path:
        return None
    ext = os.path.splitext(file_path)[1].lower()
    if os.path.isfile(file_path) and ext in SPREADSHEET_EXTENSIONS:
        return file_path
    return None


def build_spreadsheet_context(file_path: str) -> str:
    """
    Reads a small, safe profile of the attached spreadsheet — sheet names,
    column names/dtypes, and a handful of sample rows per sheet — formatted
    as plain text to append to the prompt. Never sends the entire file.
    """
    ext = os.path.splitext(file_path)[1].lower()
    try:
        if ext == ".csv":
            sheets = {"Sheet1": pd.read_csv(file_path)}
        else:
            xl = pd.ExcelFile(file_path)
            names = xl.sheet_names[:MAX_SPREADSHEET_SHEETS_PROFILED]
            sheets = {name: xl.parse(name) for name in names}
    except Exception as e:
        return f"[Could not read attached spreadsheet '{os.path.basename(file_path)}': {e}]"

    blocks = [f"Attached spreadsheet: {os.path.basename(file_path)}"]
    for name, df in sheets.items():
        blocks.append(
            f"\nSheet '{name}' — {len(df)} rows, {len(df.columns)} columns\n"
            f"Columns (name: dtype): "
            + ", ".join(f"{c} ({df[c].dtype})" for c in df.columns)
        )
        sample = df.head(MAX_SPREADSHEET_SAMPLE_ROWS)
        if not sample.empty:
            blocks.append(f"Sample rows:\n{sample.to_csv(index=False)}")
    return "\n".join(blocks)


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
def ask_assistant(prompt: str, image_paths: list, provider: str, api_key: str,
                   model_name: str = None, spreadsheet_path: str = None) -> str:
    """
    prompt: the user's instruction, e.g. "Write the DAX measure for this".
    image_paths: 0-3 paths to screenshots/images providing context.
    spreadsheet_path: optional path to an attached .xlsx/.xls/.csv file — its
        real column names/dtypes and a small sample of rows are read (never
        the whole file) and folded into the prompt as ground-truth context.
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

    full_prompt = prompt.strip()
    sheet_path = validate_spreadsheet(spreadsheet_path) if spreadsheet_path else None
    if sheet_path:
        full_prompt = f"{full_prompt}\n\n{build_spreadsheet_context(sheet_path)}"

    model_name = (model_name or "").strip()
    if not model_name:
        model_name = pe.DEFAULT_OPENAI_MODEL if provider == pe.PROVIDER_OPENAI else pe.DEFAULT_GEMINI_MODEL

    if provider == pe.PROVIDER_OPENAI:
        client = OpenAI(api_key=api_key)
        return _call_openai_assist(client, model_name, full_prompt, images)
    else:
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel(model_name, system_instruction=SYSTEM_PROMPT)
        return _call_gemini_assist(model, full_prompt, images)
