"""
ai_assistant.py
Backend logic for the "AI Assistant" sidebar page in LinkHarvest.

A Copilot-style helper: attach up to MAX_ATTACHMENTS (4) files of any
supported type via a single unified picker — images (screenshots of a
spreadsheet, a table, an error message, a chart, a whiteboard sketch),
an Excel/CSV file, audio, PDFs, or even a whole folder (expanded to its
supported files) — plus a plain-English prompt. It writes the logic you
asked for — SQL, DAX, Python, an Excel formula, VBA, or anything else the
prompt requests — reasoning from what's in the images and/or the real
column names, dtypes, and a small sample of rows read directly from the
attached spreadsheet, plus transcribed audio and extracted PDF text.

Reuses the same provider plumbing (constants, saved .env keys, default
models) as pdf_extractor.py so a key already saved there works here
automatically. Supports 4 providers: Gemini, OpenAI, Grok (xAI), and
Ollama (local, no API key needed).

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


MAX_IMAGES = 3  # kept for backwards compatibility; the unified picker below uses MAX_ATTACHMENTS
MAX_ATTACHMENTS = 4  # total files across all types in one request (the unified "+" picker)
IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".webp", ".gif", ".bmp", ".tiff", ".tif"}
ALLOWED_EXTENSIONS = IMAGE_EXTENSIONS  # backwards-compat alias
SPREADSHEET_EXTENSIONS = {".xlsx", ".xls", ".csv"}
AUDIO_EXTENSIONS = {".mp3", ".wav", ".m4a", ".flac", ".ogg", ".aac"}
PDF_EXTENSIONS = {".pdf"}
ATTACHMENT_EXTENSIONS = IMAGE_EXTENSIONS | SPREADSHEET_EXTENSIONS | AUDIO_EXTENSIONS | PDF_EXTENSIONS
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


def build_pdf_context(file_path: str) -> str:
    """Reads a PDF's text (via pdf_extractor's reader) to fold in as ground-truth context."""
    try:
        text = pe.extract_text_from_pdf(file_path)
    except RuntimeError as e:
        return f"[Could not read attached PDF '{os.path.basename(file_path)}': {e}]"
    return f"Attached PDF: {os.path.basename(file_path)}\n{text[:pe.MAX_TEXT_CHARS]}"


def transcribe_audio(provider: str, client, audio_path: str) -> str:
    """
    Transcribes an attached audio file to fold in as text context.
    Supported for OpenAI (Whisper) and Gemini (native audio understanding).
    Grok/Ollama don't currently have an audio path here — returns a note
    instead of raising, so the rest of the attachments still go through.
    """
    name = os.path.basename(audio_path)
    if provider == pe.PROVIDER_OPENAI:
        try:
            with open(audio_path, "rb") as f:
                text = client.audio.transcriptions.create(model="whisper-1", file=f).text
            return f"Attached audio ({name}) transcript:\n{text}"
        except OpenAIAPIError as e:
            return f"[Could not transcribe attached audio '{name}': {e}]"
    elif provider == pe.PROVIDER_GEMINI:
        try:
            uploaded = genai.upload_file(audio_path)
            resp = client.generate_content(["Transcribe this audio verbatim.", uploaded])
            return f"Attached audio ({name}) transcript:\n{resp.text}"
        except GoogleAPIError as e:
            return f"[Could not transcribe attached audio '{name}': {e}]"
    return f"[Audio attachment '{name}' skipped — transcription isn't supported for the '{provider}' provider yet.]"


def is_provider_available(provider: str) -> bool:
    if provider == pe.PROVIDER_GEMINI:
        return GEMINI_AVAILABLE
    return OPENAI_AVAILABLE  # OpenAI, Grok, and Ollama all ride the openai package


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


def classify_attachment(path: str) -> str:
    """Returns 'image' | 'spreadsheet' | 'audio' | 'pdf' | 'folder' | 'unknown'."""
    if not path:
        return "unknown"
    if os.path.isdir(path):
        return "folder"
    ext = os.path.splitext(path)[1].lower()
    if ext in IMAGE_EXTENSIONS:
        return "image"
    if ext in SPREADSHEET_EXTENSIONS:
        return "spreadsheet"
    if ext in AUDIO_EXTENSIONS:
        return "audio"
    if ext in PDF_EXTENSIONS:
        return "pdf"
    return "unknown"


def expand_attachments(paths: list) -> list:
    """
    Expands any folder entries into their contained files (non-recursive,
    only recognized extensions) and drops unreadable/unsupported paths.
    Applied before validate_attachments so a folder counts as "however many
    files it contained" against the MAX_ATTACHMENTS cap.
    """
    expanded = []
    for path in paths:
        if not path:
            continue
        if os.path.isdir(path):
            for name in sorted(os.listdir(path)):
                full = os.path.join(path, name)
                if os.path.isfile(full) and os.path.splitext(full)[1].lower() in ATTACHMENT_EXTENSIONS:
                    expanded.append(full)
        elif os.path.isfile(path):
            expanded.append(path)
    return expanded


def validate_attachments(paths: list) -> dict:
    """
    Takes up to MAX_ATTACHMENTS raw paths (files and/or folders) from the
    unified '+' picker and sorts them into the buckets ask_assistant needs.
    Only the first attached spreadsheet is used as ground-truth context
    (matches the existing single-spreadsheet-context behavior); extra
    spreadsheets beyond the first are ignored with no error.
    """
    files = expand_attachments(paths)[:MAX_ATTACHMENTS]
    buckets = {"images": [], "spreadsheet": None, "audio": [], "pdfs": []}
    for f in files:
        kind = classify_attachment(f)
        if kind == "image":
            buckets["images"].append(f)
        elif kind == "spreadsheet" and buckets["spreadsheet"] is None:
            buckets["spreadsheet"] = f
        elif kind == "audio":
            buckets["audio"].append(f)
        elif kind == "pdf":
            buckets["pdfs"].append(f)
    return buckets


# --------------------------------------------------------------------------
# OpenAI (vision via image_url data URLs)
# --------------------------------------------------------------------------
def _encode_image_data_url(path: str) -> str:
    mime, _ = mimetypes.guess_type(path)
    mime = mime or "image/png"
    with open(path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode("ascii")
    return f"data:{mime};base64,{b64}"


def _call_openai_compat_assist(client, model_name, prompt, image_paths):
    """Shared by OpenAI, Grok, and Ollama (llava etc.) — same chat.completions shape."""
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
            raise AuthError(f"Invalid or missing API key: {e}")
        except (OpenAIAPITimeoutError, OpenAIRateLimitError) as e:
            last_error = e
            if attempt < MAX_RETRIES:
                time.sleep(RETRY_BACKOFF_SECONDS * (2 ** (attempt - 1)))
                continue
            raise RuntimeError(f"API timed out/rate-limited after {MAX_RETRIES} attempts: {e}")
        except OpenAIAPIError as e:
            raise RuntimeError(f"API error (check the model name '{model_name}' supports images): {e}")
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
def _make_client(provider: str, api_key: str, model_name: str = None, base_url: str = None):
    if provider == pe.PROVIDER_OPENAI:
        return OpenAI(api_key=api_key)
    elif provider == pe.PROVIDER_GROK:
        return OpenAI(api_key=api_key, base_url=pe.GROK_BASE_URL)
    elif provider == pe.PROVIDER_OLLAMA:
        return OpenAI(api_key=api_key or pe.OLLAMA_PLACEHOLDER_KEY, base_url=base_url or pe.OLLAMA_DEFAULT_BASE_URL)
    elif provider == pe.PROVIDER_GEMINI:
        genai.configure(api_key=api_key)
        return genai.GenerativeModel(model_name or pe.DEFAULT_GEMINI_MODEL, system_instruction=SYSTEM_PROMPT)
    raise ValueError(f"Unknown provider: {provider}")


def ask_assistant(prompt: str, attachment_paths: list, provider: str, api_key: str,
                   model_name: str = None, base_url: str = None,
                   image_paths: list = None, spreadsheet_path: str = None) -> str:
    """
    prompt: the user's instruction, e.g. "Write the DAX measure for this".
    attachment_paths: 0-MAX_ATTACHMENTS paths from the unified '+' picker —
        any mix of images, one spreadsheet, audio files, PDFs, and/or
        folders (folders are expanded to their contained supported files).
    provider: one of pe.PROVIDER_OPENAI/GEMINI/GROK/OLLAMA.
    base_url: only used for PROVIDER_OLLAMA — the local server URL.
    image_paths / spreadsheet_path: deprecated, kept for backwards
        compatibility — merged into attachment_paths if attachment_paths
        is empty and these are given instead.
    Returns the assistant's markdown-formatted response (explanation + code block).
    Raises AuthError on a missing/invalid key, RuntimeError on other failures.
    """
    if provider not in pe.ALL_PROVIDERS:
        raise ValueError(f"Unknown provider: {provider}")
    if not is_provider_available(provider):
        pkg = "google-generativeai" if provider == pe.PROVIDER_GEMINI else "openai"
        raise RuntimeError(f"The '{pkg}' package isn't installed. Run: pip install {pkg}")
    if pe.provider_requires_api_key(provider) and not api_key:
        raise AuthError(f"No {provider.title()} API key configured.")
    if not prompt or not prompt.strip():
        raise ValueError("Please describe what you want written.")

    # Backwards-compat: callers still passing the old separate params.
    if not attachment_paths and (image_paths or spreadsheet_path):
        attachment_paths = list(image_paths or []) + ([spreadsheet_path] if spreadsheet_path else [])

    buckets = validate_attachments(attachment_paths or [])
    images = buckets["images"]

    model_name = (model_name or "").strip()
    if not model_name:
        model_name = {
            pe.PROVIDER_OPENAI: pe.DEFAULT_OPENAI_MODEL,
            pe.PROVIDER_GEMINI: pe.DEFAULT_GEMINI_MODEL,
            pe.PROVIDER_GROK: pe.DEFAULT_GROK_MODEL,
            pe.PROVIDER_OLLAMA: pe.DEFAULT_OLLAMA_MODEL,
        }[provider]

    client = _make_client(provider, api_key, model_name=model_name, base_url=base_url)

    full_prompt = prompt.strip()
    if buckets["spreadsheet"]:
        full_prompt = f"{full_prompt}\n\n{build_spreadsheet_context(buckets['spreadsheet'])}"
    for pdf_path in buckets["pdfs"]:
        full_prompt = f"{full_prompt}\n\n{build_pdf_context(pdf_path)}"
    for audio_path in buckets["audio"]:
        full_prompt = f"{full_prompt}\n\n{transcribe_audio(provider, client, audio_path)}"

    if provider == pe.PROVIDER_GEMINI:
        return _call_gemini_assist(client, full_prompt, images)
    else:
        return _call_openai_compat_assist(client, model_name, full_prompt, images)
