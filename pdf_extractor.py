"""
pdf_extractor.py

Reads a folder of PDFs, sends each one's text plus a natural-language
instruction to either OpenAI or Google Gemini (user's choice), gets
back structured JSON, and compiles everything into an Excel sheet.

Design goals (per spec):
- Accuracy over guessing: the model is instructed to return null/blank
  for anything not explicitly present, never to infer or approximate.
- Concurrency via ThreadPoolExecutor (1-20 threads, default 10).
- Resilient to encrypted/corrupted PDFs, API timeouts, and rate limits.
- Provider-agnostic: OpenAI and Gemini are both supported behind one
  interface (run_extraction), selected by the `provider` argument.
"""
import os
import json
import time
import glob
from concurrent.futures import ThreadPoolExecutor, as_completed

import pdfplumber
import pandas as pd

PROVIDER_OPENAI = "openai"
PROVIDER_GEMINI = "gemini"

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
        Cancelled as GoogleCancelled,
        InvalidArgument as GoogleInvalidArgument,
        PermissionDenied as GooglePermissionDenied,
        ResourceExhausted as GoogleResourceExhausted,
        DeadlineExceeded as GoogleDeadlineExceeded,
    )
    GEMINI_AVAILABLE = True
except ImportError:
    GEMINI_AVAILABLE = False

try:
    from dotenv import load_dotenv, set_key
except ImportError:
    load_dotenv = None
    set_key = None

ENV_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")

DEFAULT_OPENAI_MODEL = "gpt-5"
# "-latest" lets Google auto-roll this without code changes; it currently
# resolves to Gemini 3.5 Flash.
DEFAULT_GEMINI_MODEL = "gemini-flash-latest"

# Shown as dropdown suggestions in the UI, but the field is freely editable —
# both providers rename/retire models often enough that a hardcoded-only
# list would go stale. Check https://platform.openai.com/docs/models or
# https://ai.google.dev/gemini-api/docs/models for the current lineup.
SUGGESTED_MODELS = {
    PROVIDER_OPENAI: ["gpt-5", "gpt-5-mini", "gpt-5-nano"],
    PROVIDER_GEMINI: ["gemini-flash-latest", "gemini-flash-lite-latest", "gemini-pro-latest"],
}

MAX_RETRIES = 3
RETRY_BACKOFF_SECONDS = 2  # doubles each retry: 2s, 4s, 8s
API_TIMEOUT = 60
# Cap raw text sent per PDF to stay well inside the model's context window
# on very long documents (roughly ~40k characters ~ 10k tokens).
MAX_TEXT_CHARS = 40000

ENV_KEY_NAMES = {
    PROVIDER_OPENAI: "OPENAI_API_KEY",
    PROVIDER_GEMINI: "GEMINI_API_KEY",
}


class AuthError(Exception):
    """Raised when the selected provider's API key is missing or invalid."""


def is_provider_available(provider: str) -> bool:
    return OPENAI_AVAILABLE if provider == PROVIDER_OPENAI else GEMINI_AVAILABLE


# --------------------------------------------------------------------------
# API key handling (.env file) — one slot per provider
# --------------------------------------------------------------------------
def load_api_key(provider: str = PROVIDER_OPENAI) -> str:
    """Reads the given provider's API key from .env (if present) or the environment."""
    if load_dotenv:
        load_dotenv(ENV_PATH, override=False)
    return os.environ.get(ENV_KEY_NAMES[provider], "")


def save_api_key(key: str, provider: str = PROVIDER_OPENAI) -> None:
    """Writes/updates the given provider's API key in the local .env file."""
    key = (key or "").strip()
    if not key:
        raise ValueError("API key is empty.")
    env_name = ENV_KEY_NAMES[provider]
    if not os.path.exists(ENV_PATH):
        with open(ENV_PATH, "w") as f:
            f.write("")
    if set_key:
        set_key(ENV_PATH, env_name, key)
    else:
        # Fallback if python-dotenv isn't installed for some reason.
        with open(ENV_PATH, "a") as f:
            f.write(f"\n{env_name}={key}\n")
    os.environ[env_name] = key


# --------------------------------------------------------------------------
# PDF text extraction
# --------------------------------------------------------------------------
def extract_text_from_pdf(pdf_path: str) -> str:
    """
    Extracts raw text from every page of a PDF using pdfplumber.
    Raises a descriptive exception on encrypted/corrupted/unreadable files
    so the caller can log a clear, per-file error instead of crashing.
    """
    try:
        text_parts = []
        with pdfplumber.open(pdf_path) as pdf:
            if getattr(pdf, "is_encrypted", False):
                raise ValueError("PDF is password-protected/encrypted")
            for page in pdf.pages:
                page_text = page.extract_text() or ""
                if page_text:
                    text_parts.append(page_text)
        text = "\n".join(text_parts).strip()
        if not text:
            raise ValueError("No extractable text found (possibly a scanned/image-only PDF)")
        return text[:MAX_TEXT_CHARS]
    except Exception as e:
        reason = str(e).strip() or (
            "file could not be opened — it may be password-protected/encrypted "
            "or corrupted"
        )
        raise RuntimeError(f"Could not read PDF: {reason}")


# --------------------------------------------------------------------------
# Shared prompt construction (identical instructions regardless of provider,
# so behavior is consistent no matter which one the user picks)
# --------------------------------------------------------------------------
def _build_system_prompt() -> str:
    return (
        "You are a precision data extraction analyst. You will be given the raw "
        "text of a single document and a list of fields the user wants extracted. "
        "Rules you must follow exactly:\n"
        "1. Only return information that is explicitly and unambiguously present "
        "in the provided text.\n"
        "2. Never guess, infer, approximate, or generate plausible-sounding values.\n"
        "3. If a requested field cannot be found in the text, its value MUST be "
        "null (JSON null), never an empty guess or a placeholder like 'N/A'.\n"
        "4. Return ONLY a single flat JSON object mapping each requested field "
        "name to its extracted value (string) or null. No nested objects, no "
        "commentary, no markdown formatting, no extra keys."
    )


def _build_user_prompt(instructions: str, pdf_text: str) -> str:
    return (
        f"Fields to extract (as described by the user):\n{instructions}\n\n"
        f"--- DOCUMENT TEXT START ---\n{pdf_text}\n--- DOCUMENT TEXT END ---\n\n"
        "Return a single flat JSON object with one key per requested field."
    )


def _parse_json_object(raw_text: str) -> dict:
    data = json.loads(raw_text)
    if not isinstance(data, dict):
        raise ValueError("Model did not return a JSON object")
    return data


# --------------------------------------------------------------------------
# OpenAI structured extraction
# --------------------------------------------------------------------------
def _call_openai_extract(client, model_name: str, instructions: str, pdf_text: str) -> dict:
    last_error = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            response = client.chat.completions.create(
                model=model_name,
                messages=[
                    {"role": "system", "content": _build_system_prompt()},
                    {"role": "user", "content": _build_user_prompt(instructions, pdf_text)},
                ],
                response_format={"type": "json_object"},
                temperature=0,
                timeout=API_TIMEOUT,
            )
            return _parse_json_object(response.choices[0].message.content)
        except OpenAIAuthenticationError as e:
            raise AuthError(f"Invalid or missing OpenAI API key: {e}")
        except (OpenAIAPITimeoutError, OpenAIRateLimitError) as e:
            last_error = e
            if attempt < MAX_RETRIES:
                time.sleep(RETRY_BACKOFF_SECONDS * (2 ** (attempt - 1)))
                continue
            raise RuntimeError(f"API timed out/rate-limited after {MAX_RETRIES} attempts: {e}")
        except json.JSONDecodeError as e:
            last_error = e
            if attempt < MAX_RETRIES:
                continue
            raise RuntimeError(f"Model returned invalid JSON after {MAX_RETRIES} attempts: {e}")
        except OpenAIAPIError as e:
            # Includes "model not found" style errors if a bad/retired model
            # string was typed into the model field — surface it clearly.
            raise RuntimeError(f"OpenAI API error (check the model name '{model_name}' is valid): {e}")
    raise RuntimeError(f"Extraction failed: {last_error}")


# --------------------------------------------------------------------------
# Gemini structured extraction
# --------------------------------------------------------------------------
def _call_gemini_extract(model, instructions: str, pdf_text: str) -> dict:
    """
    `model` is a genai.GenerativeModel already configured with the system
    instruction. Gemini's JSON mode is requested via generation_config's
    response_mime_type, mirroring OpenAI's response_format json_object.
    """
    last_error = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            response = model.generate_content(
                _build_user_prompt(instructions, pdf_text),
                generation_config=genai.types.GenerationConfig(
                    response_mime_type="application/json",
                    temperature=0,
                ),
                request_options={"timeout": API_TIMEOUT},
            )
            return _parse_json_object(response.text)
        except GooglePermissionDenied as e:
            raise AuthError(f"Invalid or missing Gemini API key: {e}")
        except GoogleInvalidArgument as e:
            # Usually a malformed request rather than a bad key, but an
            # invalid/malformed API key can also surface here depending on
            # SDK version — treat it as auth if the message hints at a key
            # problem, otherwise treat as a per-file failure.
            msg = str(e).lower()
            if "api key" in msg or "api_key" in msg:
                raise AuthError(f"Invalid Gemini API key: {e}")
            raise RuntimeError(f"Gemini rejected the request: {e}")
        except (GoogleCancelled, GoogleDeadlineExceeded, GoogleResourceExhausted) as e:
            last_error = e
            if attempt < MAX_RETRIES:
                time.sleep(RETRY_BACKOFF_SECONDS * (2 ** (attempt - 1)))
                continue
            raise RuntimeError(f"Gemini timed out/rate-limited after {MAX_RETRIES} attempts: {e}")
        except json.JSONDecodeError as e:
            last_error = e
            if attempt < MAX_RETRIES:
                continue
            raise RuntimeError(f"Model returned invalid JSON after {MAX_RETRIES} attempts: {e}")
        except GoogleAPIError as e:
            last_error = e
            if attempt < MAX_RETRIES:
                time.sleep(RETRY_BACKOFF_SECONDS * (2 ** (attempt - 1)))
                continue
            raise RuntimeError(f"Gemini API error after {MAX_RETRIES} attempts: {e}")
    raise RuntimeError(f"Extraction failed: {last_error}")


# --------------------------------------------------------------------------
# Provider dispatch
# --------------------------------------------------------------------------
def _make_client(provider: str, api_key: str, model_name: str):
    """Builds whatever object the per-file worker needs to call the API."""
    if provider == PROVIDER_OPENAI:
        return OpenAI(api_key=api_key)
    elif provider == PROVIDER_GEMINI:
        genai.configure(api_key=api_key)
        return genai.GenerativeModel(
            model_name,
            system_instruction=_build_system_prompt(),
        )
    raise ValueError(f"Unknown provider: {provider}")


def _call_ai_extract(provider: str, client, model_name: str, instructions: str, pdf_text: str) -> dict:
    if provider == PROVIDER_OPENAI:
        return _call_openai_extract(client, model_name, instructions, pdf_text)
    elif provider == PROVIDER_GEMINI:
        # model_name is already baked into `client` (a GenerativeModel) at
        # construction time in _make_client, so it's unused here.
        return _call_gemini_extract(client, instructions, pdf_text)
    raise ValueError(f"Unknown provider: {provider}")


# --------------------------------------------------------------------------
# Per-file processing + thread pool orchestration
# --------------------------------------------------------------------------
def _process_single_pdf(pdf_path: str, instructions: str, provider: str, model_name: str, client) -> dict:
    """
    Returns a result dict always containing 'Source File' plus whatever
    fields the model returned (or an '_error' key on failure).
    """
    filename = os.path.basename(pdf_path)
    try:
        text = extract_text_from_pdf(pdf_path)
    except RuntimeError as e:
        return {"Source File": filename, "_error": str(e)}

    try:
        data = _call_ai_extract(provider, client, model_name, instructions, text)
    except AuthError:
        raise  # bubble up immediately, stop the whole run
    except RuntimeError as e:
        return {"Source File": filename, "_error": str(e)}

    result = {"Source File": filename}
    for k, v in data.items():
        result[k] = v if v not in (None, "null") else ""
    return result


def run_extraction(pdf_paths, instructions: str, provider: str, api_key: str,
                    max_workers: int, progress_cb, log_cb, model_name: str = None):
    """
    provider: PROVIDER_OPENAI or PROVIDER_GEMINI.
    model_name: which model string to send to the provider's API. If left
        blank/None, falls back to that provider's default (see
        DEFAULT_OPENAI_MODEL / DEFAULT_GEMINI_MODEL above).
    progress_cb(done, total) called after every completed file.
    log_cb(message) called for start/skip/error/found-partial events.
    Returns list of result dicts (one per PDF, in completion order).
    Raises AuthError immediately if the API key is invalid (stops early).
    """
    if provider not in (PROVIDER_OPENAI, PROVIDER_GEMINI):
        raise ValueError(f"Unknown provider: {provider}")
    if not is_provider_available(provider):
        pkg = "openai" if provider == PROVIDER_OPENAI else "google-generativeai"
        raise RuntimeError(f"The '{pkg}' package isn't installed. Run: pip install {pkg}")
    if not api_key:
        raise AuthError(f"No {provider.title()} API key configured.")

    model_name = (model_name or "").strip()
    if not model_name:
        model_name = DEFAULT_OPENAI_MODEL if provider == PROVIDER_OPENAI else DEFAULT_GEMINI_MODEL

    client = _make_client(provider, api_key, model_name)
    results = []
    total = len(pdf_paths)
    done = 0

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_to_path = {
            executor.submit(_process_single_pdf, path, instructions, provider, model_name, client): path
            for path in pdf_paths
        }
        for future in as_completed(future_to_path):
            path = future_to_path[future]
            filename = os.path.basename(path)
            timestamp = time.strftime("%H:%M:%S")
            try:
                result = future.result()
            except AuthError as e:
                log_cb(f"{timestamp} - AUTH ERROR: {e} — stopping.")
                executor.shutdown(wait=False, cancel_futures=True)
                raise

            done += 1
            if "_error" in result:
                log_cb(f"{timestamp} - FAILED {filename}: {result['_error']}")
            else:
                missing = [k for k, v in result.items()
                           if k != "Source File" and (v is None or v == "")]
                if missing:
                    log_cb(f"{timestamp} - Processed {filename} "
                           f"(not found: {', '.join(missing)})")
                else:
                    log_cb(f"{timestamp} - Processed {filename} — all fields found")
            results.append(result)
            progress_cb(done, total)

    return results


# --------------------------------------------------------------------------
# Excel compilation (append to existing sheet if present)
# --------------------------------------------------------------------------
def compile_to_excel(results: list, excel_path: str, sheet_name: str):
    """
    Writes results to excel_path/sheet_name. If the file+sheet already
    exist, new rows are appended underneath the existing data instead of
    overwriting it. Column headers are the union of all keys seen
    (excluding the internal '_error' marker).
    """
    clean_rows = [{k: v for k, v in r.items() if k != "_error"} for r in results]
    new_df = pd.DataFrame(clean_rows)

    if os.path.exists(excel_path):
        try:
            existing_df = pd.read_excel(excel_path, sheet_name=sheet_name, engine="openpyxl")
            combined_df = pd.concat([existing_df, new_df], ignore_index=True)
        except ValueError:
            # sheet doesn't exist yet in this workbook
            combined_df = new_df
        with pd.ExcelWriter(excel_path, engine="openpyxl", mode="a",
                             if_sheet_exists="replace") as writer:
            combined_df.to_excel(writer, sheet_name=sheet_name, index=False)
    else:
        with pd.ExcelWriter(excel_path, engine="openpyxl", mode="w") as writer:
            new_df.to_excel(writer, sheet_name=sheet_name, index=False)


def list_pdfs(folder_path: str):
    return sorted(glob.glob(os.path.join(folder_path, "*.pdf")))
