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
PROVIDER_GROK = "grok"
PROVIDER_OLLAMA = "ollama"

ALL_PROVIDERS = [PROVIDER_GEMINI, PROVIDER_OPENAI, PROVIDER_GROK, PROVIDER_OLLAMA]

# Grok (xAI) and Ollama both speak the OpenAI-compatible chat.completions
# API, so they're driven through the same `openai` client, just pointed at
# a different base_url. Ollama runs locally and needs no real API key.
GROK_BASE_URL = "https://api.x.ai/v1"
OLLAMA_DEFAULT_BASE_URL = "http://localhost:11434/v1"
OLLAMA_PLACEHOLDER_KEY = "ollama"  # the SDK requires a non-empty string even though Ollama ignores it

try:
    from openai import OpenAI
    from openai import (
        AuthenticationError as OpenAIAuthenticationError,
        APITimeoutError as OpenAIAPITimeoutError,
        RateLimitError as OpenAIRateLimitError,
        APIConnectionError as OpenAIAPIConnectionError,
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
DEFAULT_GROK_MODEL = "grok-4-fast"
DEFAULT_OLLAMA_MODEL = "llama3.2"

# Shown as dropdown suggestions in the UI, but every field is freely
# editable — all four providers rename/retire models often enough that a
# hardcoded-only list would go stale. "(free)" tags are best-effort and can
# change; always double check the provider's own pricing/docs page:
#   OpenAI  -> https://platform.openai.com/docs/models
#   Gemini  -> https://ai.google.dev/gemini-api/docs/models
#   Grok    -> https://docs.x.ai/docs/models
#   Ollama  -> https://ollama.com/library  (all models are free/local)
SUGGESTED_MODELS = {
    PROVIDER_OPENAI: ["gpt-5", "gpt-5-mini", "gpt-5-nano"],
    PROVIDER_GEMINI: [
        "gemini-flash-latest",          # rolling alias, currently Gemini 3.5 Flash (free tier)
        "gemini-flash-lite-latest",     # cheapest/fastest, free tier
        "gemini-2.5-flash",             # free tier
        "gemini-2.5-flash-lite",        # free tier
        "gemini-3-flash",               # preview, free tier w/ tighter limits
        "gemini-pro-latest",            # paid only
        "gemini-3.1-pro",               # paid only, 2M context
        "gemini-2.5-flash-image",       # "Nano Banana" — image generation, not text extraction
        "gemini-3.1-flash-image",       # "Nano Banana 2" — image generation, not text extraction
    ],
    PROVIDER_GROK: [
        "grok-4-fast",       # cheapest current text model; xAI's free/trial API tier covers this first
        "grok-4.1-fast",
        "grok-4.3",
        "grok-code-fast-1",
    ],
    PROVIDER_OLLAMA: [
        "llama3.2",
        "llama3.1",
        "mistral",
        "qwen2.5",
        "gemma2",
        "phi3",
        "llava",             # vision-capable, for scanned/image documents
    ],
}

# Providers/models that are free to use out of the box (no billing needed).
# Ollama is entirely free/local. Gemini's Flash + Flash-Lite family is free
# tier (Pro models are paid-only). OpenAI and Grok have no standing free
# tier as of this writing (Grok has a limited free/trial credit, not a
# permanently free model) — shown for completeness, not flagged "free".
FREE_MODELS = {
    PROVIDER_GEMINI: {
        "gemini-flash-latest", "gemini-flash-lite-latest",
        "gemini-2.5-flash", "gemini-2.5-flash-lite", "gemini-3-flash",
    },
    PROVIDER_OLLAMA: set(SUGGESTED_MODELS[PROVIDER_OLLAMA]),  # all of them
}


def is_free_model(provider: str, model_name: str) -> bool:
    return model_name in FREE_MODELS.get(provider, set())


MAX_RETRIES = 3
RETRY_BACKOFF_SECONDS = 2  # doubles each retry: 2s, 4s, 8s
API_TIMEOUT = 60
# Cap raw text sent per PDF to stay well inside the model's context window
# on very long documents (roughly ~40k characters ~ 10k tokens).
MAX_TEXT_CHARS = 40000

ENV_KEY_NAMES = {
    PROVIDER_OPENAI: "OPENAI_API_KEY",
    PROVIDER_GEMINI: "GEMINI_API_KEY",
    PROVIDER_GROK: "XAI_API_KEY",
    PROVIDER_OLLAMA: "OLLAMA_API_KEY",  # unused by Ollama itself, kept for interface symmetry
}


class AuthError(Exception):
    """Raised when the selected provider's API key is missing or invalid."""


def is_provider_available(provider: str) -> bool:
    """
    OpenAI, Grok, and Ollama all go through the `openai` package (Grok/Ollama
    just point it at a different base_url), so they share one availability
    check. Gemini uses the separate google-generativeai package.
    """
    if provider == PROVIDER_GEMINI:
        return GEMINI_AVAILABLE
    return OPENAI_AVAILABLE


def provider_requires_api_key(provider: str) -> bool:
    """Ollama runs locally and doesn't need a real API key."""
    return provider != PROVIDER_OLLAMA


# --------------------------------------------------------------------------
# API key handling (.env file) — one slot per provider
# --------------------------------------------------------------------------
def load_api_key(provider: str = PROVIDER_OPENAI) -> str:
    """Reads the given provider's API key from .env (if present) or the environment."""
    if load_dotenv:
        load_dotenv(ENV_PATH, override=False)
    return os.environ.get(ENV_KEY_NAMES[provider], "")


def load_ollama_base_url() -> str:
    """Reads a custom Ollama server URL from .env, falling back to localhost."""
    if load_dotenv:
        load_dotenv(ENV_PATH, override=False)
    return os.environ.get("OLLAMA_BASE_URL", OLLAMA_DEFAULT_BASE_URL)


def save_ollama_base_url(url: str) -> None:
    url = (url or "").strip() or OLLAMA_DEFAULT_BASE_URL
    if not os.path.exists(ENV_PATH):
        with open(ENV_PATH, "w") as f:
            f.write("")
    if set_key:
        set_key(ENV_PATH, "OLLAMA_BASE_URL", url)
    else:
        with open(ENV_PATH, "a") as f:
            f.write(f"\nOLLAMA_BASE_URL={url}\n")
    os.environ["OLLAMA_BASE_URL"] = url


def save_api_key(key: str, provider: str = PROVIDER_OPENAI) -> None:
    """Writes/updates the given provider's API key in the local .env file."""
    key = (key or "").strip()
    if not key:
        if provider == PROVIDER_OLLAMA:
            key = OLLAMA_PLACEHOLDER_KEY  # Ollama ignores the key but the SDK needs a non-empty string
        else:
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
        "commentary, no markdown formatting, no extra keys.\n"
        "5. PARTY-MATCHING PRECISION: documents like invoices/receipts/POs often "
        "contain multiple named parties, and different documents label them "
        "differently. Step 1 — ALWAYS search the text first for the exact word "
        "the requested field uses (e.g. a field called 'vendor name' should "
        "look for a label literally containing 'Vendor' before anything else). "
        "Step 2 — only if that exact label is genuinely absent from the "
        "document, fall back to matching by role: the party providing the "
        "goods/services and getting paid (commonly labelled seller, vendor, "
        "supplier, biller, or 'billed by') versus the party being charged "
        "(commonly labelled buyer, customer, bill-to, ship-to, or client). "
        "CRITICAL: 'vendor' and 'seller' are NOT always the same party — some "
        "documents (e.g. marketplace/dropship/consignment paperwork, multi-tier "
        "POs) name a vendor/supplier separately from a seller/reseller, each "
        "with their own address and details. Never assume two differently-"
        "labelled roles are the same party just because they're often "
        "synonyms elsewhere; if the document has separate blocks for both, "
        "extract each field from the block matching ITS OWN exact label, not "
        "from whichever similarly-themed block appears first, is largest, or "
        "is otherwise easiest to find. Never return one party's details for a "
        "field that names a different party."
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
# Batch prompts — pack several documents into ONE request to cut down on
# API calls (helps both free-tier rate limits and paid per-request/latency
# overhead). Used when the "Batch Size" slider in the UI is set above 1.
# --------------------------------------------------------------------------
def _build_batch_system_prompt() -> str:
    return _build_system_prompt() + (
        "\n\n6. BATCH MODE: you will receive MULTIPLE documents in this one request, each "
        "marked with its own '--- DOCUMENT: <label> ---' header (for images, a text "
        "marker immediately precedes each image). Treat every document completely "
        "independently — never let information from one document fill in, override, "
        "or influence the extracted fields of another. Return ONE JSON object whose "
        "top-level keys are exactly the document labels given (verbatim, unchanged), "
        "and whose value for each key is the flat {field: value|null} object for that "
        "one document, following all the rules above. Include every label given, even "
        "if a document yields no matches (all null)."
    )


# Keeps a single batched request from ballooning past the model's context
# window as batch size grows — each document gets a smaller slice, not the
# full MAX_TEXT_CHARS, when it's sharing a request with others.
BATCH_PER_DOC_CHAR_BUDGET = 8000


def _build_batch_user_prompt(instructions: str, items: list) -> str:
    """items: list of (label, text) tuples."""
    sections = [f"--- DOCUMENT: {label} ---\n{text[:BATCH_PER_DOC_CHAR_BUDGET]}" for label, text in items]
    body = "\n\n".join(sections)
    return f"Fields to extract from EACH document below: {instructions}\n\n{body}"


# --------------------------------------------------------------------------
# OpenAI structured extraction
# --------------------------------------------------------------------------
def _call_openai_chat_json(client, model_name: str, system_prompt: str, user_prompt: str,
                            force_json_mode: bool = True) -> dict:
    """
    Core retry/parse logic shared by every OpenAI-compatible call site
    (OpenAI, Grok, Ollama) — single-doc and batch alike. `force_json_mode`
    requests strict JSON-object output; some Ollama models don't support
    that flag, so the caller can retry with it off and parse best-effort.
    """
    last_error = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            kwargs = dict(
                model=model_name,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt},
                ],
                temperature=0,
                timeout=API_TIMEOUT,
            )
            if force_json_mode:
                kwargs["response_format"] = {"type": "json_object"}
            response = client.chat.completions.create(**kwargs)
            content = response.choices[0].message.content
            return _parse_json_object(_strip_json_fences(content))
        except OpenAIAuthenticationError as e:
            raise AuthError(f"Invalid or missing API key: {e}")
        except OpenAIAPIConnectionError as e:
            # Most common cause: Ollama's local server isn't running.
            raise RuntimeError(
                f"Couldn't connect to the API server ({e}). If you're using Ollama, "
                f"make sure it's running (`ollama serve`) and the server URL is correct."
            )
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
            raise RuntimeError(f"API error (check the model name '{model_name}' is valid): {e}")
    raise RuntimeError(f"Extraction failed: {last_error}")


def _call_openai_extract(client, model_name: str, instructions: str, pdf_text: str,
                          force_json_mode: bool = True) -> dict:
    """Single-document text extraction — thin wrapper around _call_openai_chat_json."""
    return _call_openai_chat_json(
        client, model_name, _build_system_prompt(), _build_user_prompt(instructions, pdf_text),
        force_json_mode=force_json_mode)


def _strip_json_fences(text: str) -> str:
    """Some models (esp. local Ollama ones without strict JSON mode) wrap
    their JSON in ```json ... ``` fences despite instructions not to."""
    text = (text or "").strip()
    if text.startswith("```"):
        text = text.split("\n", 1)[1] if "\n" in text else text
        if text.endswith("```"):
            text = text[:-3]
        if text.lower().startswith("json"):
            text = text[4:]
    return text.strip()


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
def _make_client(provider: str, api_key: str, model_name: str, base_url: str = None):
    """Builds whatever object the per-file worker needs to call the API."""
    if provider == PROVIDER_OPENAI:
        return OpenAI(api_key=api_key)
    elif provider == PROVIDER_GROK:
        return OpenAI(api_key=api_key, base_url=GROK_BASE_URL)
    elif provider == PROVIDER_OLLAMA:
        return OpenAI(api_key=api_key or OLLAMA_PLACEHOLDER_KEY,
                       base_url=base_url or OLLAMA_DEFAULT_BASE_URL)
    elif provider == PROVIDER_GEMINI:
        genai.configure(api_key=api_key)
        return genai.GenerativeModel(
            model_name,
            system_instruction=_build_system_prompt(),
        )
    raise ValueError(f"Unknown provider: {provider}")


def _call_ai_extract(provider: str, client, model_name: str, instructions: str, pdf_text: str) -> dict:
    if provider in (PROVIDER_OPENAI, PROVIDER_GROK, PROVIDER_OLLAMA):
        # All three speak the same chat.completions.create interface.
        # Ollama's JSON-mode support varies by model, so fall back to a
        # plain call + best-effort JSON parse if the strict mode errors out.
        try:
            return _call_openai_extract(client, model_name, instructions, pdf_text)
        except RuntimeError:
            if provider != PROVIDER_OLLAMA:
                raise
            return _call_openai_extract(client, model_name, instructions, pdf_text, force_json_mode=False)
    elif provider == PROVIDER_GEMINI:
        # model_name is already baked into `client` (a GenerativeModel) at
        # construction time in _make_client, so it's unused here.
        return _call_gemini_extract(client, instructions, pdf_text)
    raise ValueError(f"Unknown provider: {provider}")


def _call_ai_extract_batch_text(provider: str, client, model_name: str, instructions: str, items: list) -> dict:
    """items: list of (label, text). Returns {label: {field: value}}."""
    prompt = _build_batch_user_prompt(instructions, items)
    if provider in (PROVIDER_OPENAI, PROVIDER_GROK, PROVIDER_OLLAMA):
        try:
            return _call_openai_chat_json(client, model_name, _build_batch_system_prompt(), prompt)
        except RuntimeError:
            if provider != PROVIDER_OLLAMA:
                raise
            return _call_openai_chat_json(client, model_name, _build_batch_system_prompt(), prompt,
                                           force_json_mode=False)
    elif provider == PROVIDER_GEMINI:
        try:
            response = client.generate_content(
                prompt,
                generation_config=genai.types.GenerationConfig(response_mime_type="application/json", temperature=0),
                request_options={"timeout": API_TIMEOUT},
            )
            return _parse_json_object(response.text)
        except GooglePermissionDenied as e:
            raise AuthError(f"Invalid or missing Gemini API key: {e}")
        except GoogleAPIError as e:
            raise RuntimeError(f"Gemini batch API error: {e}")
    raise ValueError(f"Unknown provider: {provider}")


def _call_ai_extract_batch_vision(provider: str, client, model_name: str, instructions: str, items: list) -> dict:
    """items: list of (label, image_path). Returns {label: {field: value}}."""
    prompt = _build_batch_user_prompt(
        instructions, [(label, "(see attached image below — no OCR text was pre-extracted)") for label, _ in items])
    if provider == PROVIDER_GEMINI:
        parts = [prompt]
        from PIL import Image
        for label, path in items:
            parts.append(f"Image for document '{label}':")
            parts.append(Image.open(path))
        try:
            response = client.generate_content(
                parts,
                generation_config=genai.types.GenerationConfig(response_mime_type="application/json", temperature=0),
                request_options={"timeout": API_TIMEOUT},
            )
            return _parse_json_object(response.text)
        except GooglePermissionDenied as e:
            raise AuthError(f"Invalid or missing Gemini API key: {e}")
        except GoogleAPIError as e:
            raise RuntimeError(f"Gemini batch vision API error: {e}")
    else:  # OpenAI, Grok, Ollama (llava etc.) — same image_url content shape
        content = [{"type": "text", "text": prompt}]
        for label, path in items:
            content.append({"type": "text", "text": f"Image for document '{label}':"})
            content.append({"type": "image_url", "image_url": {"url": _encode_image_data_url(path)}})
        try:
            response = client.chat.completions.create(
                model=model_name,
                messages=[
                    {"role": "system", "content": _build_batch_system_prompt()},
                    {"role": "user", "content": content},
                ],
                response_format={"type": "json_object"},
                temperature=0,
                timeout=API_TIMEOUT,
            )
            return _parse_json_object(_strip_json_fences(response.choices[0].message.content))
        except OpenAIAuthenticationError as e:
            raise AuthError(f"Invalid or missing API key: {e}")
        except OpenAIAPIError as e:
            raise RuntimeError(f"Batch vision API error (check '{model_name}' supports images): {e}")


# --------------------------------------------------------------------------
# Non-PDF (image / audio / video) extraction
# --------------------------------------------------------------------------
def _encode_image_data_url(path: str) -> str:
    import mimetypes, base64
    mime, _ = mimetypes.guess_type(path)
    mime = mime or "image/png"
    with open(path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode("ascii")
    return f"data:{mime};base64,{b64}"


def _call_vision_extract_openai_compat(client, model_name: str, instructions: str, image_path: str) -> dict:
    """Vision extraction via the OpenAI-compatible chat API (OpenAI, Grok, and
    Ollama vision models like llava all accept this same image_url shape)."""
    prompt = _build_user_prompt(instructions, "(see attached image — no OCR text was pre-extracted)")
    try:
        response = client.chat.completions.create(
            model=model_name,
            messages=[
                {"role": "system", "content": _build_system_prompt()},
                {"role": "user", "content": [
                    {"type": "text", "text": prompt},
                    {"type": "image_url", "image_url": {"url": _encode_image_data_url(image_path)}},
                ]},
            ],
            response_format={"type": "json_object"},
            temperature=0,
            timeout=API_TIMEOUT,
        )
        return _parse_json_object(_strip_json_fences(response.choices[0].message.content))
    except OpenAIAuthenticationError as e:
        raise AuthError(f"Invalid or missing API key: {e}")
    except OpenAIAPIError as e:
        raise RuntimeError(f"Vision API error (check '{model_name}' supports images): {e}")


def _call_vision_extract_gemini(client, instructions: str, image_path: str) -> dict:
    from PIL import Image
    prompt = _build_user_prompt(instructions, "(see attached image — no OCR text was pre-extracted)")
    try:
        response = client.generate_content(
            [prompt, Image.open(image_path)],
            generation_config=genai.types.GenerationConfig(response_mime_type="application/json", temperature=0),
            request_options={"timeout": API_TIMEOUT},
        )
        return _parse_json_object(response.text)
    except GooglePermissionDenied as e:
        raise AuthError(f"Invalid or missing Gemini API key: {e}")
    except GoogleAPIError as e:
        raise RuntimeError(f"Gemini vision API error: {e}")


def _call_audio_extract(provider: str, client, model_name: str, instructions: str, audio_path: str) -> dict:
    if provider == PROVIDER_OPENAI:
        try:
            with open(audio_path, "rb") as f:
                transcript = client.audio.transcriptions.create(model="whisper-1", file=f).text
        except OpenAIAuthenticationError as e:
            raise AuthError(f"Invalid or missing OpenAI API key: {e}")
        except OpenAIAPIError as e:
            raise RuntimeError(f"Whisper transcription failed: {e}")
        return _call_openai_extract(client, model_name, instructions, transcript)
    elif provider == PROVIDER_GEMINI:
        try:
            uploaded = genai.upload_file(audio_path)
            prompt = _build_user_prompt(instructions, "(see attached audio file — transcribe, then extract)")
            response = client.generate_content(
                [prompt, uploaded],
                generation_config=genai.types.GenerationConfig(response_mime_type="application/json", temperature=0),
                request_options={"timeout": API_TIMEOUT},
            )
            return _parse_json_object(response.text)
        except GooglePermissionDenied as e:
            raise AuthError(f"Invalid or missing Gemini API key: {e}")
        except GoogleAPIError as e:
            raise RuntimeError(f"Gemini audio API error: {e}")
    raise RuntimeError(f"Audio files aren't supported for the '{provider}' provider yet — use OpenAI or Gemini.")


def _call_video_extract(provider: str, client, instructions: str, video_path: str) -> dict:
    if provider != PROVIDER_GEMINI:
        raise RuntimeError(f"Video files aren't supported for the '{provider}' provider yet — use Gemini.")
    try:
        uploaded = genai.upload_file(video_path)
        # Native video files need a moment to finish processing server-side.
        while getattr(uploaded, "state", None) and uploaded.state.name == "PROCESSING":
            time.sleep(2)
            uploaded = genai.get_file(uploaded.name)
        prompt = _build_user_prompt(instructions, "(see attached video file — read on-screen/spoken content, then extract)")
        response = client.generate_content(
            [prompt, uploaded],
            generation_config=genai.types.GenerationConfig(response_mime_type="application/json", temperature=0),
            request_options={"timeout": API_TIMEOUT},
        )
        return _parse_json_object(response.text)
    except GooglePermissionDenied as e:
        raise AuthError(f"Invalid or missing Gemini API key: {e}")
    except GoogleAPIError as e:
        raise RuntimeError(f"Gemini video API error: {e}")


# --------------------------------------------------------------------------
# Per-file processing + thread pool orchestration
# --------------------------------------------------------------------------
def _process_single_pdf(pdf_path: str, instructions: str, provider: str, model_name: str, client) -> dict:
    """
    Returns a result dict always containing 'Source File' plus whatever
    fields the model returned (or an '_error' key on failure). Despite the
    name (kept for backwards compatibility), this now dispatches by file
    type — PDF, image, audio, or video — based on the extension.
    """
    filename = os.path.basename(pdf_path)
    kind = file_kind(pdf_path)

    try:
        if kind == "pdf":
            text = extract_text_from_pdf(pdf_path)
            data = _call_ai_extract(provider, client, model_name, instructions, text)
        elif kind == "image":
            if provider == PROVIDER_GEMINI:
                data = _call_vision_extract_gemini(client, instructions, pdf_path)
            else:
                data = _call_vision_extract_openai_compat(client, model_name, instructions, pdf_path)
        elif kind == "audio":
            data = _call_audio_extract(provider, client, model_name, instructions, pdf_path)
        elif kind == "video":
            data = _call_video_extract(provider, client, instructions, pdf_path)
        else:
            return {"Source File": filename, "_error": f"Unsupported file type: {os.path.splitext(filename)[1]}"}
    except AuthError:
        raise  # bubble up immediately, stop the whole run
    except RuntimeError as e:
        return {"Source File": filename, "_error": str(e)}

    return _normalize_result(filename, data)


def _normalize_result(filename: str, data: dict) -> dict:
    result = {"Source File": filename}
    for k, v in data.items():
        result[k] = v if v not in (None, "null") else ""
    return result


# Only these two file kinds can currently be packed multiple-per-request.
# Audio/video go through a heavier per-file pipeline (transcription/upload)
# that doesn't batch cleanly, so they always run one file per API call
# regardless of the Batch Size slider.
BATCHABLE_KINDS = ("pdf", "image")


def _process_batch(paths: list, instructions: str, provider: str, model_name: str, client) -> list:
    """
    Processes 2+ same-kind files (all 'pdf' or all 'image') in ONE API
    request. On any failure (auth errors excepted — those bubble straight
    up) it transparently falls back to processing the batch one file at a
    time, so a single malformed document or an oversized batch can't take
    the rest of the batch down with it.
    """
    kind = file_kind(paths[0])
    labels = [os.path.basename(p) for p in paths]

    try:
        if kind == "pdf":
            items = [(os.path.basename(p), extract_text_from_pdf(p)) for p in paths]
            by_label = _call_ai_extract_batch_text(provider, client, model_name, instructions, items)
        elif kind == "image":
            items = [(os.path.basename(p), p) for p in paths]
            by_label = _call_ai_extract_batch_vision(provider, client, model_name, instructions, items)
        else:
            raise RuntimeError(f"'{kind}' isn't batchable")

        results = []
        for label in labels:
            data = by_label.get(label)
            if not isinstance(data, dict):
                results.append({"Source File": label, "_error": "Missing from batch response — retrying alone."})
            else:
                results.append(_normalize_result(label, data))
        # If the model dropped any labels entirely, retry just those individually.
        missing = [p for p, r in zip(paths, results) if "_error" in r]
        if missing:
            results = [r for r in results if "_error" not in r]
            for p in missing:
                results.append(_process_single_pdf(p, instructions, provider, model_name, client))
        return results
    except AuthError:
        raise
    except RuntimeError:
        # Whole batch failed (bad JSON, context overflow, rate limit inside
        # the combined request, etc.) — fall back to one request per file
        # rather than losing every result in the batch.
        return [_process_single_pdf(p, instructions, provider, model_name, client) for p in paths]


PROVIDER_DEFAULT_MODEL = {
    PROVIDER_OPENAI: DEFAULT_OPENAI_MODEL,
    PROVIDER_GEMINI: DEFAULT_GEMINI_MODEL,
    PROVIDER_GROK: DEFAULT_GROK_MODEL,
    PROVIDER_OLLAMA: DEFAULT_OLLAMA_MODEL,
}
_PROVIDER_DEFAULT_MODEL = PROVIDER_DEFAULT_MODEL  # backwards-compat alias for internal use
_PROVIDER_PACKAGE_HINT = {
    PROVIDER_OPENAI: "openai",
    PROVIDER_GEMINI: "google-generativeai",
    PROVIDER_GROK: "openai",  # Grok reuses the openai package via base_url
    PROVIDER_OLLAMA: "openai",  # Ollama reuses the openai package via base_url
}

MAX_BATCH_SIZE = 10  # UI slider ceiling — keeps a single request's prompt/response reasonable


def run_extraction(pdf_paths, instructions: str, provider: str, api_key: str,
                    max_workers: int, progress_cb, log_cb, model_name: str = None,
                    base_url: str = None, batch_size: int = 1):
    """
    provider: one of PROVIDER_OPENAI, PROVIDER_GEMINI, PROVIDER_GROK, PROVIDER_OLLAMA.
    model_name: which model string to send to the provider's API. If left
        blank/None, falls back to that provider's default (see
        _PROVIDER_DEFAULT_MODEL above).
    base_url: only used for PROVIDER_OLLAMA — the local server URL
        (defaults to OLLAMA_DEFAULT_BASE_URL if not given).
    batch_size: how many files to pack into a single API request (1 = old
        behavior, one request per file). Only PDF and image files batch;
        audio/video always go one-per-request regardless of this value.
        Higher values mean far fewer requests (good for free-tier daily/
        per-minute caps and for cutting paid per-request overhead), at the
        cost of a larger prompt per request and coarser progress updates.
    progress_cb(done, total) called after every completed file.
    log_cb(message) called for start/skip/error/found-partial events.
    Returns list of result dicts (one per file, in completion order).
    Raises AuthError immediately if the API key is invalid (stops early).
    """
    if provider not in ALL_PROVIDERS:
        raise ValueError(f"Unknown provider: {provider}")
    if not is_provider_available(provider):
        pkg = _PROVIDER_PACKAGE_HINT[provider]
        raise RuntimeError(f"The '{pkg}' package isn't installed. Run: pip install {pkg}")
    if provider_requires_api_key(provider) and not api_key:
        raise AuthError(f"No {provider.title()} API key configured.")

    model_name = (model_name or "").strip() or _PROVIDER_DEFAULT_MODEL[provider]
    batch_size = max(1, min(int(batch_size or 1), MAX_BATCH_SIZE))

    client = _make_client(provider, api_key, model_name, base_url=base_url)
    results = []
    total = len(pdf_paths)
    done = 0

    # Group into work units: chunks of up to `batch_size` same-kind files
    # for pdf/image, one file per unit for everything else (or whenever
    # batch_size is 1, which reproduces the exact pre-batching behavior).
    units = []
    if batch_size > 1:
        by_kind = {}
        for p in pdf_paths:
            by_kind.setdefault(file_kind(p), []).append(p)
        for kind, paths in by_kind.items():
            if kind in BATCHABLE_KINDS:
                for i in range(0, len(paths), batch_size):
                    units.append(paths[i:i + batch_size])
            else:
                units.extend([p] for p in paths)
    else:
        units = [[p] for p in pdf_paths]

    def process_unit(unit):
        if len(unit) == 1:
            return [_process_single_pdf(unit[0], instructions, provider, model_name, client)]
        return _process_batch(unit, instructions, provider, model_name, client)

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_to_unit = {executor.submit(process_unit, unit): unit for unit in units}
        for future in as_completed(future_to_unit):
            unit = future_to_unit[future]
            timestamp = time.strftime("%H:%M:%S")
            try:
                unit_results = future.result()
            except AuthError as e:
                log_cb(f"{timestamp} - AUTH ERROR: {e} — stopping.")
                executor.shutdown(wait=False, cancel_futures=True)
                raise

            for result in unit_results:
                filename = result.get("Source File", "?")
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


PDF_EXTENSIONS = (".pdf",)
IMAGE_EXTENSIONS = (".png", ".jpg", ".jpeg", ".webp", ".gif", ".bmp", ".tiff", ".tif")
AUDIO_EXTENSIONS = (".mp3", ".wav", ".m4a", ".flac", ".ogg", ".aac")
VIDEO_EXTENSIONS = (".mp4", ".mov", ".avi", ".mkv", ".webm")


def _list_by_extensions(folder_path: str, extensions: tuple) -> list:
    files = []
    for ext in extensions:
        files.extend(glob.glob(os.path.join(folder_path, f"*{ext}")))
        files.extend(glob.glob(os.path.join(folder_path, f"*{ext.upper()}")))
    return sorted(set(files))


def list_pdfs(folder_path: str):
    """Kept for backwards compatibility — PDFs only."""
    return _list_by_extensions(folder_path, PDF_EXTENSIONS)


def list_images(folder_path: str):
    return _list_by_extensions(folder_path, IMAGE_EXTENSIONS)


def list_audio(folder_path: str):
    return _list_by_extensions(folder_path, AUDIO_EXTENSIONS)


def list_video(folder_path: str):
    return _list_by_extensions(folder_path, VIDEO_EXTENSIONS)


def list_source_files(folder_path: str, include_pdf=True, include_images=False,
                       include_audio=False, include_video=False) -> list:
    """
    One combined, sorted list of every file to process from a single
    'Source Folder' — this is what backs the PDF Extractor UI's folder
    picker so the user can point it at a mixed folder (or several folders
    of different types) instead of a PDFs-only one.
    """
    files = []
    if include_pdf:
        files += list_pdfs(folder_path)
    if include_images:
        files += list_images(folder_path)
    if include_audio:
        files += list_audio(folder_path)
    if include_video:
        files += list_video(folder_path)
    return sorted(set(files))


def file_kind(path: str) -> str:
    ext = os.path.splitext(path)[1].lower()
    if ext in PDF_EXTENSIONS:
        return "pdf"
    if ext in IMAGE_EXTENSIONS:
        return "image"
    if ext in AUDIO_EXTENSIONS:
        return "audio"
    if ext in VIDEO_EXTENSIONS:
        return "video"
    return "unknown"
