"""
sheet_editor.py
Backend logic for the "Excel Editor" sidebar page in LinkHarvest.

Lets the user load a spreadsheet, describe an edit in plain English, and
have it applied to the data — with undo/redo and save.

Design choice: the AI is NEVER asked to write or execute arbitrary code.
Instead it translates the instruction into a small JSON list of operations
drawn from a fixed, safe vocabulary (rename column, filter rows, sort,
add a computed column, fill missing values, etc). Each operation is then
applied by our own deterministic pandas code in apply_operation(). This
keeps every edit predictable, undoable, and impossible to turn into
something destructive outside the spreadsheet itself.

No UI code lives here. Keep this importable and testable on its own.
"""
import os
import json
import time

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
        Cancelled as GoogleCancelled,
        InvalidArgument as GoogleInvalidArgument,
        PermissionDenied as GooglePermissionDenied,
        ResourceExhausted as GoogleResourceExhausted,
        DeadlineExceeded as GoogleDeadlineExceeded,
    )
    GEMINI_AVAILABLE = True
except ImportError:
    GEMINI_AVAILABLE = False


class AuthError(Exception):
    """Raised when the selected provider's API key is missing or invalid."""


class OperationError(Exception):
    """Raised when an operation can't be applied (bad column name, bad value, etc)."""


MAX_RETRIES = 3
RETRY_BACKOFF_SECONDS = 2
API_TIMEOUT = 60
MAX_UNDO_DEPTH = 25
MAX_SAMPLE_VALUES = 5      # sample values per column sent to the AI (profile only, never raw rows)
MAX_PREVIEW_ROWS = 300     # rows shown in the on-screen table at once


# ==========================================================================
# Loading / saving workbooks (multi-sheet aware)
# ==========================================================================
def load_workbook(filepath: str):
    """
    Returns (sheets, sheet_order):
        sheets: dict[sheet_name -> DataFrame]
        sheet_order: list of sheet names in their original order
    CSV files are treated as a single sheet named 'Sheet1'.
    """
    if not os.path.exists(filepath):
        raise ValueError(f"Could not read file: file not found: {filepath}")

    ext = os.path.splitext(filepath)[1].lower()
    try:
        if ext == ".csv":
            df = pd.read_csv(filepath)
            return {"Sheet1": df}, ["Sheet1"]
        elif ext in (".xlsx", ".xls"):
            xl = pd.ExcelFile(filepath)
            sheets = {name: xl.parse(name) for name in xl.sheet_names}
            return sheets, list(xl.sheet_names)
        else:
            raise ValueError(
                f"Could not read file: unsupported file type '{ext}'. "
                "Please provide a .csv, .xlsx, or .xls file."
            )
    except ValueError:
        raise
    except Exception as e:
        raise ValueError(f"Could not read file: {e}")


def save_workbook(sheets: dict, filepath: str):
    """Saves all sheets to filepath. CSV can only hold a single sheet."""
    ext = os.path.splitext(filepath)[1].lower()
    try:
        if ext == ".csv":
            if len(sheets) > 1:
                raise ValueError(
                    "This file has more than one sheet — save as .xlsx to keep "
                    "them all, or switch to the sheet you want and save that one as .csv."
                )
            df = next(iter(sheets.values()))
            df.to_csv(filepath, index=False)
        elif ext in (".xlsx", ".xls"):
            with pd.ExcelWriter(filepath, engine="openpyxl") as writer:
                for name, df in sheets.items():
                    writer_name = str(name)[:31]  # Excel's sheet-name length limit
                    df.to_excel(writer, sheet_name=writer_name, index=False)
        else:
            raise ValueError(f"Unsupported export type '{ext}'. Use .csv or .xlsx.")
    except ValueError:
        raise
    except Exception as e:
        raise ValueError(f"Could not save file: {e}")


# ==========================================================================
# Column profile sent to the AI (small, structural only — never raw rows)
# ==========================================================================
def build_column_profile(df: pd.DataFrame) -> list:
    profile = []
    for col in df.columns:
        series = df[col]
        sample_vals = series.dropna().astype(str).unique()[:MAX_SAMPLE_VALUES].tolist()
        profile.append({
            "column": str(col),
            "dtype": str(series.dtype),
            "missing": int(series.isna().sum()),
            "unique_count": int(series.nunique(dropna=True)),
            "sample_values": sample_vals,
        })
    return profile


# ==========================================================================
# Safe operation vocabulary
# ==========================================================================
_CASE_FUNCS = {"upper": str.upper, "lower": str.lower, "title": str.title}


def _get_col(df: pd.DataFrame, name):
    if name not in df.columns:
        raise OperationError(
            f"Column '{name}' not found. Available columns: {', '.join(map(str, df.columns))}"
        )
    return name


def _filter_mask(df: pd.DataFrame, column, operator, value):
    col = _get_col(df, column)
    series = df[col]

    if operator == "is_null":
        return series.isna()
    if operator == "not_null":
        return series.notna()

    if operator in ("==", "!=", ">", "<", ">=", "<="):
        numeric = pd.to_numeric(series, errors="coerce")
        try:
            num_value = float(value)
            is_numeric_comparable = numeric.notna().any()
        except (TypeError, ValueError):
            is_numeric_comparable = False

        if is_numeric_comparable:
            if operator == "==":
                return numeric == num_value
            if operator == "!=":
                return numeric != num_value
            if operator == ">":
                return numeric > num_value
            if operator == "<":
                return numeric < num_value
            if operator == ">=":
                return numeric >= num_value
            if operator == "<=":
                return numeric <= num_value

        if operator not in ("==", "!="):
            raise OperationError(
                f"Operator '{operator}' needs a numeric column; '{column}' isn't numeric."
            )
        str_series = series.astype(str)
        str_value = str(value)
        return str_series == str_value if operator == "==" else str_series != str_value

    if operator == "contains":
        return series.astype(str).str.contains(str(value), case=False, na=False)
    if operator == "not_contains":
        return ~series.astype(str).str.contains(str(value), case=False, na=False)

    raise OperationError(f"Unknown filter operator '{operator}'.")


def apply_operation(df: pd.DataFrame, op: dict) -> pd.DataFrame:
    """Applies a single operation dict to df. Returns a NEW DataFrame; df is untouched."""
    df = df.copy()
    kind = op.get("op")

    if kind == "rename_column":
        col = _get_col(df, op["column"])
        return df.rename(columns={col: op["new_name"]})

    if kind == "drop_column":
        col = _get_col(df, op["column"])
        return df.drop(columns=[col])

    if kind == "keep_rows":
        mask = _filter_mask(df, op["column"], op["operator"], op.get("value"))
        return df[mask].reset_index(drop=True)

    if kind == "remove_rows":
        mask = _filter_mask(df, op["column"], op["operator"], op.get("value"))
        return df[~mask].reset_index(drop=True)

    if kind == "sort_rows":
        col = _get_col(df, op["column"])
        ascending = bool(op.get("ascending", True))
        return df.sort_values(by=col, ascending=ascending, kind="mergesort").reset_index(drop=True)

    if kind == "add_column":
        new_col = op["new_column"]
        formula = op.get("formula")
        constant = op.get("constant")
        if formula:
            try:
                df[new_col] = df.eval(formula, engine="python")
            except Exception as e:
                raise OperationError(f"Couldn't compute '{formula}': {e}")
        elif constant is not None:
            df[new_col] = constant
        else:
            raise OperationError("add_column needs either 'formula' or 'constant'.")
        return df

    if kind == "replace_values":
        col = _get_col(df, op["column"])
        df[col] = df[col].replace(op["find"], op["replace"])
        return df

    if kind == "fill_missing":
        col_name = op.get("column")
        method = op.get("method", "mean")
        cols = [_get_col(df, col_name)] if col_name else list(df.columns)
        for col in cols:
            series = df[col]
            if method in ("mean", "median"):
                numeric = pd.to_numeric(series, errors="coerce")
                if numeric.notna().sum() == 0:
                    continue  # non-numeric column: leave it alone rather than corrupt it
                fill_val = numeric.mean() if method == "mean" else numeric.median()
                df[col] = numeric.fillna(fill_val)
            elif method == "mode":
                modes = series.mode(dropna=True)
                if len(modes):
                    df[col] = series.fillna(modes.iloc[0])
            elif method == "constant":
                df[col] = series.fillna(op.get("value", ""))
            elif method == "ffill":
                df[col] = series.ffill()
            elif method == "bfill":
                df[col] = series.bfill()
            elif method == "drop_rows":
                df = df.dropna(subset=[col]).reset_index(drop=True)
            else:
                raise OperationError(f"Unknown fill method '{method}'.")
        return df

    if kind == "drop_duplicates":
        subset = op.get("columns")
        if subset:
            subset = [_get_col(df, c) for c in subset]
        return df.drop_duplicates(subset=subset, keep="first").reset_index(drop=True)

    if kind == "drop_blank_rows":
        return df.dropna(how="all").reset_index(drop=True)

    if kind == "change_case":
        col = _get_col(df, op["column"])
        func = _CASE_FUNCS.get(op.get("mode", "upper"))
        if not func:
            raise OperationError(f"Unknown case mode '{op.get('mode')}'.")
        df[col] = df[col].apply(lambda v: func(v) if isinstance(v, str) else v)
        return df

    if kind == "strip_whitespace":
        col_name = op.get("column")
        cols = [_get_col(df, col_name)] if col_name else df.select_dtypes(include="object").columns.tolist()
        for col in cols:
            df[col] = df[col].apply(lambda v: v.strip() if isinstance(v, str) else v)
        return df

    if kind == "round_numbers":
        col = _get_col(df, op["column"])
        decimals = int(op.get("decimals", 2))
        df[col] = pd.to_numeric(df[col], errors="coerce").round(decimals)
        return df

    if kind == "change_dtype":
        col = _get_col(df, op["column"])
        target = op.get("dtype", "text")
        try:
            if target == "number":
                df[col] = pd.to_numeric(df[col], errors="coerce")
            elif target == "text":
                df[col] = df[col].astype(str)
            elif target == "date":
                df[col] = pd.to_datetime(df[col], errors="coerce")
            else:
                raise OperationError(f"Unknown dtype '{target}'.")
        except OperationError:
            raise
        except Exception as e:
            raise OperationError(f"Couldn't convert '{col}' to {target}: {e}")
        return df

    if kind == "reorder_columns":
        order = [_get_col(df, c) for c in op.get("columns", [])]
        remaining = [c for c in df.columns if c not in order]
        return df[order + remaining]

    raise OperationError(f"Unsupported operation '{kind}'.")


def apply_operations(df: pd.DataFrame, ops: list) -> pd.DataFrame:
    for op in ops:
        df = apply_operation(df, op)
    return df


# ==========================================================================
# AI: translate a plain-English instruction into an operations plan
# ==========================================================================
def _build_system_prompt() -> str:
    return (
        "You are a spreadsheet-editing assistant. You will be given a profile of a "
        "table's columns (name, data type, missing count, unique count, a few sample "
        "values) and a plain-English instruction describing an edit the user wants. "
        "Translate the instruction into a JSON object with exactly this shape:\n"
        '{"operations": [ {"op": "<name>", ...fields}, ... ], '
        '"explanation": "<one short plain-English sentence summarizing what will change>"}\n\n'
        "Only use these operation names and fields, exactly as specified:\n"
        "- rename_column: {\"column\", \"new_name\"}\n"
        "- drop_column: {\"column\"}\n"
        "- keep_rows: {\"column\", \"operator\" (==,!=,>,<,>=,<=,contains,not_contains,"
        "is_null,not_null), \"value\" (omit for is_null/not_null)} — keeps only matching rows\n"
        "- remove_rows: same fields as keep_rows, but removes matching rows instead\n"
        "- sort_rows: {\"column\", \"ascending\" (bool)}\n"
        "- add_column: {\"new_column\", \"formula\" (a pandas-eval expression over existing "
        "column names, e.g. \"price * qty\") OR \"constant\" (a fixed value for every row)}\n"
        "- replace_values: {\"column\", \"find\", \"replace\"}\n"
        "- fill_missing: {\"column\" (omit for all columns), \"method\" "
        "(mean,median,mode,constant,ffill,bfill,drop_rows), \"value\" (only for constant)}\n"
        "- drop_duplicates: {\"columns\" (list, omit for all columns)}\n"
        "- drop_blank_rows: {}\n"
        "- change_case: {\"column\", \"mode\" (upper,lower,title)}\n"
        "- strip_whitespace: {\"column\" (omit for all text columns)}\n"
        "- round_numbers: {\"column\", \"decimals\" (int)}\n"
        "- change_dtype: {\"column\", \"dtype\" (number,text,date)}\n"
        "- reorder_columns: {\"columns\" (list, in the desired order; remaining columns keep "
        "their existing order after them)}\n\n"
        "Rules:\n"
        "1. Only reference column names that actually appear in the provided profile — never invent one.\n"
        "2. If the instruction is ambiguous, impossible with these operations, or needs a column "
        "that doesn't exist, return {\"operations\": [], \"explanation\": \"<why you couldn't do it>\"}.\n"
        "3. Use the smallest number of operations that satisfies the instruction.\n"
        "4. Return ONLY the JSON object — no markdown, no commentary outside the JSON."
    )


def _build_user_prompt(instruction, column_profile, sheet_name, row_count) -> str:
    return (
        f"Sheet: '{sheet_name}' — {row_count} rows, {len(column_profile)} columns.\n"
        f"Column profile:\n{json.dumps(column_profile, default=str)}\n\n"
        f"User instruction:\n{instruction}\n\n"
        "Return the JSON object described in the system instructions."
    )


def _parse_plan(raw_text: str):
    data = json.loads(raw_text)
    if not isinstance(data, dict) or "operations" not in data:
        raise ValueError("Model did not return the expected {operations, explanation} JSON shape.")
    ops = data.get("operations")
    if not isinstance(ops, list):
        raise ValueError("'operations' must be a list.")
    explanation = data.get("explanation", "") or ""
    return ops, explanation


def _make_client(provider: str, api_key: str, model_name: str):
    if provider == pe.PROVIDER_OPENAI:
        return OpenAI(api_key=api_key)
    elif provider == pe.PROVIDER_GEMINI:
        genai.configure(api_key=api_key)
        return genai.GenerativeModel(model_name, system_instruction=_build_system_prompt())
    raise ValueError(f"Unknown provider: {provider}")


def _call_openai_edit(client, model_name, instruction, column_profile, sheet_name, row_count):
    last_error = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            response = client.chat.completions.create(
                model=model_name,
                messages=[
                    {"role": "system", "content": _build_system_prompt()},
                    {"role": "user", "content": _build_user_prompt(
                        instruction, column_profile, sheet_name, row_count)},
                ],
                response_format={"type": "json_object"},
                temperature=0,
                timeout=API_TIMEOUT,
            )
            return _parse_plan(response.choices[0].message.content)
        except OpenAIAuthenticationError as e:
            raise AuthError(f"Invalid or missing OpenAI API key: {e}")
        except (OpenAIRateLimitError, OpenAIAPITimeoutError) as e:
            last_error = e
            if attempt < MAX_RETRIES:
                time.sleep(RETRY_BACKOFF_SECONDS * (2 ** (attempt - 1)))
                continue
            raise RuntimeError(f"OpenAI timed out/rate-limited after {MAX_RETRIES} attempts: {e}")
        except json.JSONDecodeError as e:
            last_error = e
            if attempt < MAX_RETRIES:
                continue
            raise RuntimeError(f"Model returned invalid JSON after {MAX_RETRIES} attempts: {e}")
        except OpenAIAPIError as e:
            last_error = e
            if attempt < MAX_RETRIES:
                time.sleep(RETRY_BACKOFF_SECONDS * (2 ** (attempt - 1)))
                continue
            raise RuntimeError(f"OpenAI API error after {MAX_RETRIES} attempts: {e}")
    raise RuntimeError(f"Edit planning failed: {last_error}")


def _call_gemini_edit(model, instruction, column_profile, sheet_name, row_count):
    last_error = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            response = model.generate_content(
                _build_user_prompt(instruction, column_profile, sheet_name, row_count),
                generation_config=genai.types.GenerationConfig(
                    response_mime_type="application/json",
                    temperature=0,
                ),
                request_options={"timeout": API_TIMEOUT},
            )
            return _parse_plan(response.text)
        except GooglePermissionDenied as e:
            raise AuthError(f"Invalid or missing Gemini API key: {e}")
        except GoogleInvalidArgument as e:
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
    raise RuntimeError(f"Edit planning failed: {last_error}")


def get_edit_plan(instruction: str, df: pd.DataFrame, sheet_name: str,
                   provider: str, api_key: str, model_name: str = None):
    """
    Returns (operations, explanation) for the given instruction.
    Raises AuthError on a missing/invalid key, RuntimeError on other API
    failures. Does NOT apply anything — see apply_operations() for that.
    """
    if provider not in (pe.PROVIDER_OPENAI, pe.PROVIDER_GEMINI):
        raise ValueError(f"Unknown provider: {provider}")
    if not is_provider_available(provider):
        pkg = "openai" if provider == pe.PROVIDER_OPENAI else "google-generativeai"
        raise RuntimeError(f"The '{pkg}' package isn't installed. Run: pip install {pkg}")
    if not api_key:
        raise AuthError(f"No {provider.title()} API key configured.")

    model_name = (model_name or "").strip()
    if not model_name:
        model_name = pe.DEFAULT_OPENAI_MODEL if provider == pe.PROVIDER_OPENAI else pe.DEFAULT_GEMINI_MODEL

    column_profile = build_column_profile(df)
    client = _make_client(provider, api_key, model_name)

    if provider == pe.PROVIDER_OPENAI:
        return _call_openai_edit(client, model_name, instruction, column_profile, sheet_name, len(df))
    return _call_gemini_edit(client, instruction, column_profile, sheet_name, len(df))


def is_provider_available(provider: str) -> bool:
    return OPENAI_AVAILABLE if provider == pe.PROVIDER_OPENAI else GEMINI_AVAILABLE


# ==========================================================================
# Edit session: current state of every sheet + per-sheet undo/redo + history
# ==========================================================================
class EditSession:
    """
    Holds all sheets loaded from one workbook, tracks the active sheet, and
    keeps a per-sheet undo/redo stack of DataFrame snapshots. Undo/redo is
    scoped to whichever sheet is active when you call it, since edits on
    different sheets are independent of each other.
    """

    def __init__(self, filepath: str, sheets: dict, sheet_order: list):
        self.filepath = filepath
        self.sheets = sheets
        self.sheet_order = list(sheet_order)
        self.active_sheet = sheet_order[0]
        self._undo_stacks = {name: [] for name in sheet_order}
        self._redo_stacks = {name: [] for name in sheet_order}
        self.history = []  # [{"sheet", "instruction", "explanation", "ops"}]

    @property
    def df(self) -> pd.DataFrame:
        return self.sheets[self.active_sheet]

    def set_active_sheet(self, name: str):
        if name not in self.sheets:
            raise ValueError(f"Sheet '{name}' not found.")
        self.active_sheet = name

    def apply(self, ops: list, instruction: str = "", explanation: str = "") -> pd.DataFrame:
        """Applies ops to the active sheet. Raises OperationError without changing state on failure."""
        current = self.sheets[self.active_sheet]
        new_df = apply_operations(current, ops)  # raises OperationError before anything is mutated

        stack = self._undo_stacks[self.active_sheet]
        stack.append(current)
        if len(stack) > MAX_UNDO_DEPTH:
            stack.pop(0)
        self._redo_stacks[self.active_sheet] = []

        self.sheets[self.active_sheet] = new_df
        self.history.append({
            "sheet": self.active_sheet,
            "instruction": instruction,
            "explanation": explanation,
            "ops": ops,
        })
        return new_df

    def can_undo(self) -> bool:
        return bool(self._undo_stacks.get(self.active_sheet))

    def can_redo(self) -> bool:
        return bool(self._redo_stacks.get(self.active_sheet))

    def undo(self) -> pd.DataFrame:
        stack = self._undo_stacks[self.active_sheet]
        if not stack:
            raise OperationError("Nothing to undo on this sheet.")
        self._redo_stacks[self.active_sheet].append(self.sheets[self.active_sheet])
        self.sheets[self.active_sheet] = stack.pop()
        return self.sheets[self.active_sheet]

    def redo(self) -> pd.DataFrame:
        redo_stack = self._redo_stacks[self.active_sheet]
        if not redo_stack:
            raise OperationError("Nothing to redo on this sheet.")
        self._undo_stacks[self.active_sheet].append(self.sheets[self.active_sheet])
        self.sheets[self.active_sheet] = redo_stack.pop()
        return self.sheets[self.active_sheet]

    def save(self, filepath: str = None) -> str:
        target = filepath or self.filepath
        save_workbook(self.sheets, target)
        self.filepath = target
        return target
