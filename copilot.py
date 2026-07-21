"""
copilot.py

AI Data Copilot: takes a plain-English instruction, asks the configured
AI provider to translate it into a structured, step-by-step workflow
plan (JSON), and then executes that plan by calling into the app's
existing modules (downloader, pdf_extractor, data_profiler, sheet_editor,
dashboard_builder, ai_assistant).

Design choices, on purpose:
- No scheduler and no drag-drop canvas here. This module only does
  "understand -> plan -> run once, now". Recurring/scheduled workflows
  and a visual node editor are separate features to build on top of
  this, later — bolting a fake scheduler on top of a single-shot
  executor would just silently drop the "every morning" part of a
  request.
- The model NEVER invents file paths. Any path-like parameter it can't
  read directly from the instruction is left blank ("") and the UI is
  expected to ask the user to fill it in (browse dialog) before running.
  This avoids a workflow silently reading/writing the wrong folder.
- Requests for things this app can't do yet (send email, run on a
  timer, touch files outside what the user picks) are reported back
  as `unsupported` rather than silently dropped or faked.
"""
import os
import glob
import json

import pandas as pd

import pdf_extractor as pe
import data_profiler as dp
import sheet_editor as se
import dashboard_builder as db
import ai_assistant as aa
import downloader as dl


class PlanError(Exception):
    """Raised when the AI's plan can't be parsed or is structurally invalid."""


class StepError(Exception):
    """Raised when a single workflow step fails during execution."""


# --------------------------------------------------------------------------
# Action schema — the single source of truth for both the planning prompt
# and the executor's dispatch table. Add a new action by adding one entry
# here, one dispatch case in run_step(), and (if it needs one) one function.
# --------------------------------------------------------------------------
ACTION_SCHEMA = {
    "merge_excel": {
        "description": "Combine every .xlsx/.xls/.csv file in a folder into one sheet.",
        "params": {"folder": "path to the folder of files to combine",
                    "output_path": "path to write the combined .xlsx to"},
    },
    "clean_data": {
        "description": "Remove duplicate/blank rows, trim whitespace, on a spreadsheet.",
        "params": {"input_path": "path to the spreadsheet to clean",
                    "output_path": "path to write the cleaned file to",
                    "remove_duplicates": "true/false",
                    "remove_blank_rows": "true/false",
                    "trim_whitespace": "true/false",
                    "dedupe_keep": "'first' or 'last' — which duplicate row to keep",
                    "dedupe_sort_by": "column name to sort by before deduping (e.g. a date "
                                       "column), so 'keep last' means 'keep most recent'; "
                                       "null if not applicable"},
    },
    "extract_pdfs": {
        "description": "Read every PDF in a folder and extract requested fields into Excel.",
        "params": {"folder": "path to the folder of PDFs",
                    "fields": "plain-English description of what fields to extract",
                    "output_path": "path to write the extracted-data .xlsx to",
                    "sheet_name": "sheet name to write to, default 'Extracted Data'"},
    },
    "download_links": {
        "description": "Download every URL found in a workbook's cells to a folder.",
        "params": {"workbook_path": "path to the .xlsx containing links",
                    "sheet_name": "sheet name to scan for links",
                    "save_folder": "folder to save downloaded files into"},
    },
    "edit_sheet": {
        "description": "Apply a plain-English edit instruction to a spreadsheet (filter/sort/etc).",
        "params": {"input_path": "path to the spreadsheet to edit",
                    "sheet_name": "which sheet to edit",
                    "instruction": "the plain-English edit instruction",
                    "output_path": "path to save the edited file to"},
    },
    "build_dashboard": {
        "description": "Auto-generate a chart dashboard (Excel with embedded charts) from a spreadsheet.",
        "params": {"input_path": "path to the source spreadsheet",
                    "output_path": "path to write the dashboard .xlsx to",
                    "include_ai_summary": "true/false — add an AI-written summary of findings"},
    },
    "export_power_bi": {
        "description": "Export a cleaned, Power-BI-ready copy of a spreadsheet.",
        "params": {"input_path": "path to the source spreadsheet",
                    "output_path": "path to write the Power BI-ready .xlsx to"},
    },
    "ask_ai": {
        "description": "Freeform: ask the AI Assistant to write SQL/DAX/Python/VBA/etc.",
        "params": {"prompt": "the freeform question/instruction"},
    },
}

# Things people commonly ask for that this app genuinely can't do yet.
# The planner is told to name these explicitly instead of pretending.
KNOWN_UNSUPPORTED = [
    "sending email (no email/SMTP integration exists yet)",
    "running on a schedule / recurring automation (workflows run once, on demand, for now)",
    "reading files the user hasn't pointed the app at (no unrestricted filesystem/cloud access)",
]


def _system_prompt() -> str:
    schema_text = json.dumps(ACTION_SCHEMA, indent=2)
    return (
        "You are the planning engine for a desktop data-automation app. Convert the "
        "user's plain-English request into a JSON workflow plan using ONLY the action "
        "types below — never invent a new action type.\n\n"
        f"Available actions:\n{schema_text}\n\n"
        "Rules:\n"
        "1. Output ONLY a single JSON object, no markdown fences, no commentary.\n"
        "2. Shape: {\"workflow_name\": str, \"steps\": [{\"type\": str, \"label\": str, "
        "\"params\": {...}}], \"unsupported\": [str], \"clarifying_question\": str|null}\n"
        "3. For any path-like param (folder, input_path, output_path, workbook_path, "
        "save_folder) that the user's message doesn't literally specify, set it to an "
        "empty string \"\" — never guess a real filesystem path.\n"
        "4. For non-path params, fill in your best reasonable interpretation of the "
        "request (e.g. 'remove duplicates, keep latest' -> remove_duplicates: true, "
        "dedupe_keep: 'last').\n"
        "5. If the request includes something outside the action list (sending email, "
        "scheduling/recurring runs, etc.), do NOT add a fake step for it — instead add a "
        "short plain-English description of it to \"unsupported\".\n"
        "6. \"label\" is a short human-readable title for that step, e.g. 'Combine Excel "
        "files from folder'.\n"
        "7. If the request is too vague to plan at all (e.g. no data source of any kind "
        "mentioned), leave \"steps\" empty and set \"clarifying_question\" to a single "
        "specific question. Otherwise clarifying_question is null.\n"
    )


def _call_ai_json(provider: str, api_key: str, model_name: str, user_prompt: str) -> dict:
    """Shared JSON-mode call, mirroring the pattern already used in pdf_extractor.py
    and sheet_editor.py, but generic across both providers for the planner's use."""
    model_name = (model_name or "").strip()
    if not model_name:
        model_name = pe.DEFAULT_OPENAI_MODEL if provider == pe.PROVIDER_OPENAI else pe.DEFAULT_GEMINI_MODEL

    if provider == pe.PROVIDER_OPENAI:
        client = pe.OpenAI(api_key=api_key)
        resp = client.chat.completions.create(
            model=model_name,
            messages=[{"role": "system", "content": _system_prompt()},
                      {"role": "user", "content": user_prompt}],
            response_format={"type": "json_object"},
            temperature=0,
        )
        raw = resp.choices[0].message.content
    elif provider == pe.PROVIDER_GEMINI:
        pe.genai.configure(api_key=api_key)
        model = pe.genai.GenerativeModel(model_name, system_instruction=_system_prompt())
        resp = model.generate_content(
            user_prompt,
            generation_config=pe.genai.types.GenerationConfig(
                response_mime_type="application/json", temperature=0),
            request_options={"timeout": pe.API_TIMEOUT},
        )
        raw = resp.text
    else:
        raise ValueError(f"Unknown provider: {provider}")

    try:
        data = json.loads(raw)
    except (json.JSONDecodeError, TypeError) as e:
        raise PlanError(f"AI didn't return valid JSON: {e}")
    if not isinstance(data, dict):
        raise PlanError("AI response wasn't a JSON object.")
    return data


def plan_workflow(instruction: str, provider: str, api_key: str, model_name: str = None) -> dict:
    """
    Returns {"workflow_name", "steps", "unsupported", "clarifying_question"}.
    Raises pdf_extractor.AuthError on a bad/missing key, PlanError on a
    malformed response, RuntimeError on other API failures.
    """
    if not instruction or not instruction.strip():
        raise ValueError("Please describe what you want the workflow to do.")
    if provider not in (pe.PROVIDER_OPENAI, pe.PROVIDER_GEMINI):
        raise ValueError(f"Unknown provider: {provider}")
    if not pe.is_provider_available(provider):
        pkg = "openai" if provider == pe.PROVIDER_OPENAI else "google-generativeai"
        raise RuntimeError(f"The '{pkg}' package isn't installed. Run: pip install {pkg}")
    if not api_key:
        raise pe.AuthError(f"No {provider.title()} API key configured.")

    data = _call_ai_json(provider, api_key, model_name, instruction.strip())

    steps = data.get("steps") or []
    if not isinstance(steps, list):
        raise PlanError("Plan's 'steps' field wasn't a list.")
    for i, step in enumerate(steps):
        if not isinstance(step, dict) or "type" not in step:
            raise PlanError(f"Step {i} is malformed.")
        if step["type"] not in ACTION_SCHEMA:
            raise PlanError(f"Step {i} used an unknown action type: {step['type']}")
        step.setdefault("params", {})
        step.setdefault("label", step["type"].replace("_", " ").title())

    return {
        "workflow_name": data.get("workflow_name") or "Untitled Workflow",
        "steps": steps,
        "unsupported": data.get("unsupported") or [],
        "clarifying_question": data.get("clarifying_question"),
    }


# --------------------------------------------------------------------------
# Executor
# --------------------------------------------------------------------------
def merge_excel_files(folder: str, output_path: str, log_cb=None) -> pd.DataFrame:
    if not folder or not os.path.isdir(folder):
        raise StepError(f"Folder not found: {folder}")
    paths = sorted(
        p for p in glob.glob(os.path.join(folder, "*"))
        if p.lower().endswith((".xlsx", ".xls", ".csv"))
    )
    if not paths:
        raise StepError(f"No .xlsx/.xls/.csv files found in {folder}")

    frames = []
    for p in paths:
        try:
            if p.lower().endswith(".csv"):
                frame = pd.read_csv(p)
            else:
                frame = pd.read_excel(p)
            frame["Source File"] = os.path.basename(p)
            frames.append(frame)
            if log_cb:
                log_cb(f"  Read {os.path.basename(p)} ({len(frame)} rows)")
        except Exception as e:
            if log_cb:
                log_cb(f"  Skipped {os.path.basename(p)}: {e}")

    if not frames:
        raise StepError("None of the files in that folder could be read.")

    combined = pd.concat(frames, ignore_index=True, sort=False)
    combined.to_excel(output_path, index=False)
    return combined


def _require(params: dict, *keys):
    missing = [k for k in keys if not params.get(k)]
    if missing:
        raise StepError(f"Missing required value(s): {', '.join(missing)}")


def run_step(step: dict, provider: str, api_key: str, model_name: str,
             context: dict, log_cb) -> dict:
    """
    Executes one step. `context` carries state between steps in a run
    (currently just the last successful output path, so e.g. a clean_data
    step after a merge_excel step can default to using its output).
    Returns a small result dict; raises StepError on failure.
    """
    step_type = step["type"]
    params = dict(step.get("params") or {})
    log_cb(f"-> {step.get('label', step_type)}")

    if step_type == "merge_excel":
        _require(params, "folder", "output_path")
        merge_excel_files(params["folder"], params["output_path"], log_cb)
        context["last_output"] = params["output_path"]
        return {"output_path": params["output_path"]}

    elif step_type == "clean_data":
        input_path = params.get("input_path") or context.get("last_output")
        _require({"input_path": input_path, "output_path": params.get("output_path")},
                 "input_path", "output_path")
        df = dp.load_data(input_path)
        if params.get("dedupe_sort_by"):
            try:
                df = df.sort_values(params["dedupe_sort_by"])
            except KeyError:
                log_cb(f"  Note: column '{params['dedupe_sort_by']}' not found, skipping sort.")
        filters = {
            "duplicate_rows": bool(params.get("remove_duplicates")),
            "blank_rows": bool(params.get("remove_blank_rows")),
            "extra_spaces": bool(params.get("trim_whitespace")),
        }
        cleaned = dp.clean_data(df, filters)
        dp.save_data(cleaned, params["output_path"])
        context["last_output"] = params["output_path"]
        return {"output_path": params["output_path"], "rows": len(cleaned)}

    elif step_type == "extract_pdfs":
        _require(params, "folder", "fields", "output_path")
        pdfs = pe.list_pdfs(params["folder"])
        if not pdfs:
            raise StepError(f"No PDFs found in {params['folder']}")
        results = pe.run_extraction(
            pdfs, params["fields"], provider, api_key,
            max_workers=5, progress_cb=lambda d, t: None, log_cb=log_cb, model_name=model_name)
        pe.compile_to_excel(results, params["output_path"], params.get("sheet_name") or "Extracted Data")
        context["last_output"] = params["output_path"]
        return {"output_path": params["output_path"], "count": len(results)}

    elif step_type == "download_links":
        _require(params, "workbook_path", "sheet_name", "save_folder")
        os.makedirs(params["save_folder"], exist_ok=True)
        ws = dl.load_sheet(params["workbook_path"], params["sheet_name"])
        tasks = dl.build_download_tasks(ws, 2, ws.max_row, 1, ws.max_column, params["save_folder"])
        ok, failed = dl.run_downloads(tasks, max_workers=10, timeout=30,
                                       progress_cb=lambda d, t: None, log_cb=log_cb)
        return {"downloaded": ok, "failed": failed}

    elif step_type == "edit_sheet":
        input_path = params.get("input_path") or context.get("last_output")
        _require({"input_path": input_path, "instruction": params.get("instruction"),
                   "output_path": params.get("output_path")},
                 "input_path", "instruction", "output_path")
        sheets, order = se.load_workbook(input_path)
        sheet_name = params.get("sheet_name") or order[0]
        df = sheets[sheet_name]
        ops, explanation = se.get_edit_plan(params["instruction"], df, sheet_name,
                                             provider, api_key, model_name)
        sheets[sheet_name] = se.apply_operations(df, ops)
        se.save_workbook(sheets, params["output_path"])
        context["last_output"] = params["output_path"]
        return {"output_path": params["output_path"], "explanation": explanation}

    elif step_type == "build_dashboard":
        input_path = params.get("input_path") or context.get("last_output")
        _require({"input_path": input_path, "output_path": params.get("output_path")},
                 "input_path", "output_path")
        sheets, order = se.load_workbook(input_path)
        df = sheets[order[0]]
        column_info = db.classify_columns(df)
        plan = db.suggest_chart_plan(df, column_info)
        summary = ""
        if params.get("include_ai_summary"):
            aggregates = db.compute_group_aggregates(df, column_info)
            summary = db.generate_ai_summary(column_info, plan, provider, api_key,
                                              model_name=model_name, aggregates=aggregates)
        db.build_excel_dashboard(df, plan, params["output_path"], summary_text=summary)
        context["last_output"] = params["output_path"]
        return {"output_path": params["output_path"], "charts": len(plan)}

    elif step_type == "export_power_bi":
        input_path = params.get("input_path") or context.get("last_output")
        _require({"input_path": input_path, "output_path": params.get("output_path")},
                 "input_path", "output_path")
        sheets, order = se.load_workbook(input_path)
        df = sheets[order[0]]
        out = db.export_power_bi_ready(df, params["output_path"])
        context["last_output"] = out
        return {"output_path": out}

    elif step_type == "ask_ai":
        _require(params, "prompt")
        response = aa.ask_assistant(params["prompt"], [], provider, api_key, model_name=model_name)
        return {"response": response}

    else:
        raise StepError(f"Unknown action type: {step_type}")


def run_workflow(steps: list, provider: str, api_key: str, model_name: str, log_cb):
    """
    Runs steps in order. Stops at the first failure (later steps typically
    depend on earlier output) and returns (results, error_index_or_None).
    """
    context = {}
    results = []
    for i, step in enumerate(steps):
        try:
            result = run_step(step, provider, api_key, model_name, context, log_cb)
            log_cb(f"   done: {result}")
            results.append({"ok": True, "result": result})
        except (StepError, pe.AuthError, ValueError) as e:
            log_cb(f"   FAILED: {e}")
            results.append({"ok": False, "error": str(e)})
            return results, i
        except Exception as e:
            log_cb(f"   FAILED (unexpected): {e}")
            results.append({"ok": False, "error": str(e)})
            return results, i
    return results, None
