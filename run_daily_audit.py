# ============================================================================
# ArabWheels AE — MDM Spec Audit (GitHub Actions daily runner)
# ============================================================================
# Same audit logic as the Colab/OpenRouter version, but every file read/write
# goes through the Drive API (via drive_utils.py) instead of a local mounted
# filesystem, since GitHub Actions runners don't have Drive mounted.
#
# Each run:
#   1. Downloads the CSV from Drive (the audited one if it already exists
#      there, otherwise the original).
#   2. Lists every PDF under your brochures folder (metadata only — no PDF
#      bytes downloaded yet) and re-runs matching fresh every time (cheap,
#      local, and keeps things correct across the Colab -> Actions switch).
#   3. Processes brochure groups until DAILY_REQUEST_CAP is hit or the queue
#      is empty, downloading only the PDFs it actually needs to read today.
#   4. Uploads the updated CSV, highlighted .xlsx, discrepancy/new-trim
#      reports, audit log, and progress snapshot back to the SAME Drive
#      folder, overwriting the previous version of each — so tomorrow's run
#      picks up exactly where today's left off, same as the Colab version.
# ============================================================================

# ============================================================================
# ArabWheels AE — MDM Spec Audit (GitHub Actions daily runner — Gemini)
# ============================================================================
# Same audit logic as the Colab version, but every file read/write goes
# through the Drive API (via drive_utils.py) instead of a local mounted
# filesystem, since GitHub Actions runners don't have Drive mounted.
#
# Each run:
#   1. Downloads the CSV from Drive (the audited one if it already exists
#      there, otherwise the original).
#   2. Lists every PDF under your brochures folder (metadata only — no PDF
#      bytes downloaded yet) and re-runs matching fresh every time (cheap,
#      local, and keeps things correct across the Colab -> Actions switch).
#   3. Processes brochure groups until DAILY_REQUEST_CAP is hit or the queue
#      is empty, downloading only the PDFs it actually needs to read today.
#   4. Uploads the updated CSV, highlighted .xlsx, discrepancy/new-trim
#      reports, audit log, and progress snapshot back to the SAME Drive
#      folder, overwriting the previous version of each — so tomorrow's run
#      picks up exactly where today's left off, same as the Colab version.
# ============================================================================

import os
import re
import io
import json
import time
import uuid
import random
import logging
import threading
from datetime import datetime, timezone

import pandas as pd
import pypdf
from google import genai
from google.genai import types
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

import drive_utils

logging.getLogger("pypdf").setLevel(logging.ERROR)

# ============================== CONFIG =====================================

# Folder IDs from the Drive URL (…/folders/<THIS PART>). Set as GitHub repo
# "Variables" (not secrets — folder IDs aren't sensitive), read here as env vars.
GDRIVE_ROOT_FOLDER_ID = os.environ['GDRIVE_ROOT_FOLDER_ID']            # "AW AE MDM Project" folder
GDRIVE_BROCHURES_FOLDER_ID = os.environ['GDRIVE_BROCHURES_FOLDER_ID']  # extracted-PDFs folder

SOURCE_CSV_NAME = 'Arabwheels AE MDM data - MDM Version specs.csv'   # only used if no audited copy exists yet
OUTPUT_CSV_NAME = 'AUDITED_Arabwheels_MDM_data.csv'
OUTPUT_XLSX_HIGHLIGHTED_NAME = 'AUDITED_Arabwheels_MDM_data_HIGHLIGHTED.xlsx'
DISCREPANCY_DETAIL_NAME = 'Discrepancy_Detail.csv'
NEW_TRIMS_NAME = 'New_Trims_Found_In_Brochures.csv'
AUDIT_LOG_NAME = 'audit_log.jsonl'
SUMMARY_NAME = 'audit_summary.txt'
PROGRESS_NAME = 'audit_progress.json'
NOT_YET_AUDITED_NAME = 'Not_Yet_Audited.csv'

LOCAL_DIR = './run_data'
os.makedirs(LOCAL_DIR, exist_ok=True)
def local(name):
    return os.path.join(LOCAL_DIR, name)

GEMINI_API_KEY = os.environ['GEMINI_API_KEY']
client = genai.Client(api_key=GEMINI_API_KEY)

MODEL_NAME = os.environ.get('GEMINI_MODEL', 'gemini-3.6-flash')

MAX_TEXT_CHARS = 40_000
GROUP_BATCH_SIZE = 10
SAVE_EVERY_N_GROUPS = 5
MAX_RETRIES = 5

def get_int_env(name, default):
    """Reads an int env var, falling back to default if unset OR empty.

    GitHub Actions sets an env var to '' (not omitted) when it's assigned
    from a repo variable/secret that doesn't exist, e.g.
    `RATE_LIMIT_RPM: ${{ vars.RATE_LIMIT_RPM }}` with no such repo variable
    defined. In that case os.environ.get(name, default) still returns ''
    (the key IS present), so int() blows up. This checks for that case
    explicitly.
    """
    val = os.environ.get(name, '').strip()
    return int(val) if val else default


# Free-tier Gemini Flash is typically ~10-15 RPM depending on your account —
# check https://ai.google.dev/gemini-api/docs/rate-limits and set this to
# roughly 80% of your actual limit.
RATE_LIMIT_RPM = get_int_env('RATE_LIMIT_RPM', 10)
SLEEP_BETWEEN_REQUESTS = 60.0 / RATE_LIMIT_RPM

# This is what makes it a genuine daily partition: each GitHub Actions run
# does at most this many requests, then stops cleanly and picks back up on
# tomorrow's scheduled run. Gemini free tier is also capped per-day (varies
# by model/account) — set this to comfortably under your actual daily quota.
DAILY_REQUEST_CAP = get_int_env('DAILY_REQUEST_CAP', 100)

MAKE_COL, MODEL_COL, GEN_COL, VERSION_COL = (
    'manufacturer_name', 'model_name', 'generation_name', 'version_name'
)
ID_COL = 'version_spec_id'


BINARY_SPECS = [
    'air_conditioner', 'power_windows', 'power_door_locks', 'power_steering',
    'anti_lock_braking', 'traction_control', 'immobilizer', 'cup_holders',
    'rear_folding_seat', 'rear_wiper', 'alloy_wheels', 'tubeless_tyres',
    'central_locking', 'remote_boot', 'steering_adjustment', 'tachometer',
    'child_safety_locks', 'fog_lights', 'defroster', 'defogger', 'am_fm_radio',
    'cassette_player', 'cd_player', 'sun_roof', 'moon_roof', 'is_imported',
    'cool_box', 'dvd_player', 'cruise_control', 'keyless_entry', 'rear_ac_vents',
    'front_speakers', 'rear_speakers', 'usb_and_auxillary_cable', 'heated_seats',
    'steering_switches', 'front_camera', 'rear_camera', 'push_start',
    'hill_start_assist_control', 'isofix_child_seat_anchors', 'vehicle_stability_control'
]

NUMERIC_SPECS = [
    'overall_height', 'overall_length', 'overall_width', 'wheel_base',
    'ground_clearance', 'boot_space', 'kerb_weight', 'no_of_doors',
    'displacement', 'engine_power', 'torque', 'no_of_cylinders',
    'valves_per_cylinder', 'turning_radius', 'seating_capacity',
    'fuel_tank_capacity', 'mileage_high_way', 'mileage_city', 'mileage_overall',
    'max_speed', 'wheel_size', 'number_of_airbags', 'electric_motor_power',
    'charging_time', 'range', 'battery_capacity', 'compression_ratio',
]

TEXT_SPECS = [
    'valve_mechanism', 'cylinder_config', 'engine_type', 'fuel_system',
    'steering_type', 'power_assisted', 'transmission_type', 'gear_box',
    'wheel_type', 'tyre_size', 'suspensions', 'brakes', 'seat_material_type',
    'key_type', 'drive_train', 'battery_type', 'spare_tyre', 'spare_tyre_size',
    'transmission_technology', 'driving_modes',
]

NEW_ROW_COPY_COLS = [
    'make_id', 'model_id', 'gen_id', MAKE_COL, MODEL_COL, GEN_COL,
    'gen_launch_date', 'gen_close_date',
]

TERMINAL_STATUSES = {
    "Verified Accurate",
    "Flagged: Spec Discrepancy",
    "Flagged: Mismatched Brochure PDF",
    "Flagged: Missing Brochure PDF",
    "Flagged: Missing Make/Model in Sheet",
    "Flagged: Ambiguous Brochure Match",
    "Flagged: Image-Only or Unreadable PDF",
}
# These two are matching-time only (no LLM call) — safe to re-open if a
# brochure that didn't exist before now does.
REMATCHABLE_STATUSES = {"Flagged: Missing Brochure PDF", "Flagged: Ambiguous Brochure Match"}

FILENAME_NOISE_TOKENS = {
    'brochure', 'catalogue', 'catalog', 'spec', 'specs', 'specsheet', 'sheet',
    'product', 'compressed', 'final', 'pdf', 'en', 'uae', 'ksa', 'huge', 'pro',
    'v1', 'v2', 'v3', 'v4', 'new', 'update', 'updated', 'copy', 'draft',
}

_detail_lock = threading.Lock()
_new_rows_lock = threading.Lock()
_log_lock = threading.Lock()

spec_discrepancy_details = []
new_rows_from_brochures = []
_missing_trim_seen = set()
CHANGED_CELLS = {}

_daily_lock = threading.Lock()
_daily_count = 0

def daily_cap_try_consume():
    global _daily_count
    with _daily_lock:
        if _daily_count >= DAILY_REQUEST_CAP:
            return False
        _daily_count += 1
        return True

def daily_cap_exhausted():
    with _daily_lock:
        return _daily_count >= DAILY_REQUEST_CAP

def values_differ(old_val, new_val):
    if pd.isna(old_val) and (new_val is None or str(new_val).strip() == ''):
        return False
    try:
        return float(old_val) != float(new_val)
    except (ValueError, TypeError):
        return str(old_val).strip().lower() != str(new_val).strip().lower()

def log_event(payload: dict):
    payload["ts"] = datetime.now(timezone.utc).isoformat()
    with _log_lock:
        with open(local(AUDIT_LOG_NAME), "a", encoding="utf-8") as f:
            f.write(json.dumps(payload, ensure_ascii=False) + "\n")

# ============================== DRIVE SETUP =================================

print("Authenticating to Google Drive...")
drive, service_account_email = drive_utils.get_drive_service()

print("Checking folder access...")
drive_utils.check_folder_access(drive, GDRIVE_ROOT_FOLDER_ID, "GDRIVE_ROOT_FOLDER_ID", service_account_email)
drive_utils.check_folder_access(drive, GDRIVE_BROCHURES_FOLDER_ID, "GDRIVE_BROCHURES_FOLDER_ID", service_account_email)

existing_output_id = drive_utils.find_file_id_by_name(drive, GDRIVE_ROOT_FOLDER_ID, OUTPUT_CSV_NAME)
if existing_output_id:
    print("Found existing audited CSV on Drive — resuming from it.")
    drive_utils.download_to_path(drive, existing_output_id, local(OUTPUT_CSV_NAME))
    df = pd.read_csv(local(OUTPUT_CSV_NAME), low_memory=False)
else:
    print("No audited CSV on Drive yet — starting from the original source CSV.")
    source_id = drive_utils.find_file_id_by_name(drive, GDRIVE_ROOT_FOLDER_ID, SOURCE_CSV_NAME)
    if not source_id:
        raise FileNotFoundError(f"Couldn't find '{SOURCE_CSV_NAME}' in the Drive root folder.")
    drive_utils.download_to_path(drive, source_id, local(SOURCE_CSV_NAME))
    df = pd.read_csv(local(SOURCE_CSV_NAME), low_memory=False)

for col in ['Brochure_File_Found', 'Matched_Brochure_Path', 'Accuracy_Status', 'Discrepancies_Flagged', 'Match_Confidence']:
    if col not in df.columns:
        df[col] = False if col == 'Brochure_File_Found' else ""

df['Discrepancies_Flagged'] = df['Discrepancies_Flagged'].astype('object')
df['Accuracy_Status'] = df['Accuracy_Status'].astype('object')
for spec in BINARY_SPECS:
    if spec in df.columns:
        df[spec] = df[spec].astype('object')

total_rows = len(df)
already_terminal = df['Accuracy_Status'].isin(TERMINAL_STATUSES).sum()
print(f"Sheet has {total_rows} total rows — {already_terminal} already have a final audit "
      f"status, {total_rows - already_terminal} still need one.")

# ============================ MATCH ENGINE ===================================
# Re-run every time (cheap, local, no API cost) rather than trusting a
# previous run's Matched_Brochure_Path, since that column may have been
# written by a different environment (e.g. a prior Colab run) with a
# different path format.

_word_re = re.compile(r'[a-z0-9]+')

def tokenize(s: str) -> list:
    return _word_re.findall(str(s).lower())

print("\nListing brochure PDFs on Drive...")
pdf_index = {}      # file_id -> set(tokens from its virtual path)
pdf_vpaths = {}      # file_id -> human-readable virtual path (for logs/filenames)
for file_id, vpath in drive_utils.list_pdfs_recursive(drive, GDRIVE_BROCHURES_FOLDER_ID):
    clean = re.sub(r'[^a-z0-9]', ' ', vpath.lower())
    pdf_index[file_id] = set(tokenize(clean))
    pdf_vpaths[file_id] = vpath
print(f"Total PDFs found: {len(pdf_index)}")

def find_best_match(make, model, generation, start_year, end_year):
    make_str = str(make).lower().strip() if pd.notna(make) else ''
    model_str = str(model).lower().strip() if pd.notna(model) else ''

    if not make_str or not model_str:
        return None, "Flagged: Missing Make/Model in Sheet", None

    make_tokens = set(tokenize(make_str))
    model_tokens = set(tokenize(model_str)) - FILENAME_NOISE_TOKENS

    if not make_tokens or not model_tokens:
        return None, "Flagged: Missing Make/Model in Sheet", None

    candidates = []
    for file_id, path_tokens in pdf_index.items():
        path_lower = pdf_vpaths[file_id].lower()
        if not make_tokens.issubset(path_tokens):
            continue

        match_all = True
        for m_token in model_tokens:
            pattern = r'(?<![a-z0-9])' + re.escape(m_token) + r'(?![a-z0-9])'
            if not re.search(pattern, path_lower):
                match_all = False
                break
        if not match_all:
            continue

        gen_tokens = set(tokenize(generation)) if generation and str(generation).lower() != 'nan' else set()
        gen_hits = len(gen_tokens & path_tokens)
        score = (gen_hits, -len(path_tokens))
        candidates.append((score, file_id))

    if not candidates:
        return None, "Flagged: Missing Brochure PDF", None

    candidates.sort(key=lambda x: x[0], reverse=True)
    return candidates[0][1], None, "Exact"

print("Matching rows against brochures...")
rematched_count = 0
for index, row in df.iterrows():
    prior_status = row.get('Accuracy_Status')
    result, note, confidence = find_best_match(
        row.get(MAKE_COL, ''), row.get(MODEL_COL, ''), row.get(GEN_COL, ''),
        row.get('start_year'), row.get('end_year')
    )

    if note:
        df.at[index, 'Brochure_File_Found'] = False
        # Only overwrite a matching-time status; never touch an LLM-derived one.
        if prior_status not in TERMINAL_STATUSES or prior_status in REMATCHABLE_STATUSES:
            df.at[index, 'Accuracy_Status'] = note
    else:
        df.at[index, 'Brochure_File_Found'] = True
        df.at[index, 'Matched_Brochure_Path'] = pdf_vpaths[result]  # human-readable, for the sheet
        df.at[index, 'Match_Confidence'] = confidence
        if prior_status in REMATCHABLE_STATUSES:
            df.at[index, 'Accuracy_Status'] = ""  # newly matchable — reopen for audit
            rematched_count += 1

# lookup: vpath -> file_id, since Matched_Brochure_Path stores the readable
# path but downloading needs the Drive file id
vpath_to_id = {v: k for k, v in pdf_vpaths.items()}

if rematched_count:
    print(f"{rematched_count} row(s) that previously had no brochure now match one "
          f"(new PDFs added since the last run) — reopened for audit.")

# ============================ PDF TEXT (lazy download + cache) ==============

_pdf_text_cache = {}

def extract_pdf_text(vpath: str) -> str:
    if vpath in _pdf_text_cache:
        return _pdf_text_cache[vpath]
    file_id = vpath_to_id.get(vpath)
    if not file_id:
        return ""
    local_pdf_path = local(f"pdf_cache_{file_id}.pdf")
    try:
        drive_utils.download_to_path(drive, file_id, local_pdf_path)
        reader = pypdf.PdfReader(local_pdf_path)
        parts = []
        for i, page in enumerate(reader.pages):
            parts.append(f"\n--- PAGE {i + 1} ---\n" + (page.extract_text() or ""))
        full_text = "".join(parts)
        if len(full_text) > MAX_TEXT_CHARS:
            half = MAX_TEXT_CHARS // 2
            full_text = full_text[:half] + "\n\n...[TRUNCATED]...\n\n" + full_text[-half:]
    except Exception as e:
        print(f"  ! Failed to read/download {vpath}: {e}")
        full_text = ""
    finally:
        if os.path.exists(local_pdf_path):
            os.remove(local_pdf_path)  # runner disk is limited; don't hoard PDFs across the run
    _pdf_text_cache[vpath] = full_text
    return full_text

# ============================ GEMINI API CALL ================================

SYSTEM_INSTRUCTION = """
You are an automotive spec auditor. Do not assume the current flags/specs given to you are
correct — some are known to be wrong (e.g. a car marked as having a sunroof it doesn't, or
missing a radio it actually has). Re-derive every binary flag and spec value yourself from the
brochure text. Only report a spec_discrepancy when the brochure clearly states a different value
(ignore trivial rounding/unit-notation differences). Only report a missing trim when confident
it's genuinely absent from the KNOWN TRIMS list given to you.
"""

# Native structured output — Gemini enforces this shape server-side, so no
# markdown-fence stripping or "please return only JSON" prompting needed.
SPEC_DISCREPANCY_SCHEMA = types.Schema(
    type=types.Type.OBJECT,
    properties={
        "field": types.Schema(type=types.Type.STRING),
        "current_value": types.Schema(type=types.Type.STRING),
        "brochure_value": types.Schema(type=types.Type.STRING),
        "page_reference": types.Schema(type=types.Type.STRING),
    },
    required=["field", "current_value", "brochure_value"],
)
RESULT_ITEM_SCHEMA = types.Schema(
    type=types.Type.OBJECT,
    properties={
        "version_spec_id": types.Schema(type=types.Type.STRING),
        "brochure_covers_this_version": types.Schema(type=types.Type.BOOLEAN),
        "detected_brochure_model": types.Schema(type=types.Type.STRING),
        "verified_flags": types.Schema(
            type=types.Type.OBJECT,
            properties={spec: types.Schema(type=types.Type.INTEGER) for spec in BINARY_SPECS},
            required=list(BINARY_SPECS),
        ),
        "discrepancies": types.Schema(type=types.Type.ARRAY, items=types.Schema(type=types.Type.STRING)),
        "spec_discrepancies": types.Schema(type=types.Type.ARRAY, items=SPEC_DISCREPANCY_SCHEMA),
    },
    required=["version_spec_id", "brochure_covers_this_version", "verified_flags", "discrepancies"],
)
MISSING_TRIM_SCHEMA = types.Schema(
    type=types.Type.OBJECT,
    properties={
        "trim_name": types.Schema(type=types.Type.STRING),
        "key_specs": types.Schema(type=types.Type.OBJECT, properties={
            f: types.Schema(type=types.Type.STRING) for f in (NUMERIC_SPECS + TEXT_SPECS)
        }),
        "page_reference": types.Schema(type=types.Type.STRING),
    },
    required=["trim_name"],
)
RESPONSE_SCHEMA = types.Schema(
    type=types.Type.OBJECT,
    properties={
        "results": types.Schema(type=types.Type.ARRAY, items=RESULT_ITEM_SCHEMA),
        "missing_trims_in_brochure": types.Schema(type=types.Type.ARRAY, items=MISSING_TRIM_SCHEMA),
    },
    required=["results"],
)
GEN_CONFIG = types.GenerateContentConfig(
    system_instruction=SYSTEM_INSTRUCTION,
    response_mime_type="application/json",
    response_schema=RESPONSE_SCHEMA,
    temperature=0,
)

def parse_retry_delay_seconds(err_msg: str):
    m = re.search(r'retryDelay["\']?\s*[:=]\s*["\']?(\d+(?:\.\d+)?)s', err_msg)
    if m:
        return float(m.group(1))
    m = re.search(r'seconds["\']?\s*[:=]\s*(\d+)', err_msg)
    if m:
        return float(m.group(1))
    return None

def call_gemini_with_retry(user_prompt: str):
    for attempt in range(MAX_RETRIES):
        if not daily_cap_try_consume():
            return None, f"Daily request cap ({DAILY_REQUEST_CAP}) reached — stopping for today."
        try:
            time.sleep(SLEEP_BETWEEN_REQUESTS)
            response = client.models.generate_content(model=MODEL_NAME, contents=user_prompt, config=GEN_CONFIG)
            return response.text, None
        except Exception as e:
            msg = str(e)
            is_daily_quota = "PerDay" in msg or "daily" in msg.lower()
            if is_daily_quota:
                return None, f"Gemini Daily Quota Error: {msg[:300]}"
            if any(code in msg for code in ("503", "UNAVAILABLE", "429", "RESOURCE_EXHAUSTED")):
                delay = parse_retry_delay_seconds(msg)
                if delay is None:
                    delay = (2 ** (attempt + 1)) + random.uniform(1, 2)
                print(f"[Rate Limit] attempt {attempt + 1}/{MAX_RETRIES}, waiting {delay:.1f}s...")
                time.sleep(delay)
            else:
                return None, f"Gemini API Error: {msg}"
    return None, "Gemini API Error: Exceeded Retries"

def normalize_trim_key(name):
    return re.sub(r'[^a-z0-9]', '', str(name).lower())

def audit_group_gemini(make, model, generation, pdf_text, group_rows):
    all_results = []
    all_missing_trims = []
    all_version_names = [str(row.get(VERSION_COL, '')) for _, row in group_rows]

    for i in range(0, len(group_rows), GROUP_BATCH_SIZE):
        batch = group_rows[i:i + GROUP_BATCH_SIZE]
        versions_payload = []
        for idx, row in batch:
            current_flags = {c: int(row.get(c)) for c in BINARY_SPECS if c in df.columns and pd.notna(row.get(c))}
            current_specs = {c: str(row.get(c)) for c in (NUMERIC_SPECS + TEXT_SPECS) if c in df.columns and pd.notna(row.get(c))}
            versions_payload.append({
                "version_spec_id": str(row.get(ID_COL, idx)),
                "version_name": str(row.get(VERSION_COL, '')),
                "current_binary_flags": current_flags,
                "current_specs": current_specs
            })

        user_prompt = f"""
VEHICLE: {make} {model} ({generation})
VERSIONS: {json.dumps(versions_payload)}
KNOWN TRIMS: {json.dumps(all_version_names)}

BROCHURE TEXT:
{pdf_text}
"""
        res_text, err = call_gemini_with_retry(user_prompt)

        if err or not res_text:
            for idx, row in batch:
                all_results.append({"version_spec_id": str(row.get(ID_COL, idx)), "error": err})
            continue

        try:
            parsed = json.loads(res_text)
            all_results.extend(parsed.get("results", []))
            all_missing_trims.extend(parsed.get("missing_trims_in_brochure", []))
        except Exception as parse_err:
            for idx, row in batch:
                all_results.append({"version_spec_id": str(row.get(ID_COL, idx)), "error": f"Parse Failure: {parse_err}"})

    return all_results, all_missing_trims

def make_new_row_from_trim(sample_row, vpath, trim):
    trim_name = trim.get("trim_name", "").strip()
    if not trim_name:
        return None
    dedup_key = (vpath, normalize_trim_key(trim_name))
    if dedup_key in _missing_trim_seen:
        return None
    _missing_trim_seen.add(dedup_key)

    new_row = {col: None for col in df.columns}
    for col in NEW_ROW_COPY_COLS:
        if col in df.columns:
            new_row[col] = sample_row.get(col)
    new_row[VERSION_COL] = trim_name
    new_row[ID_COL] = f"NEW_{uuid.uuid4().hex[:10]}"
    key_specs = trim.get("key_specs", {}) or {}
    for k, v in key_specs.items():
        if k in df.columns and v not in (None, ""):
            new_row[k] = v
    new_row['Brochure_File_Found'] = True
    new_row['Matched_Brochure_Path'] = vpath
    new_row['Accuracy_Status'] = "New: Found In Brochure, Not In Sheet — Needs Review"
    page_ref = trim.get("page_reference", "")
    new_row['Discrepancies_Flagged'] = (
        f"Auto-extracted from brochure{(' (' + page_ref + ')') if page_ref else ''}. "
        f"version_spec_id is a placeholder ({new_row[ID_COL]}) — replace with a real ID before publishing."
    )
    return new_row

# ============================ EXECUTION =====================================

def write_progress_snapshot(status, groups_done, groups_total, extra=None):
    snapshot = {
        "status": status,
        "timestamp": datetime.now(timezone.utc).isoformat(),
        "groups_done_this_run": groups_done,
        "groups_remaining": max(0, groups_total - groups_done),
        **(extra or {}),
    }
    with open(local(PROGRESS_NAME), "w", encoding="utf-8") as f:
        json.dump(snapshot, f, ensure_ascii=False, indent=2)
    return snapshot

def report_unaudited_rows():
    pending_mask = (
        (df['Brochure_File_Found'] == True)
        & (df['Matched_Brochure_Path'].astype(str).str.len() > 0)
        & (~df['Accuracy_Status'].isin(TERMINAL_STATUSES))
    )
    pending = df.loc[pending_mask]
    if len(pending):
        cols = [c for c in [ID_COL, MAKE_COL, MODEL_COL, GEN_COL, VERSION_COL,
                             'Matched_Brochure_Path', 'Accuracy_Status'] if c in df.columns]
        pending[cols].to_csv(local(NOT_YET_AUDITED_NAME), index=False)
    return len(pending)

def upload_outputs():
    for fname in [OUTPUT_CSV_NAME, OUTPUT_XLSX_HIGHLIGHTED_NAME, DISCREPANCY_DETAIL_NAME,
                  NEW_TRIMS_NAME, AUDIT_LOG_NAME, SUMMARY_NAME, PROGRESS_NAME, NOT_YET_AUDITED_NAME]:
        path = local(fname)
        if os.path.exists(path):
            drive_utils.upload_or_update(drive, GDRIVE_ROOT_FOLDER_ID, path, remote_name=fname)
            print(f"  Uploaded {fname} to Drive.")

def run_pipeline():
    matched_df = df[(df['Brochure_File_Found'] == True) & (df['Matched_Brochure_Path'].astype(str).str.len() > 0)]
    to_process = [idx for idx, row in matched_df.iterrows() if row.get('Accuracy_Status') not in TERMINAL_STATUSES]

    grouped_by_pdf = {}
    for idx in to_process:
        vpath = df.at[idx, 'Matched_Brochure_Path']
        grouped_by_pdf.setdefault(vpath, []).append((idx, df.loc[idx]))

    total_groups = len(grouped_by_pdf)
    print(f"Total brochure groups pending: {total_groups} "
          f"(today's cap: {DAILY_REQUEST_CAP} requests, ~{DAILY_REQUEST_CAP // GROUP_BATCH_SIZE if GROUP_BATCH_SIZE else DAILY_REQUEST_CAP} "
          f"group(s) worth, depending on trims per brochure)")

    group_count = 0
    stopped_early = False
    for vpath, g_rows in grouped_by_pdf.items():
        if daily_cap_exhausted():
            print(f"\nDaily request cap ({DAILY_REQUEST_CAP}) reached — stopping cleanly here. "
                  f"Tomorrow's scheduled run will continue automatically.")
            stopped_early = True
            break

        pdf_text = extract_pdf_text(vpath)
        if not pdf_text.strip():
            for idx, _ in g_rows:
                df.at[idx, 'Accuracy_Status'] = "Flagged: Image-Only or Unreadable PDF"
            continue

        first_row = g_rows[0][1]
        results, missing_trims = audit_group_gemini(
            first_row.get(MAKE_COL), first_row.get(MODEL_COL), first_row.get(GEN_COL), pdf_text, g_rows
        )

        for trim in missing_trims:
            new_row = make_new_row_from_trim(first_row, vpath, trim)
            if new_row is not None:
                with _new_rows_lock:
                    new_rows_from_brochures.append(new_row)

        res_map = {str(r.get('version_spec_id')): r for r in results if 'version_spec_id' in r}

        for idx, row in g_rows:
            v_id = str(row.get(ID_COL, idx))
            res = res_map.get(v_id, {})

            if res.get('error'):
                df.at[idx, 'Accuracy_Status'] = res['error']
                log_event({"version_spec_id": v_id, "status": "error", "detail": res['error'], "brochure": vpath})
                continue

            if not res.get('brochure_covers_this_version', True):
                df.at[idx, 'Accuracy_Status'] = "Flagged: Mismatched Brochure PDF"
                log_event({"version_spec_id": v_id, "status": "mismatched_brochure", "brochure": vpath})
                continue

            verified_flags = res.get('verified_flags', {}) or {}
            spec_discrepancies = res.get('spec_discrepancies', []) or []
            row_changes = []

            for spec_col, val in verified_flags.items():
                if spec_col not in BINARY_SPECS or spec_col not in df.columns:
                    continue
                try:
                    new_val = 1 if int(val) != 0 else 0
                except (ValueError, TypeError):
                    continue
                old_val = df.at[idx, spec_col]
                if values_differ(old_val, new_val):
                    row_changes.append((spec_col, old_val, new_val, ''))
                    CHANGED_CELLS[(idx, spec_col)] = True
                df.at[idx, spec_col] = new_val

            for sd in spec_discrepancies:
                field = sd.get('field', '')
                brochure_val = sd.get('brochure_value', '')
                page_ref = sd.get('page_reference', '')
                if not field or field not in df.columns or brochure_val in (None, ''):
                    continue
                old_val = df.at[idx, field]
                if values_differ(old_val, brochure_val):
                    row_changes.append((field, old_val, brochure_val, page_ref))
                    CHANGED_CELLS[(idx, field)] = True
                    df.at[idx, field] = brochure_val
                with _detail_lock:
                    spec_discrepancy_details.append({
                        ID_COL: v_id, MAKE_COL: row.get(MAKE_COL, ''), MODEL_COL: row.get(MODEL_COL, ''),
                        VERSION_COL: row.get(VERSION_COL, ''), 'field': field, 'current_value': old_val,
                        'brochure_value': brochure_val, 'page_reference': page_ref, 'brochure_path': vpath,
                    })

            if row_changes:
                df.at[idx, 'Accuracy_Status'] = "Flagged: Spec Discrepancy"
                issue_parts = [f"{c}: {o} → {n}" + (f" (Page {p})" if p else "") for c, o, n, p in row_changes]
                df.at[idx, 'Discrepancies_Flagged'] = " | ".join(issue_parts)
            else:
                df.at[idx, 'Accuracy_Status'] = "Verified Accurate"
                df.at[idx, 'Discrepancies_Flagged'] = "No issues found — matches brochure."

            log_event({
                "version_spec_id": v_id, "status": df.at[idx, 'Accuracy_Status'],
                "changes": [{"field": c[0], "old": str(c[1]), "new": str(c[2]), "page": c[3]} for c in row_changes],
                "brochure": vpath,
            })

        group_count += 1
        if group_count % SAVE_EVERY_N_GROUPS == 0:
            df.to_csv(local(OUTPUT_CSV_NAME), index=False)
            write_progress_snapshot("in_progress", group_count, total_groups, {"last_brochure": vpath})
            upload_outputs()
            print(f"[{group_count}/{total_groups}] groups done, checkpoint uploaded to Drive.")

    df.to_csv(local(OUTPUT_CSV_NAME), index=False)

    unaudited_count = report_unaudited_rows()
    write_progress_snapshot(
        "stopped_daily_cap" if stopped_early else "complete",
        group_count, total_groups, {"rows_still_unaudited": int(unaudited_count)},
    )

    if CHANGED_CELLS:
        df.to_excel(local(OUTPUT_XLSX_HIGHLIGHTED_NAME), index=False, sheet_name='Audited', engine='openpyxl')
        import openpyxl
        wb = openpyxl.load_workbook(local(OUTPUT_XLSX_HIGHLIGHTED_NAME))
        ws = wb['Audited']
        red_fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
        col_positions = {name: i + 1 for i, name in enumerate(df.columns)}
        row_positions = {label: pos for pos, label in enumerate(df.index)}
        for (row_idx, col_name), _ in CHANGED_CELLS.items():
            if col_name not in col_positions or row_idx not in row_positions:
                continue
            excel_row = row_positions[row_idx] + 2
            excel_col_letter = get_column_letter(col_positions[col_name])
            ws[f'{excel_col_letter}{excel_row}'].fill = red_fill
        wb.save(local(OUTPUT_XLSX_HIGHLIGHTED_NAME))
        print(f"Highlighted {len(CHANGED_CELLS)} corrected cell(s) red in {OUTPUT_XLSX_HIGHLIGHTED_NAME}")

    if spec_discrepancy_details:
        pd.DataFrame(spec_discrepancy_details).to_csv(local(DISCREPANCY_DETAIL_NAME), index=False)
    if new_rows_from_brochures:
        pd.DataFrame(new_rows_from_brochures).to_csv(local(NEW_TRIMS_NAME), index=False)

    summary_counts = df['Accuracy_Status'].value_counts(dropna=False)
    summary_lines = ["AUDIT SUMMARY", "=" * 40] + [f"{s or '(unaudited)'}: {c}" for s, c in summary_counts.items()]
    summary_text = "\n".join(summary_lines)
    with open(local(SUMMARY_NAME), "w", encoding="utf-8") as f:
        f.write(summary_text)
    print("\n" + summary_text)

    print("\nUploading final outputs to Drive...")
    upload_outputs()

    if stopped_early:
        print(f"\nStopped after {group_count}/{total_groups} groups this run (daily cap). "
              f"Tomorrow's scheduled GitHub Actions run picks up automatically.")
    elif unaudited_count:
        print(f"\n{unaudited_count} row(s) still unaudited (likely a PDF read/parse issue) — "
              f"see {NOT_YET_AUDITED_NAME}.")
    else:
        print("\nAll pending groups processed — nothing left in the queue.")

if __name__ == "__main__":
    run_pipeline()
