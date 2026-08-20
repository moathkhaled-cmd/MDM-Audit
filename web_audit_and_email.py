#!/usr/bin/env python3
"""
web_audit_and_email.py

- Loads the existing audited CSV from Drive (AUDITED_... CSV) or the source CSV if none.
- For rows that do NOT have a brochure (Brochure_File_Found == False or empty), performs a web audit:
  - Uses Drive + Bing search to gather official website snippets (requires BING_API_KEY).
  - Calls Gemini (GEMINI_API_KEY) with a strict JSON schema asking for verified_flags (0/1 integers) for binary specs.
  - Updates the CSV row values for BINARY_SPECS with 0/1 from the model (only when model returns explicit 0/1).
  - Marks audited rows as "Audited" and audited_source as "Web-AI" and note: "No brochure — audited from web".
  - Rows where the model response is missing/ambiguous remain "Pending".
- Writes CSV and XLSX (highlight light yellow for rows audited from web) locally and uploads them back to Drive.

Environment variables expected:
- GDRIVE_ROOT_FOLDER_ID (required)
- GDRIVE_BROCHURES_FOLDER_ID (optional, not used here)
- GEMINI_API_KEY (required)
- GEMINI_MODEL (optional, default gemini-3.6-flash)
- BING_API_KEY (optional but recommended)

This script integrates non-brochure auditing without changing the existing run_daily_audit pipeline.
"""

import os
import json
import time
import re
from datetime import datetime

import pandas as pd
import requests
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from google import genai
from google.genai import types

import drive_utils

# ------------------ Config ------------------
GDRIVE_ROOT_FOLDER_ID = os.environ['GDRIVE_ROOT_FOLDER_ID']
OUTPUT_CSV_NAME = os.environ.get('OUTPUT_CSV_NAME', 'AUDITED_Arabwheels_MDM_data.csv')
OUTPUT_XLSX_NAME = os.environ.get('OUTPUT_XLSX_NAME', 'AUDITED_Arabwheels_MDM_data_WEB_AUDIT.xlsx')
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

# Gemini setup
GEMINI_API_KEY = os.environ['GEMINI_API_KEY']
client = genai.Client(api_key=GEMINI_API_KEY)
MODEL_NAME = os.environ.get('GEMINI_MODEL', 'gemini-3.6-flash')

# Bing
BING_KEY = os.environ.get('BING_API_KEY')
BING_ENDPOINT = 'https://api.bing.microsoft.com/v7.0/search'

# Behavior
TOP_K_PAGES = int(os.environ.get('WEB_AUDIT_TOP_K', '3'))
SLEEP_BETWEEN = float(os.environ.get('WEB_AUDIT_SLEEP', '0.3'))

# ------------------ Helpers ------------------

def bing_search(query, top_k=3):
    if not BING_KEY:
        return []
    headers = {'Ocp-Apim-Subscription-Key': BING_KEY}
    params = {'q': query, 'mkt': 'en-US', 'count': top_k}
    r = requests.get(BING_ENDPOINT, headers=headers, params=params, timeout=15)
    if r.status_code != 200:
        return []
    data = r.json()
    results = []
    for item in data.get('webPages', {}).get('value', [])[:top_k]:
        results.append({'url': item.get('url'), 'snippet': item.get('snippet'), 'name': item.get('name')})
    return results


# Schema: verified_flags: object with binary specs -> integer (0/1)
VERIFIED_SCHEMA = types.Schema(
    type=types.Type.OBJECT,
    properties={
        'verified_flags': types.Schema(
            type=types.Type.OBJECT,
            properties={spec: types.Schema(type=types.Type.INTEGER) for spec in BINARY_SPECS},
            required=list(BINARY_SPECS),
        ),
        'audit_confident': types.Schema(type=types.Type.BOOLEAN),
    },
    required=['verified_flags', 'audit_confident'],
)

GEN_CONFIG = types.GenerateContentConfig(
    system_instruction=(
        "You are an automotive spec auditor. You will be given a small set of official website snippets. "
        "For each binary feature, return 0 if the official sources do NOT show the feature, 1 if they clearly show it. "
        "If the sources are ambiguous or do not contain evidence, return 0.\n"
        "REPLY ONLY WITH JSON that matches the schema: {\"verified_flags\": {...}, \"audit_confident\": true|false}. "
        "Do NOT add any commentary or extra fields. Values for features must be integers 0 or 1."
    ),
    response_mime_type='application/json',
    response_schema=VERIFIED_SCHEMA,
    temperature=0,
)


def call_gemini_with_schema(contents):
    # wrapper similar to run_daily_audit
    try:
        response = client.models.generate_content(model=MODEL_NAME, contents=contents, config=GEN_CONFIG)
        return response.text, None
    except Exception as e:
        return None, str(e)


# ------------------ Main flow ------------------

def main():
    print('Authenticating to Drive...')
    drive, _ = drive_utils.get_drive_service()

    # find the audited CSV on Drive; if not found, fall back to original source (same name used in run_daily_audit)
    file_id = drive_utils.find_file_id_by_name(drive, GDRIVE_ROOT_FOLDER_ID, OUTPUT_CSV_NAME)
    if not file_id:
        # try the original source name used in run_daily_audit
        print(f'{OUTPUT_CSV_NAME} not found on Drive; looking for source CSV')
        # fallback names (common)
        candidates = ['Arabwheels AE MDM data - MDM Version specs.csv', OUTPUT_CSV_NAME]
        file_id = None
        for c in candidates:
            file_id = drive_utils.find_file_id_by_name(drive, GDRIVE_ROOT_FOLDER_ID, c)
            if file_id:
                break
        if not file_id:
            raise SystemExit('Could not find any CSV to audit on Drive. Put your CSV in the root Drive folder and retry.')

    local_csv = './run_data/web_audit_input.csv'
    os.makedirs(os.path.dirname(local_csv) or '.', exist_ok=True)
    drive_utils.download_to_path(drive, file_id, local_csv)

    df = pd.read_csv(local_csv, low_memory=False, dtype=object)

    # ensure audit columns exist
    if 'audited_status' not in df.columns:
        df['audited_status'] = ''
    if 'audited_source' not in df.columns:
        df['audited_source'] = ''
    if 'note' not in df.columns:
        df['note'] = ''

    # target rows: no brochure found OR flagged missing brochure
    mask_no_brochure = (~df.get('Brochure_File_Found', False).astype(bool))
    target_idxs = df[mask_no_brochure].index.tolist()
    print(f'Rows without brochures to web-audit: {len(target_idxs)}')

    updated_count = 0
    for idx in target_idxs:
        make = str(df.at[idx, 'manufacturer_name'] or '')
        model = str(df.at[idx, 'model_name'] or '')
        generation = str(df.at[idx, 'generation_name'] or '')
        version = str(df.at[idx, 'version_name'] or '')

        query = f"{make} {model} {generation} official specifications site"
        pages = bing_search(query, top_k=TOP_K_PAGES)
        pages_text = '\n'.join([f"{p['url']} - {p['snippet']}" for p in pages]) if pages else 'No pages found.'

        # Build prompt content
        contents = f"VEHICLE: {make} {model} {generation} ({version})\nSOURCES:\n{pages_text}\n\n"
        contents += "ASK: For each binary feature, return verified_flags mapping 0/1 and audit_confident:true|false."

        res_text, err = call_gemini_with_schema(contents)
        if err or not res_text:
            print(f'Gemini error for row {idx}: {err} — marking Pending')
            df.at[idx, 'audited_status'] = 'Pending'
            df.at[idx, 'audited_source'] = ''
            time.sleep(SLEEP_BETWEEN)
            continue

        # parse JSON only (Gemini returns .text which should be JSON string)
        try:
            parsed = json.loads(res_text)
        except Exception as e:
            print(f'Parse failure for row {idx}: {e} — marking Pending')
            df.at[idx, 'audited_status'] = 'Pending'
            df.at[idx, 'audited_source'] = ''
            time.sleep(SLEEP_BETWEEN)
            continue

        verified = parsed.get('verified_flags', {}) or {}
        audit_confident = bool(parsed.get('audit_confident'))

        # apply only explicit 0/1 integers
        any_written = False
        for spec in BINARY_SPECS:
            if spec in verified:
                try:
                    v = int(verified[spec])
                except Exception:
                    continue
                if v in (0, 1):
                    df.at[idx, spec] = v
                    any_written = True

        if any_written:
            df.at[idx, 'audited_status'] = 'Audited'
            df.at[idx, 'audited_source'] = 'Web-AI'
            df.at[idx, 'note'] = 'No brochure — audited from web'
            updated_count += 1
        else:
            df.at[idx, 'audited_status'] = 'Pending'
            df.at[idx, 'audited_source'] = ''

        # polite sleep
        time.sleep(SLEEP_BETWEEN)

    print(f'Web audit updated {updated_count} row(s).')

    # save CSV and XLSX (highlight yellow rows which have the note)
    out_csv_local = './run_data/' + OUTPUT_CSV_NAME
    df.to_csv(out_csv_local, index=False)

    # XLSX with yellow highlight for note containing "No brochure — audited from web"
    wb = Workbook()
    ws = wb.active
    ws.title = 'Audit'
    headers = list(df.columns)
    ws.append(headers)
    yellow = PatternFill(start_color='FFF7B0', end_color='FFF7B0', fill_type='solid')
    for _, row in df.iterrows():
        ws.append([row.get(h, '') for h in headers])
    # apply highlight
    note_col = None
    for i, h in enumerate(headers, start=1):
        if h == 'note':
            note_col = i
            break
    if note_col:
        for r_idx in range(2, ws.max_row + 1):
            cell = ws.cell(row=r_idx, column=note_col)
            if isinstance(cell.value, str) and 'No brochure — audited from web' in cell.value:
                for c in range(1, len(headers) + 1):
                    ws.cell(row=r_idx, column=c).fill = yellow

    out_xlsx_local = './run_data/' + OUTPUT_XLSX_NAME
    wb.save(out_xlsx_local)

    # upload back to Drive
    drive_utils.upload_or_update(drive, GDRIVE_ROOT_FOLDER_ID, out_csv_local, remote_name=OUTPUT_CSV_NAME)
    drive_utils.upload_or_update(drive, GDRIVE_ROOT_FOLDER_ID, out_xlsx_local, remote_name=OUTPUT_XLSX_NAME)
    print('Uploaded updated CSV and XLSX to Drive.')


if __name__ == '__main__':
    main()
