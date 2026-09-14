# MDM-Audit: Fixing Free Tier Quota Exhaustion

## The Core Problem

**Your 6 Gemini API keys are NOT from 6 different Google Cloud projects** — they're 6 keys from the **same project**.

This means:
- All 6 keys share ONE daily quota pool (~10 requests/day on free tier)
- Having more keys doesn't multiply your quota
- All 6 keys exhaust simultaneously, then ALL rotate away
- Result: Your audit fails with 391 web-search errors without trying other approaches

---

## Fix #1: Use 6 SEPARATE Google Cloud Projects (REQUIRED)

### Step 1: Create 6 New Google Cloud Projects

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Click the project dropdown → **NEW PROJECT**
3. Name each: `mdm-audit-1`, `mdm-audit-2`, ... `mdm-audit-6`
4. For each project:
   - Enable the **Generative Language API**
   - Enable the **Google Drive API**
   - Create an **API key** in Credentials section
   - Copy the key (you'll need 6 keys total)

### Step 2: Update GitHub Secrets

Replace your current 6 keys with keys from 6 DIFFERENT projects:

```
GEMINI_API_KEY       → key from project 1
GEMINI_API_KEY_2     → key from project 2
GEMINI_API_KEY_3     → key from project 3
GEMINI_API_KEY_4     → key from project 4
GEMINI_API_KEY_5     → key from project 5
GEMINI_API_KEY_6     → key from project 6
```

Each key now has its OWN quota:
- **Free tier per project**: ~10 requests/day per project
- **Total: ~60 requests/day** (6 projects × 10 requests)
- **Web search grounding**: ~1,500 requests/day per project on free tier
- **Total web search: ~9,000 requests/day** (if each project uses free grounding)

---

## Fix #2: Optimize the Script for Free Tier

Apply these changes to `run_daily_audit.py`:

### Change 1: Reduce batch sizes (less tokens per call)

```python
# Line 195
MAX_TEXT_CHARS = get_int_env('MAX_TEXT_CHARS', 12_000)  # Cut in half to 12K

# Line 203
GROUP_BATCH_SIZE = get_int_env('GROUP_BATCH_SIZE', 8)   # Cut from 16 to 8 trims/call
```

**Why**: Smaller batches = fewer tokens per request = more requests fit in daily quota

### Change 2: Increase rate limits (use all projects effectively)

```python
# Line 187-188
RATE_LIMIT_RPM = get_int_env('RATE_LIMIT_RPM', 30)              # 30 RPM (6 projects × 5 RPM each)
WEB_SEARCH_RATE_LIMIT_RPM = get_int_env('WEB_SEARCH_RATE_LIMIT_RPM', 12)  # 12 RPM
```

**Why**: Each project can handle 1-2 RPM safely; 6 projects = 6-12 total RPM capacity

### Change 3: Lower daily cap (match free tier reality)

```python
# Line 220
DAILY_REQUEST_CAP = get_int_env('DAILY_REQUEST_CAP', 50)  # 50 requests/day (6 projects × 8-9 each)
```

**Why**: Free tier is ~10 req/day per project; 50 total is realistic and safe

### Change 4: Reduce timeout

```python
# In .github/workflows/daily_audit.yml, line 22:
timeout-minutes: 45  # Down from 90; free tier runs finish faster anyway
```

### Change 5: Simplify the model

```python
# Line 158
MODEL_NAME = os.environ.get('GEMINI_MODEL', 'gemini-2.5-flash')  # Use Flash instead of 3.6-flash
```

**Why**: Flash is cheaper on tokens and faster; adequate for spec extraction

---

## Fix #3: Restructure for High-Impact Filtering

**Before calling Gemini for every vehicle**, filter out obvious non-matches first:

### Add to run_daily_audit.py before the brochure audit loop:

```python
def quick_filter_requires_audit(row):
    """Skip rows that definitely don't need Gemini."""
    # Already audited with verified result
    if row.get('Accuracy_Status') in ("Verified Accurate", "Flagged: Spec Discrepancy"):
        return False
    # No brochure and no make/model (can't search)
    if row.get('Brochure_File_Found') != True and (
        not row.get('manufacturer_name') or not row.get('model_name')
    ):
        return False
    return True

# Apply before brochure loop (around line 1719):
to_process = [idx for idx, row in matched_df.iterrows() 
              if quick_filter_requires_audit(row) and row.get('Accuracy_Status') not in TERMINAL_STATUSES]
```

---

## Fix #4: Add Sampling for Very Large Datasets

If you have thousands of vehicles, sample strategically instead of auditing all:

```python
# Add before run_pipeline():

SAMPLE_FRACTION = float(os.environ.get('SAMPLE_FRACTION', '1.0'))  # 1.0 = 100%, 0.1 = 10%

if SAMPLE_FRACTION < 1.0:
    import random
    to_process = random.sample(to_process, int(len(to_process) * SAMPLE_FRACTION))
    print(f"Sampling {len(to_process)} vehicles ({SAMPLE_FRACTION*100:.0f}% of {len(matched_df)})")
```

---

## Recommended GitHub Actions Config Changes

Create/update `.github/workflows/` variables:

```yaml
SAMPLE_FRACTION: "0.2"                 # Start with 20% of vehicles
DAILY_REQUEST_CAP: "50"
RATE_LIMIT_RPM: "30"
WEB_SEARCH_RATE_LIMIT_RPM: "12"
MAX_TEXT_CHARS: "12000"
GROUP_BATCH_SIZE: "8"
MAX_RUNTIME_MINUTES: "45"
WEB_SEARCH_GIVE_UP_THRESHOLD: "5"      # Give up faster on free tier
```

---

## Expected Results After Fixes

**Before**: 
- 6 keys, 1 quota pool → ~10 requests/day total → 391 failures
- Only 88 vehicles verified

**After**:
- 6 keys, 6 quota pools → ~60 requests/day total → Can audit 50-100 vehicles/day
- 6-8 vehicles per day with full accuracy
- Web search fallback available for 15-20 no-brochure vehicles/day

---

## Timeline to Deploy

1. **Day 1**: Create 6 new projects + get 6 new keys (~30 min)
2. **Day 1**: Update GitHub secrets (5 min)
3. **Day 1 PM**: Test with 1 run (~10 min)
4. **Day 2**: If working, enable on schedule; otherwise iterate

---

## If You Still Want Free + Need Higher Throughput

Your **only** other option without paying is:
- Spread runs across **multiple days** (run audit 3x/week instead of daily)
- **Focus sampling**: Audit only high-value vehicles (new models, popular makes)
- **Rotate projects**: Use a different set of 3 projects each run so quotas reset naturally

---

## Questions?

If any step fails or quotas still exhaust, share:
1. Screenshot of Google Cloud Console showing all 6 projects + their API keys
2. The output from a test run after deployment
3. Any error messages from the GitHub Actions log
