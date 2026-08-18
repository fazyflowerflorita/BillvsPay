# Fix: 502 Bad Gateway Error

## The Problem

The error message shows:
- ❌ **502 Bad Gateway** - Flask app crashed
- **"Failed to load resource: the server responded with a status of 502"**

This means the Python backend is not running or crashed on startup.

---

## Root Causes

The most likely reasons:

1. **reconciliation.py has a syntax error**
2. **Missing or wrong imports in server.py**
3. **Dependencies not installed correctly**
4. **Module import failed at startup**

---

## Solution: Use Minimal Server

I've created `server_minimal.py` which:

✅ Delays imports (imports happen during request, not startup)
✅ Better error handling and logging
✅ Won't crash even if imports fail
✅ Provides detailed error messages

### Steps:

1. **Download:** `server_minimal.py`
2. **Rename:** to `server.py`
3. **Upload to GitHub:** Replace `app/server.py`
4. **Render auto-deploys** in 2-3 minutes

---

## What's Different

### Old (Crashes on startup):
```python
# ❌ This fails if import errors occur
from reconciliation import BillPayReconciler
import pandas as pd

app = Flask(__name__)
# If imports fail → 502 error
```

### New (Safe):
```python
# ✅ Imports only when needed
app = Flask(__name__)

@app.route('/reconcile', methods=['POST'])
def reconcile():
    try:
        import pandas as pd  # Import here, not at startup
        from reconciliation import BillPayReconciler
        # ... rest of code ...
    except ImportError as e:
        return f'Import error: {e}', 500
```

---

## Debugging Steps

If you still get errors after updating:

### Step 1: Check Render Logs

Go to https://dashboard.render.com:

1. Select your service
2. Click "Logs"
3. Look for error messages like:
   - `ImportError: No module named 'reconciliation'`
   - `SyntaxError in server.py`
   - `ModuleNotFoundError`

### Step 2: Verify File Structure

Your GitHub should have:

```
app/
├── server.py              ← Use server_minimal.py
├── reconciliation.py      ← Must exist and be valid Python
└── __init__.py           ← Empty file
```

**Check the files exist:**
```bash
git ls-files | grep app/
```

Should show:
```
app/__init__.py
app/reconciliation.py
app/server.py
```

### Step 3: Check reconciliation.py Syntax

The reconciliation.py must be valid Python. To test locally:

```bash
python3 -m py_compile reconciliation.py
```

If there are syntax errors, it will tell you the line number.

### Step 4: Verify requirements.txt

Make sure your `requirements.txt` has:

```
pandas==2.1.3
openpyxl==3.1.2
Flask==2.3.3
gunicorn==21.2.0
```

NOT:
```
pandas==3.0.1  # ❌ Too new, unstable
```

---

## Quick Fix Checklist

- [ ] Download `server_minimal.py`
- [ ] Rename to `server.py`
- [ ] Upload to GitHub `app/` folder
- [ ] Check Render logs for errors
- [ ] Wait 2-3 minutes for Render to rebuild
- [ ] Refresh page (Ctrl+Shift+R)
- [ ] Try upload again

---

## If Error Persists

### Option 1: Check Render Logs

```
Render Dashboard → Your Service → Logs
```

Look for error messages like:
- `ImportError: cannot import name 'BillPayReconciler'`
- `ModuleNotFoundError: No module named 'reconciliation'`
- `SyntaxError: invalid syntax`

### Option 2: Roll Back

If you're not sure what went wrong:

1. Go back to the original `server.py`
2. Upload to GitHub
3. Render rebuilds (2-3 min)
4. Then try again carefully

### Option 3: Manual File Check

Make sure these files are in your repo:

1. `requirements.txt` - has correct versions
2. `app/server.py` - valid Python (no syntax errors)
3. `app/reconciliation.py` - valid Python (no syntax errors)
4. `app/__init__.py` - exists (can be empty)
5. `render.yaml` - has correct config

---

## Common Causes & Fixes

| Error | Cause | Fix |
|-------|-------|-----|
| 502 Bad Gateway | Server crashed | Use server_minimal.py |
| ImportError: reconciliation | File not uploaded | Check GitHub has app/reconciliation.py |
| SyntaxError | Invalid Python | Check reconciliation.py for errors |
| ModuleNotFoundError: pandas | Version mismatch | Update requirements.txt |

---

## Expected Behavior After Fix

1. ✅ Home page loads (no errors)
2. ✅ Can select files without errors
3. ✅ Progress bar appears when uploading
4. ✅ Progress updates every 500ms
5. ✅ Excel file downloads after reconciliation
6. ✅ No 502 errors or JSON errors

---

## Next Steps

1. **Immediate:** Download and upload `server_minimal.py`
2. **Wait:** Render rebuilds (2-3 minutes)
3. **Test:** Try uploading files again
4. **Check Logs:** If still fails, check Render logs for detailed error

**The minimal server is designed to be bulletproof - it should work!**
