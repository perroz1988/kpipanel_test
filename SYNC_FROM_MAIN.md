# Test Environment — Sync from Main

**Date Synced**: 2026-06-08  
**Status**: ✅ Ready for Testing

---

## Files Synced from Main

| File | Purpose | Last Updated |
|------|---------|--------------|
| `dashboard.html` | UI fixes (Creative Paid, collapse, aggregation) | 2026-06-08 |
| `update.py` | CSV parser fixes (cell count, post_text) | 2026-06-08 |
| `CHECKLIST_UPDATE_2026-06.md` | 7-step verification guide | 2026-06-08 |
| `TROUBLESHOOTING.md` | Error troubleshooting reference | 2026-06-08 |

---

## What Was Fixed

✅ CSV parser recognizes 22-cell rows (Always On campaigns)  
✅ Post text extraction from multi-line CSV rows  
✅ Creative Paid table: 10 columns, proper metrics (ER, CTR)  
✅ Post detail: aggregates Organico + Sponsorizzato when Totale missing  
✅ Breakdown: calculates correct Totale (never from raw data)  
✅ Collapse UI: toggles ALL rows of a type together (not just first)  
✅ Chevron indicators (▼/▶) show collapse state  
✅ Amplified posts included in Sponsorizzati filter  

---

## Test This Environment

### Before First Use
```bash
# Make sure you have latest Python packages
pip install pandas

# Test the parser
python3 update.py
```

### After `python3 update.py`
1. Open `dashboard.html` in browser
2. Use **CHECKLIST_UPDATE_2026-06.md** to verify (7 steps, 15 min)
3. If any issue → check **TROUBLESHOOTING.md**

---

## How This Was Created

Synced from **Main KPI Dashboard** repository:
- https://github.com/perroz1988/kpipanel  
- Branch: `main`  
- Commit: All recent fixes applied

This test environment is now **in sync** with production main. You can test confidently!

---

## Next Steps

1. **Upload test CSV** to `inbox/`
2. **Run update**: `python3 update.py`
3. **Verify**: Use CHECKLIST_UPDATE_2026-06.md
4. **Report findings**: Any issues? Check TROUBLESHOOTING.md

---

**Ready to test!** 🚀
