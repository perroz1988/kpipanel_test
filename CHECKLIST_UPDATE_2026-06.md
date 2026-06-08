# KPI Dashboard Update Checklist — Giugno 2026

**Data Update**: [DATA AGGIORNAMENTO - aggiorna quando esegui]
**Verificato da**: [TUO NOME]
**Status**: ⏳ In Attesa

---

## PRE-UPDATE: Verifica Files

### CSV Input Verification
- [ ] `inbox/rs-italia-organico.csv` presente
- [ ] `inbox/rs-italia-paid.csv` presente  
- [ ] `inbox/creative_performance.csv` presente
- [ ] Tutti i file hanno righe con dati (non vuoti)

### CSV Structure Check
```
Creative CSV deve avere:
- Colonna 6: spent (Spesa)
- Colonna 7: impressions (Impressioni)
- Colonna 8: clicks (Clic)
- Colonna 11: reactions (Reazioni)
- Colonna 12: comments (Commenti)
- Colonna 13: shares (Condivisioni)
- post_text field popolato per amplificazioni
```
- [ ] Apri un file CSV in Excel → verifica colonne nel posto giusto
- [ ] Verifica che almeno un post abbia post_text (non vuoto)

---

## RUN UPDATE

```bash
cd '/Users/francescoperrotta/Dropbox/Fannel_Design/ClaudePrj/KPI Dashboard'
python3 update.py
```

**Aspetta il completamento** (stampa "✓ Aggiornamento completato")

---

## POST-UPDATE: Verification (15 min)

### 1️⃣ Creative Paid Table
Apri Dashboard → Tab "Creative Paid — Campagne Krein"

- [ ] **Header corretto**: Campagna | Ad/Creative | Impressioni | Reach | Click | CTR % | Engagement | ER % | Spesa | Lead
- [ ] **10 colonne visibili** (non tagliato)
- [ ] **Post grouping funziona**:
  - [ ] Clicca su un post aggregato (riga grigia)
  - [ ] ✓ Chevron cambia (▼ → ▶)
  - [ ] ✓ TUTTE le creative sotto scompaiono insieme
- [ ] **Metriche corrette su un post random**:
  - [ ] Impressioni = somma creative sotto
  - [ ] CTR % = clicks totali / impressioni totali × 100
  - [ ] Engagement = reactions + comments + shares
  - [ ] ER % = engagement / impressioni × 100

**Esempio**: Se post ha 2 creative:
```
Post: 1000 imp, 50 click, 100 engagement → CTR=5%, ER=10%
  Creative1: 600 imp, 30 click, 60 eng
  Creative2: 400 imp, 20 click, 40 eng
```

### 2️⃣ Singolo Post Detail
- [ ] Clicca su un post qualsiasi dalla lista
- [ ] Sezione "Metriche chiave (Totale)" mostra:
  - [ ] ✓ Impressioni corrette (Organico + Amplificati)
  - [ ] ✓ ER % allineato con Breakdown
  - [ ] ✓ CTR % calcolato correttamente

### 3️⃣ Breakdown Organico vs Sponsorizzato vs Totale
Nel singolo post, sezione "Breakdown...":

- [ ] **Tre righe presenti**: Organico | Sponsorizzato | Totale
- [ ] **Totale = Organico + Sponsorizzato**:
  - Impressioni(Tot) = Impressioni(Org) + Impressioni(Spon)
  - ER(Tot) = (reactions+comments+shares)_totale / impressioni_totale × 100
  - CTR(Tot) = clicks_totale / impressioni_totale × 100
- [ ] ✓ Numeri coerenti in tutte le righe

### 4️⃣ Amplificazione Per Campagna
Nel singolo post, sezione "📢 Amplificazione per campagna":

- [ ] ✓ Mostra post amplificati (se esistono)
- [ ] ✓ Numero di creatività corretto
- [ ] ✓ Impressioni, ER%, CTR% calcolati
- [ ] ✓ Reazioni | Commenti | Condivisioni visibili

### 5️⃣ Contenuti Organici Tab
- [ ] **Collapsible groups funzionano**:
  - [ ] Clicca "Organico" → tutte le righe scompaiono (▼ → ▶)
  - [ ] Clicca "Sponsorizzato" → tutte le righe scompaiono
  - [ ] Clicca "Totale" → tutte le righe scompaiono
- [ ] **"Solo Sponsorizzati" filter**:
  - [ ] Include post organici con amplificazioni ✓
  - [ ] Non duplica post base

### 6️⃣ Data Consistency Check
**Prendi un post: "Senza misurare, non si ottimizza..."**

| Metrica | Creative Paid | Post Detail | Breakdown Totale | ✓ Match? |
|---------|--------------|-------------|------------------|---------|
| Impressioni | [num] | [num] | [num] | [ ] |
| Engagement | [num] | [num] | [num] | [ ] |
| ER % | [num]% | [num]% | [num]% | [ ] |
| CTR % | [num]% | [num]% | [num]% | [ ] |

Se tutti i numeri sono identici → ✅ PERFETTO

### 7️⃣ Edge Cases
- [ ] Post con SOLO Organico (nessuna amplificazione) → no errors
- [ ] Post con SOLO Sponsorizzato → Organico row = 0
- [ ] Periodo vuoto → "Nessun post nel periodo" message

---

## ❌ If Something's Wrong

### ⚠️ Wrong ER% or CTR%
**Sintomo**: Numeri diversi in Creative Paid vs Post Detail

**Diagnostica**:
1. Apri DevTools (F12)
2. Console → cercaquery `CREATIVE_DATA[0]`
3. Verifica: reactions, comments, shares, clicks, impressions presenti?
4. Se mancano campi → problema nel CSV parser (update.py)

**Fix**: Verifica indices in update.py linee 1279-1292

### ⚠️ Collapse not working
**Sintomo**: Clicchi su un post, scomparisce solo la prima creative

**Causa**: querySelector non funziona (class attribute non impostato)

**Fix**: Verifica dashboard.html:1513, 1612 usano `class="${id}"` non `id="${id}"`

### ⚠️ Amplificated posts not in sponsored filter
**Sintomo**: Organico posts non appare in "Solo Sponsorizzati"

**Causa**: Substring matching fallito (post_text format cambiato)

**Fix**: Verifica post_text in CREATIVE_DATA è simile ai titoli in RS_DATA

### ⚠️ Table columns misaligned
**Sintomo**: Colonne non si allineano tra header e dati

**Causa**: Headers (633-645) != data columns (renderCreatives)

**Fix**: Conta colonne in entrambi, assicura 10 colonne in ordine identico

---

## Final Sign-Off

- [ ] Creative Paid table ✓
- [ ] Post detail aggregation ✓
- [ ] Breakdown calculations ✓
- [ ] Amplificazione visibile ✓
- [ ] Collapse UI funziona ✓
- [ ] Zero data inconsistencies ✓
- [ ] Dashboard performante ✓

**Status**: ✅ READY FOR PRODUCTION

**Last verified**: [INSERISCI DATA/ORA]

---

**Questions?** Referencealla documentazione completa in: `/Users/francescoperrotta/.claude/projects/.../memory/kpi_dashboard_fixes_2026.md`
