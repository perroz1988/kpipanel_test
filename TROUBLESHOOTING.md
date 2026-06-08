# KPI Dashboard Update — Troubleshooting Guide

**Se qualcosa non funziona dopo l'aggiornamento, consulta questo file.**

---

## 🔴 CRITICAL ERRORS

### Error #1: "ValueError: not enough values to unpack"
**When**: Durante `python3 update.py`

**Cause**: CSV ha righe con meno di 20 celle (parser aspetta almeno 20)

**Solution**:
```bash
# Verifica il numero di colonne nei CSV
head -1 inbox/creative_performance.csv | tr ',' '\n' | wc -l

# Deve essere ≥ 20 colonne
# Se < 20, il CSV è corrotto o formato diverso
```

**Fix**:
1. Apri CSV in Excel
2. Verifica che non ci siano righe vuote o mal formattate all'inizio
3. Conta le colonne: se < 20, contatta la fonte dati
4. Salva come CSV UTF-8 (non Excel native format)
5. Ripeti `python3 update.py`

---

### Error #2: "KeyError: 'post_text'" oppure "post_text not found"
**When**: Nel browser, sezione "Amplificazione per campagna" è vuota o mancante

**Cause**: Colonna post_text non estratta dal CSV durante parse

**Solution**:
```python
# Nel file update.py, riga ~1270, verifica che ci sia:
# text_lines = [cells[i] for i in range(len(cells)) if cells[i]]

# Se manca, aggiungi manualmente in update.py intorno riga 1285:
'post_text': ' '.join(text_lines) if text_lines else ''
```

**Debug**:
1. Apri `rs_data.json` in browser
2. Premi F12 (DevTools)
3. Console: `CREATIVE_DATA[0]` → Enter
4. Verifica se ha campo `post_text`
5. Se vuoto o mancante → è un bug di parsing

**Fix**: 
- Verifica che CSV abbia colonna con testo del post
- Se sì, verifica indices in update.py (linee 1262-1270)
- Se no, post_text resterà vuoto (amplificazioni potrebbero non funzionare)

---

### Error #3: Impressioni/CTR/ER% totalmente diversi tra viste
**When**: Post mostra 1000 imp in Creative Paid, 500 in Post Detail

**Cause**: Indici delle colonne sbagliati nel CSV parser

**Solution**:
```python
# Verifica indices in update.py linee 1279-1292:
spent = float(l[6])           # Colonna 6 = Spesa?
impressions = float(l[7])     # Colonna 7 = Impressioni?
clicks = float(l[8])          # Colonna 8 = Click?
reactions = float(l[11])      # Colonna 11 = Reazioni?
comments = float(l[12])       # Colonna 12 = Commenti?
shares = float(l[13])         # Colonna 13 = Condivisioni?
```

**Debug**:
1. Apri `rs_data.json` nel browser
2. Apri primo record CREATIVE_DATA
3. Controlla: impressions, clicks, reactions, comments, shares presenti?
4. Se numeri sono troppo piccoli o 0 → indices sbagliati

**Fix**:
1. Apri CSV in Excel
2. Conta la posizione reale delle colonne (0-indexed):
   - Colonna A = indice 0
   - Colonna B = indice 1
   - etc.
3. Aggiorna gli indices nel file update.py
4. Ripeti `python3 update.py`

---

### Error #4: "Breakdown Organico vs Sponsorizzato" mostra valori diversi dalla riga post aggregata
**When**: Totale nel breakdown ≠ Totale nel Creative Paid

**Cause**: breakdownTable() non sta calcolando correttamente oppure dati incompleti

**Solution**:
```javascript
// File dashboard.html, riga 1797-1821
// Verifica che la funzione calcoli così:

const totale = {
  impressioni: (org?.impressioni||0) + (spon?.impressioni||0),
  clic: (org?.clic||0) + (spon?.clic||0),
  reazioni: (org?.reazioni||0) + (spon?.reazioni||0),
  commenti: (org?.commenti||0) + (spon?.commenti||0),
  diffusioni: (org?.diffusioni||0) + (spon?.diffusioni||0),
}
```

**Debug**:
1. Apri Developer Tools (F12)
2. Vai al singolo post
3. Scrolla fino al "Breakdown"
4. Confronta i numeri riga per riga

**Fix**:
- Se sono uguali → OK (non è un bug)
- Se diversi → file dashboard.html corrotto
  - Revertire a versione precedente: `git checkout HEAD -- dashboard.html`
  - Poi esegui commit più recente

---

## 🟡 MEDIUM PRIORITY ISSUES

### Issue #5: Collapse UI non funziona (clicchi post, scomparisce solo primo)
**Symptom**: Clicchi riga post aggregata, scompaiono solo le prime creative, altre restano visibili

**Cause**: querySelector non trova tutte le righe (probabilmente usando id invece di class)

**Solution**:
Verifica dashboard.html righe **1496, 1593, 1513, 1612**:
```javascript
// CORRETTO: usa querySelectorAll con class
const els = document.querySelectorAll('tr.${id}');
els.forEach(el => el.style.display = show ? 'table-row' : 'none');

// SBAGLIATO: usa getElementById (tocca solo primo elemento)
const el = document.getElementById('${id}');
el.style.display = el.style.display === 'none' ? 'table-row' : 'none';
```

**Fix**:
```bash
# Verifica che dashboard.html usi querySelectorAll
grep -n "querySelectorAll('tr\." dashboard.html

# Deve trovare risultati alla riga 1496, 1593
# Se non li trova, il file è stato rollback accidentalmente

# Fix: reimplementa il collapse fix
```

---

### Issue #6: "Solo Sponsorizzati" non include post organici amplificati
**Symptom**: Vedi un post organico nel tab Organici, ma NON appare in Sponsorizzati

**Cause**: Substring matching fallito (post_text ≠ titolo post)

**Solution**:
Verifica dashboard.html linee **1352-1377** (renderContenuti):
```javascript
const normalize = t => (t||'').replace(/\s+/g,' ').toLowerCase().trim();
const pNorm = normalize(p.titolo);
return (CREATIVE_DATA||[]).some(c => {
  const cNorm = normalize(c.post_text || '');
  // Matching: primi 30 caratteri
  return cNorm.includes(pNorm.substring(0,30)) || 
         pNorm.includes(cNorm.substring(0,30));
});
```

**Debug**:
1. Apri DevTools (F12)
2. Console:
```javascript
const p = RS_DATA.posts.find(x => x.titolo.includes('Senza misurare'));
const pNorm = p.titolo.replace(/\s+/g,' ').toLowerCase().trim();
const c = CREATIVE_DATA.find(x => x.post_text);
const cNorm = (c.post_text || '').replace(/\s+/g,' ').toLowerCase().trim();

console.log('Post:', pNorm.substring(0,30));
console.log('Creative:', cNorm.substring(0,30));
console.log('Match:', cNorm.includes(pNorm.substring(0,30)));
```

**Fix**:
- Se Match = false: aumenta da 30 → 50 caratteri
- Se Match = true ma post non appare: problema in renderContenuti logic

---

### Issue #7: CTR% o ER% sono decimali strani (0.05% invece di 5%)
**Symptom**: CTR mostra "0.05%" quando dovrebbe essere "5%"

**Cause**: Moltiplicazione doppia per 100 o formula sbagliata

**Solution**:
Verifica le formule:
```javascript
// CORRETTO:
const ctr = agg.imp>0 ? nfDec.format(agg.clicks/agg.imp*100) : '—';
// Risultato: 5 → fmtPct lo converte a "5.00%"

// SBAGLIATO:
const ctr = agg.imp>0 ? nfDec.format((agg.clicks/agg.imp)*100) : '—';
// Moltiplicato due volte (una in formula, una in fmtPct)
```

**Debug**:
1. DevTools Console:
```javascript
const c = CREATIVE_DATA[0];
console.log('clicks:', c.clicks, 'imp:', c.impressions);
console.log('Manual CTR:', (c.clicks/c.impressions*100).toFixed(2) + '%');

// Confronta con quello che vedi nel browser
```

**Fix**:
- Controlla che NON ci sia moltiplicazione extra per 100 dopo il calcolo
- Assicura che fmtPct() riceva un numero 0-100, non 0-1

---

## 🟢 MINOR ISSUES

### Issue #8: Column headers non si allineano con i dati
**Symptom**: Header dice "Campagna" ma colonna mostra logo/immagine

**Cause**: Numero di colonne diverso tra header e data rows

**Solution**:
Conta colonne in:
- Header (dashboard.html linea 633-645): 10 colonne
- Data (dashboard.html renderCreatives): 10 colonne

Devono essere **identici nell'ordine**:
1. Campagna
2. Ad/Creative
3. Impressioni
4. Reach
5. Click
6. CTR %
7. Engagement
8. ER %
9. Spesa
10. Lead

**Fix**:
Se diversi, allinea entrambi (header e data cells devono corrispondere 1:1)

---

### Issue #9: Post detail "Metriche chiave (Totale)" mostra Organico invece di Totale
**Symptom**: Nel singolo post vedi solo badge "Organico", non "Totale"

**Cause**: Aggregazione Organico+Sponsorizzato non funziona o dati mancanti

**Solution**:
Verifica dashboard.html linee **1653-1673** (renderPostDetail):
```javascript
let post = matches.find(p=>p.tipo==='Totale');

if(!post && (orgMatch || sponMatch)){
  // Aggregazione automatica
  const org = orgMatch || {...default empty...};
  const spon = sponMatch || {...default empty...};
  const totImp = org.impressioni + spon.impressioni;
  const totClick = (org.clic||0) + (spon.clic||0);
  
  post = {
    ...org,
    tipo: 'Totale',
    impressioni: totImp,
    clic: totClick,
    // etc...
  };
}
```

**Debug**:
1. Apri singolo post
2. DevTools: `RS_DATA.posts.filter(p => p.link === window.location.hash)`
3. Verifica se c'è un record con `tipo: 'Totale'`
4. Se no → l'aggregazione dovrebbe crearla

**Fix**:
- Se aggregazione non funziona → controlla che Organico E Sponsorizzato esistono nei dati
- Se solo uno esiste → dovrebbe comunque aggregarsi con valori 0 per l'altro

---

### Issue #10: Chevron indicators (▼/▶) non cambiano
**Symptom**: Clicchi post ma chevron resta sempre ▼

**Cause**: JavaScript non esegue il toggle del testo

**Solution**:
Verifica che onclick contenga:
```javascript
chev.textContent = show ? '▼' : '▶';
```

**Debug**:
1. DevTools F12
2. Clicca riga post
3. Ispeziona elemento (right-click → Inspect)
4. Verifica che chevron abbia id corretto: `id="${chevronId}"`

**Fix**:
Se l'elemento non ha l'id corretto, il toggle non funziona. Verificare che la riga sia renderizzata correttamente.

---

## 📋 Quick Diagnostic Checklist

Se qualcosa è rotto, esegui questi check in ordine:

```bash
# 1. Verifica che il file sia stato aggiornato
git log --oneline -5

# 2. Verifica che l'ultimo commit sia un update, non un revert
git show HEAD

# 3. Controlla se ci sono errori nel browser (F12 → Console)
# Dovrebbe NON avere errori rossi

# 4. Verifica che CREATIVE_DATA sia popolato
# Console: CREATIVE_DATA.length (dovrebbe essere > 0)

# 5. Verifica che RS_DATA sia popolato
# Console: RS_DATA.posts.length (dovrebbe essere > 0)

# 6. Se dati sono 0, il CSV parsing fallito
# Ricontrolla update.py output durante run
```

---

## 🆘 Se Niente Funziona

**Last Resort Actions** (in ordine di "rischio"):

### Safe Actions
1. **Hard refresh browser**: Ctrl+Shift+R (cancella cache)
2. **Controlla console errors**: F12 → Console → leggi errori rossi
3. **Verifica CSV input**: `head inbox/creative_performance.csv`

### Risky Actions (solo se necessario)
4. **Revert ultimo commit**: `git revert HEAD`
5. **Revert all changes**: `git reset --hard HEAD~1`
6. **Manuale fix**: Modifica file e re-deploy

### Nuclear Option (solo se disperato)
7. Contatta: Non c'è contatto nel file, ma il repository è versionato — controlla `git log` per storia completa

---

## 📞 Quick Reference

| Problem | First Check | File | Line |
|---------|------------|------|------|
| ER%/CTR% wrong | CSV indices | update.py | 1279-1292 |
| Missing post_text | post_text extraction | update.py | 1262-1270 |
| Collapse broken | querySelectorAll usage | dashboard.html | 1496, 1593, 1513, 1612 |
| Amplificazione missing | substring matching | dashboard.html | 1352-1377 |
| Breakdown different | aggregation logic | dashboard.html | 1797-1821 |
| Columns misaligned | header vs data | dashboard.html | 633-645 vs renderCreatives |
| Totale wrong | aggregation | dashboard.html | 1653-1673 |

---

**Last Updated**: 2026-06-08  
**Version**: 1.0  
**Status**: Ready for troubleshooting
