#!/usr/bin/env python3
"""
update.py — Aggiorna le dashboard KPI (RS Italia + Optimedia).

━━━ STRUTTURA INBOX ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  inbox/
    rs-italia/          ← XLS LinkedIn + CSV campagne
    optimedia/
      Facebook/         ← CSV Meta Business Suite Facebook
      Instagram/        ← CSV Meta Business Suite Instagram

━━━ RS ITALIA (ogni settimana) ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  inbox/rs-italia/  ← content.xls, followers.xls, campagne.csv

━━━ OPTIMEDIA (ogni settimana/mese) ════════════════════════════
  inbox/optimedia/Facebook/   ← Views.csv, Viewers.csv,
                                 Interactions.csv, Follows.csv,
                                 Link clicks.csv, Visits.csv,
                                 Audience FB.csv
  inbox/optimedia/Instagram/  ← Views.csv, Reach.csv,
                                 Interactions.csv, Follows.csv,
                                 Link clicks.csv, Visits.csv,
                                 Audience IG.csv

━━━ LANCIA ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
     python3 update.py

━━━ FLAGS ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  --no-keye      salta fetch Keye (usa ultimo archivio)
  --no-inbox     salta scansione inbox (usa solo archivio)
  --rs-only      aggiorna solo RS Italia
  --opt-only     aggiorna solo Optimedia

Richiede: pip3 install xlrd
"""

import xlrd
import json
import re
import os
import sys
import glob
import shutil
import subprocess
import urllib.request
from collections import defaultdict
from datetime import datetime

BASE = os.path.dirname(os.path.abspath(__file__))

# ── RS ITALIA ────────────────────────────────────────────────────────────────
ARCHIVE_RS       = os.path.join(BASE, 'archive', 'rs-italia')
ARCHIVE          = os.path.join(BASE, 'archive', 'rs-italia', 'linkedin')
ARCHIVE_KEYE     = os.path.join(BASE, 'archive', 'rs-italia', 'keye')
ARCHIVE_CAMPAGNE = os.path.join(BASE, 'archive', 'rs-italia', 'campagne')
ARCHIVE_PDF      = os.path.join(BASE, 'archive', 'pdf')
ARCHIVE_ROOT     = os.path.join(BASE, 'archive', 'rs-italia')
TRASH            = os.path.join(BASE, '_trash')
INBOX_RS         = os.path.join(BASE, 'inbox', 'rs-italia')
INBOX            = INBOX_RS   # alias legacy
DASHBOARD        = os.path.join(BASE, 'dashboard.html')
HISTORY_JSON     = os.path.join(BASE, 'rs_history.json')
CAMP_HISTORY_JSON= os.path.join(BASE, 'camp_history.json')
CAMP_FILTER_KEYWORDS = ['krein']

# ── OPTIMEDIA ────────────────────────────────────────────────────────────────
INBOX_OPT        = os.path.join(BASE, 'inbox', 'optimedia')
INBOX_OPT_FB     = os.path.join(BASE, 'inbox', 'optimedia', 'Facebook')
INBOX_OPT_IG     = os.path.join(BASE, 'inbox', 'optimedia', 'Instagram')
ARCHIVE_OPT_FB   = os.path.join(BASE, 'archive', 'optimedia', 'facebook')
ARCHIVE_OPT_IG   = os.path.join(BASE, 'archive', 'optimedia', 'instagram')
ARCHIVE_OPT      = os.path.join(BASE, 'archive', 'optimedia')
DASHBOARD_OPT    = os.path.join(BASE, 'optimedia.html')

KEYE_SUMMARY_URL = 'https://brand-listening.vercel.app/api/brand-summary/rs-italia?refresh=1'
KEYE_EXPORT_URL  = 'https://brand-listening.vercel.app/api/export/rs-italia'
KEYE_API_KEY     = '3bf7abb472ff1acf31e3e10a19a8779184bc22ba59de91e2'  # legacy
KEYE_COOKIE      = 'kpi6-session=kpi6-secret-2026'


# ─── DATE UTILS ──────────────────────────────────────────────────────────────

def to_iso(val):
    """MM/DD/YYYY string o numero xlrd → YYYY-MM-DD, oppure None."""
    if not val:
        return None
    s = str(val).strip()
    for fmt in ('%m/%d/%Y', '%Y-%m-%d'):
        try:
            return datetime.strptime(s, fmt).strftime('%Y-%m-%d')
        except ValueError:
            pass
    return None


# ─── PARSER XLS ──────────────────────────────────────────────────────────────

def parse_content(path):
    """Estrae metriche giornaliere e lista post da un content XLS di LinkedIn."""
    wb = xlrd.open_workbook(path)

    # ---- Metriche (riga 0 = descrizione lunga, riga 1 = headers, riga 2+ = dati)
    metriche = []
    sheet = wb.sheet_by_name('Metriche')
    for r in range(2, sheet.nrows):
        row = sheet.row_values(r)
        date = to_iso(row[0])
        if not date:
            continue
        def iv(i): return int(row[i]) if len(row) > i and row[i] != '' else 0
        def fv(i): return float(row[i]) if len(row) > i and row[i] != '' else 0.0
        metriche.append({
            'data': date,
            'impressioni_org':        iv(1),
            'impressioni_spon':       iv(2),
            'impressioni_tot':        iv(3),
            'impressioni_uniche_org': iv(4),
            'clic_org':               iv(5),
            'clic_spon':              iv(6),
            'clic_tot':               iv(7),
            'reazioni_org':           iv(8),
            'reazioni_spon':          iv(9),
            'reazioni_tot':           iv(10),
            'commenti_org':           iv(11),
            'commenti_spon':          iv(12),
            'commenti_tot':           iv(13),
            'diffusioni_org':         iv(14),
            'diffusioni_spon':        iv(15),
            'diffusioni_tot':         iv(16),
            'interesse_org':          fv(17),
            'interesse_spon':         fv(18),
            'interesse_tot':          fv(19),
        })

    # ---- Post (riga 0 = descrizione, riga 1 = headers, riga 2+ = dati)
    posts = []
    sheet2 = wb.sheet_by_name('Tutti i post')
    for r in range(2, sheet2.nrows):
        row = sheet2.row_values(r)
        if not row[0]:
            continue
        def sv(i): return str(row[i]).strip() if len(row) > i and row[i] != '' else None
        def iv(i): return int(row[i]) if len(row) > i and row[i] != '' else 0
        def fv(i): return float(row[i]) if len(row) > i and row[i] != '' else 0.0
        posts.append({
            'titolo':          sv(0),
            'link':            sv(1),
            'tipo':            sv(2),
            'campagna':        sv(3),
            'pubblicato_da':   sv(4),
            'data_creazione':  to_iso(sv(5)),
            'data_inizio_camp': to_iso(sv(6)),
            'data_fine_camp':  to_iso(sv(7)),
            'pubblico':        sv(8) or '',
            'impressioni':     iv(9),
            'visualizzazioni': iv(10),
            'vis_offsite':     iv(11),
            'clic':            iv(12),
            'ctr':             fv(13),
            'reazioni':        iv(14),  # "Volte consigliato"
            'commenti':        iv(15),
            'diffusioni':      iv(16),
            'nuovi_follower':  iv(17),
            'interesse':       fv(18),
            'tipo_contenuto':  sv(19),
        })

    return metriche, posts


def parse_followers(path):
    """Estrae follower giornalieri e demografici da un followers XLS di LinkedIn."""
    wb = xlrd.open_workbook(path)

    # ---- Nuovi follower
    followers_daily = []
    sheet = wb.sheet_by_name('Nuovi follower')
    headers = [str(v).strip() for v in sheet.row_values(0)]
    has_totale = 'Follower totali' in headers

    for r in range(1, sheet.nrows):
        row = sheet.row_values(r)
        date = to_iso(row[0])
        if not date:
            continue
        spon = int(row[1]) if row[1] else 0
        org  = int(row[2]) if row[2] else 0
        auto = int(row[3]) if row[3] else 0
        tot  = int(row[4]) if has_totale and len(row) > 4 and row[4] else spon + org + auto
        followers_daily.append({'data': date, 'spon': spon, 'org': org, 'auto': auto, 'tot': tot})

    def parse_kv(sheet_name, key_col=0, val_col=1):
        s = wb.sheet_by_name(sheet_name)
        out = []
        for r in range(1, s.nrows):
            row = s.row_values(r)
            if len(row) > val_col and row[key_col] and row[val_col] != '':
                out.append({'label': str(row[key_col]).strip(), 'value': int(row[val_col])})
        return out

    demographics = {
        'localita':   parse_kv('Località'),
        'funzione':   parse_kv('Funzione lavorativa'),
        'anzianita':  parse_kv('Anzianità'),
        'settore':    parse_kv('Settore'),
        'dim_azienda': parse_kv("Dimensioni dell’azienda"),
    }

    fan_base_totale = sum(x['value'] for x in demographics['anzianita'])
    return followers_daily, demographics, fan_base_totale


# ─── ARCHIVIO ────────────────────────────────────────────────────────────────

def archive_new_file(src_path, kind):
    """
    Legge il file, rileva la data di fine periodo e lo copia in archive/linkedin/
    con il nome YYYY-MM-DD_{kind}.xls. Restituisce il path archiviato.
    kind: 'content' | 'followers'
    """
    wb = xlrd.open_workbook(src_path)
    if kind == 'content':
        sheet = wb.sheet_by_name('Metriche')
        dates = [to_iso(sheet.row_values(r)[0]) for r in range(2, sheet.nrows)]
    else:
        sheet = wb.sheet_by_name('Nuovi follower')
        dates = [to_iso(sheet.row_values(r)[0]) for r in range(1, sheet.nrows)]

    dates = [d for d in dates if d]
    if not dates:
        raise ValueError(f"Nessuna data trovata in {src_path}")

    end_date = sorted(dates)[-1]
    dst = os.path.join(ARCHIVE, f'{end_date}_{kind}.xls')

    if os.path.exists(dst):
        ts = datetime.now().strftime('%H%M%S')
        dst = os.path.join(ARCHIVE, f'{end_date}_{kind}_{ts}.xls')

    shutil.copy2(src_path, dst)
    print(f'  Archiviato: {os.path.basename(dst)}')
    return dst


# ─── BUILD RS_DATA ───────────────────────────────────────────────────────────

def find_week_pairs():
    """Restituisce lista di (date_str, content_path, followers_path) ordinate per data."""
    content_files = sorted(glob.glob(os.path.join(ARCHIVE, '????-??-??_content*.xls')))
    pairs = []
    for cf in content_files:
        date = os.path.basename(cf)[:10]
        # Cerca il followers corrispondente per la stessa data
        ff_candidates = sorted(glob.glob(os.path.join(ARCHIVE, f'{date}_followers*.xls')))
        if ff_candidates:
            pairs.append((date, cf, ff_candidates[-1]))
    return pairs


def build_rs_data(pairs, n_weeks=2):
    """Combina i dati delle ultime n_weeks settimane."""
    selected = pairs[-n_weeks:] if len(pairs) >= n_weeks else pairs

    all_metriche = {}   # date → dict
    all_posts = {}      # link|tipo → dict
    all_followers = {}  # date → dict
    demographics      = None
    demographics_prev = None
    fan_base_totale      = 0
    fan_base_totale_prev = 0
    ultimo_aggiornamento = None

    # Itera in ordine cronologico: la prima è la più vecchia (prev), l'ultima è la corrente
    for i, (date, cf, ff) in enumerate(selected):
        print(f'  Leggo settimana {date}...')
        metriche, posts = parse_content(cf)
        fol, demo, fbt = parse_followers(ff)

        for m in metriche:
            all_metriche[m['data']] = m
        for p in posts:
            key = (p['link'] or '') + '|' + (p['tipo'] or '')
            all_posts[key] = p
        for f in fol:
            all_followers[f['data']] = f

        is_last = (i == len(selected) - 1)
        if is_last:
            demographics     = demo
            fan_base_totale  = fbt
            if metriche:
                ultimo_aggiornamento = sorted(m['data'] for m in metriche)[-1]
        else:
            demographics_prev     = demo
            fan_base_totale_prev  = fbt

    metriche_sorted  = sorted(all_metriche.values(), key=lambda x: x['data'])
    posts_sorted     = sorted(all_posts.values(), key=lambda x: x.get('data_creazione') or '')
    followers_sorted = sorted(all_followers.values(), key=lambda x: x['data'])

    data_min = metriche_sorted[0]['data']  if metriche_sorted  else ''
    data_max = metriche_sorted[-1]['data'] if metriche_sorted  else ''

    return {
        'meta': {
            'ultimo_aggiornamento': ultimo_aggiornamento or data_max,
            'data_min': data_min,
            'data_max': data_max,
            'fan_base_totale':      fan_base_totale,
            'fan_base_totale_prev': fan_base_totale_prev,
            'company_url': 'https://www.linkedin.com/company/rs-italia',
        },
        'metriche':           metriche_sorted,
        'posts':              posts_sorted,
        'followers_daily':    followers_sorted,
        'demographics':       demographics,
        'demographics_prev':  demographics_prev,
    }


# ─── STORICO INCREMENTALE ────────────────────────────────────────────────────

def update_history(pairs):
    """Aggiorna rs_history.json con tutti i file in archivio (incrementale).
    Legge lo storico esistente, aggiunge/sovrascrive solo le date nuove.
    """
    # Leggi storico esistente
    existing_metriche = {}
    existing_followers = {}
    if os.path.exists(HISTORY_JSON):
        try:
            with open(HISTORY_JSON, 'r', encoding='utf-8') as f:
                old = json.load(f)
            for m in old.get('metriche', []):
                existing_metriche[m['data']] = m
            for fol in old.get('followers_daily', []):
                existing_followers[fol['data']] = fol
        except Exception:
            pass

    # Leggi TUTTI i file in archivio (non solo n_weeks)
    for date, cf, ff in pairs:
        try:
            metriche, _ = parse_content(cf)
            followers, _, _ = parse_followers(ff)
            for m in metriche:
                existing_metriche[m['data']] = m
            for fol in followers:
                existing_followers[fol['data']] = fol
        except Exception as e:
            print(f'  WARN storico {date}: {e}')

    metriche_sorted  = sorted(existing_metriche.values(),  key=lambda x: x['data'])
    followers_sorted = sorted(existing_followers.values(), key=lambda x: x['data'])

    history = {
        'meta': {
            'aggiornato':  datetime.now().strftime('%Y-%m-%d'),
            'data_min':    metriche_sorted[0]['data']  if metriche_sorted  else '',
            'data_max':    metriche_sorted[-1]['data'] if metriche_sorted  else '',
            'giorni':      len(metriche_sorted),
        },
        'metriche':        metriche_sorted,
        'followers_daily': followers_sorted,
    }

    with open(HISTORY_JSON, 'w', encoding='utf-8') as f:
        json.dump(history, f, ensure_ascii=False, separators=(',', ':'))

    print(f'  Storico: {history["meta"]["data_min"]} → {history["meta"]["data_max"]} · {history["meta"]["giorni"]} giorni')
    return history


def _apply_krein_filter(camp_data):
    """Filtra camp_data ai soli Krein secondo CAMP_FILTER_KEYWORDS."""
    if not CAMP_FILTER_KEYWORDS:
        return camp_data
    def _match(name):
        n = name.lower()
        return any(kw.lower() in n for kw in CAMP_FILTER_KEYWORDS)
    ids_ok = {c['id'] for c in camp_data['camp_meta'] if _match(c['name'])}
    camp_data['camp_meta'] = [c for c in camp_data['camp_meta'] if c['id'] in ids_ok]
    camp_data['daily']     = [r for r in camp_data['daily'] if r.get('camp_id') in ids_ok]
    return camp_data


def rebuild_camp_history():
    """Ricostruisce camp_history.json da TUTTI i CSV in archive/campagne/.
    Scarta doppioni: per (camp_id, date), il CSV più recente vince."""
    rows_by_key = {}  # Key: (camp_id, date) → ultimo valore vince

    # Carica storico esistente
    if os.path.exists(CAMP_HISTORY_JSON):
        try:
            with open(CAMP_HISTORY_JSON, 'r', encoding='utf-8') as f:
                for r in json.load(f).get('daily', []):
                    key = (r['camp_id'], r['date'])
                    rows_by_key[key] = r
        except Exception:
            pass

    # Processa CSV in ordine cronologico (il CSV più recente sovrascrive)
    all_csvs = sorted(glob.glob(os.path.join(ARCHIVE_CAMPAGNE, '????-??-??_campagne.csv')))
    for csv_path in all_csvs:
        try:
            camp_data = _apply_krein_filter(parse_campaign_csv(csv_path))
            for r in camp_data['daily']:
                key = (r['camp_id'], r['date'])
                rows_by_key[key] = r  # CSV più recente sovrascrive
        except Exception as e:
            print(f'    WARN storico {os.path.basename(csv_path)}: {e}')

    sorted_rows = sorted(rows_by_key.values(), key=lambda x: x['date'])
    all_dates   = [r['date'] for r in sorted_rows]
    history = {
        'meta': {
            'aggiornato': datetime.now().strftime('%Y-%m-%d'),
            'data_min':   min(all_dates) if all_dates else '',
            'data_max':   max(all_dates) if all_dates else '',
            'righe':      len(sorted_rows),
        },
        'daily': sorted_rows,
    }
    with open(CAMP_HISTORY_JSON, 'w', encoding='utf-8') as f:
        json.dump(history, f, ensure_ascii=False, separators=(',', ':'))
    print(f'  Storico camp: {history["meta"]["data_min"]} → {history["meta"]["data_max"]} · {history["meta"]["righe"]} righe')
    return history


# ─── UPDATE HTML ─────────────────────────────────────────────────────────────

# ─── OPTIMEDIA ───────────────────────────────────────────────────────────────

def parse_meta_csv_series(path):
    """Legge un CSV Meta Business Suite (UTF-16, sep=,) con colonne Date,Primary."""
    with open(path, 'rb') as f:
        raw = f.read()
    text = raw.decode('utf-16', errors='replace')
    data = []
    in_data = False
    for line in text.splitlines():
        clean = line.replace('"', '').strip()
        if clean.startswith('Date'):
            in_data = True
            continue
        if in_data and clean and 'sep' not in clean.lower():
            parts = [p.strip().replace('"', '') for p in line.split(',')]
            if len(parts) >= 2:
                date_raw = parts[0].split('T')[0]
                try:
                    data.append({'data': date_raw, 'val': int(float(parts[1].strip()))})
                except ValueError:
                    pass
    return data

def parse_meta_audience_fb(path):
    """Estrae audience FB: top_countries, age_gender, top_cities."""
    with open(path, 'rb') as f:
        raw = f.read()
    text = raw.decode('utf-16', errors='replace')
    lines = [l.replace('"', '').strip() for l in text.splitlines() if l.strip() and 'sep' not in l.lower()]
    result = {'top_countries': [], 'age_gender': [], 'top_cities': []}
    section = None
    for line in lines:
        if 'Top countries' in line:   section = 'countries_header'; continue
        if 'Age & gender' in line:    section = 'age'; continue
        if 'Top cities' in line:      section = 'cities_header'; continue
        parts = [p.strip() for p in line.split(',')]
        if section == 'countries_header' and len(parts) >= 2:
            try: result['top_countries'] = [{'paese': p, 'pct': float(v)} for p, v in zip(parts, lines[lines.index(line)+1].split(',')) if p]
            except: pass
            section = None
        elif section == 'age' and len(parts) == 3 and parts[0] and parts[0][0].isdigit():
            try: result['age_gender'].append({'fascia': parts[0], 'uomini': float(parts[1]), 'donne': float(parts[2])})
            except: pass
        elif section == 'cities_header' and len(parts) >= 2:
            try: result['top_cities'] = [{'citta': p, 'pct': float(v)} for p, v in zip(parts, lines[lines.index(line)+1].split(',')) if p]
            except: pass
            section = None
    return result

def parse_meta_audience_ig(path):
    """Estrae audience IG: followers_totali, top_countries, age_gender, top_cities."""
    with open(path, 'rb') as f:
        raw = f.read()
    text = raw.decode('utf-16', errors='replace')
    lines = [l.replace('"', '').strip() for l in text.splitlines() if l.strip() and 'sep' not in l.lower()]
    result = {'followers_totali': 0, 'top_countries': [], 'age_gender': [], 'top_cities': []}
    section = None
    for i, line in enumerate(lines):
        if 'Instagram followers' == line: section = 'followers'; continue
        if 'by gender and age' in line:   section = 'age'; continue
        if 'by top cities' in line:       section = 'cities'; continue
        if 'by top countries' in line:    section = 'countries'; continue
        parts = [p.strip().rstrip('%') for p in line.split(',')]
        if section == 'followers':
            try: result['followers_totali'] = int(parts[0]); section = None
            except: pass
        elif section == 'age' and len(parts) == 3 and parts[0][0].isdigit():
            try: result['age_gender'].append({'fascia': parts[0], 'donne': float(parts[1]), 'uomini': float(parts[2])})
            except: pass
        elif section == 'cities' and len(parts) == 2:
            try: result['top_cities'].append({'citta': parts[0], 'pct': float(parts[1])})
            except: pass
        elif section == 'countries' and len(parts) == 2:
            try: result['top_countries'].append({'paese': parts[0], 'pct': float(parts[1])})
            except: pass
    return result

def build_optimedia_data(fb_archive_dir, ig_archive_dir):
    """Legge tutti i CSV archiviati e costruisce OPTIMEDIA_DATA."""
    def load_series(folder, pattern):
        files = sorted(glob.glob(os.path.join(folder, f'*_{pattern}.csv')))
        by_date = {}
        for f in files:
            for row in parse_meta_csv_series(f):
                d = row['data']
                by_date[d] = by_date.get(d, 0) + row['val']
        return [{'data': k, 'val': v} for k, v in sorted(by_date.items())]

    def merge_daily(series_dict):
        by_date = {}
        for key, rows in series_dict.items():
            for row in rows:
                d = row['data']
                if d not in by_date:
                    by_date[d] = {'data': d}
                by_date[d][key] = row['val']
        return sorted(by_date.values(), key=lambda x: x['data'])

    def totk(daily, key): return sum(d.get(key, 0) for d in daily)

    # Facebook
    fb_series = {
        'views':        load_series(fb_archive_dir, 'views'),
        'viewers':      load_series(fb_archive_dir, 'viewers'),
        'interactions': load_series(fb_archive_dir, 'interactions'),
        'follows':      load_series(fb_archive_dir, 'follows'),
        'link_clicks':  load_series(fb_archive_dir, 'link_clicks'),
        'visits':       load_series(fb_archive_dir, 'visits'),
    }
    fb_daily = merge_daily(fb_series)
    fb_t = {k: totk(fb_daily, k) for k in ['views','viewers','interactions','follows','link_clicks','visits']}
    fb_er = round(fb_t['interactions'] / fb_t['views'] * 100, 2) if fb_t['views'] else 0

    # Instagram
    ig_series = {
        'views':        load_series(ig_archive_dir, 'views'),
        'reach':        load_series(ig_archive_dir, 'reach'),
        'interactions': load_series(ig_archive_dir, 'interactions'),
        'follows':      load_series(ig_archive_dir, 'follows'),
        'link_clicks':  load_series(ig_archive_dir, 'link_clicks'),
        'visits':       load_series(ig_archive_dir, 'visits'),
    }
    ig_daily = merge_daily(ig_series)
    ig_t = {k: totk(ig_daily, k) for k in ['views','reach','interactions','follows','link_clicks','visits']}
    ig_er = round(ig_t['interactions'] / ig_t['reach'] * 100, 2) if ig_t['reach'] else 0

    # Audience (ultimo file disponibile)
    fb_aud_files = sorted(glob.glob(os.path.join(fb_archive_dir, '*_audience.csv')))
    ig_aud_files = sorted(glob.glob(os.path.join(ig_archive_dir, '*_audience.csv')))
    fb_audience = parse_meta_audience_fb(fb_aud_files[-1]) if fb_aud_files else {}
    ig_audience = parse_meta_audience_ig(ig_aud_files[-1]) if ig_aud_files else {'followers_totali': 0}

    all_dates = [r['data'] for r in fb_daily + ig_daily]
    data_min = min(all_dates) if all_dates else ''
    data_max = max(all_dates) if all_dates else ''

    return {
        'meta': {'cliente': 'Optimedia', 'data_min': data_min, 'data_max': data_max,
                 'aggiornato': datetime.now().strftime('%Y-%m-%d')},
        'facebook': {
            'kpi': {**fb_t, 'er': fb_er},
            'metriche': fb_daily,
            'audience': fb_audience,
        },
        'instagram': {
            'kpi': {'followers': ig_audience.get('followers_totali', 0),
                    'nuovi_followers': ig_t['follows'], **ig_t, 'er': ig_er},
            'metriche': ig_daily,
            'audience': ig_audience,
        },
    }

def update_optimedia_dashboard(data):
    """Sostituisce OPTIMEDIA_DATA in optimedia.html."""
    with open(DASHBOARD_OPT, 'r', encoding='utf-8') as f:
        html = f.read()
    marker = 'const OPTIMEDIA_DATA = '
    end    = ';\n'
    idx = html.find(marker)
    if idx < 0:
        print('  ERRORE: OPTIMEDIA_DATA non trovato in optimedia.html')
        return False
    end_idx = html.find(end, idx + len(marker))
    if end_idx < 0:
        print('  ERRORE: fine OPTIMEDIA_DATA non trovata')
        return False
    new_block = marker + json.dumps(data, ensure_ascii=False, separators=(',', ':')) + end
    html = html[:idx] + new_block + html[end_idx + len(end):]
    with open(DASHBOARD_OPT, 'w', encoding='utf-8') as f:
        f.write(html)
    return True

def process_inbox_optimedia():
    """Archivia i CSV Meta da inbox/optimedia/ e restituisce True se ci sono file nuovi."""
    today = datetime.now().strftime('%Y-%m-%d')
    found = False

    fb_map = {
        'Views.csv': 'views', 'Viewers.csv': 'viewers', 'Interactions.csv': 'interactions',
        'Follows.csv': 'follows', 'Link clicks.csv': 'link_clicks', 'Visits.csv': 'visits',
        'Audience FB.csv': 'audience',
    }
    ig_map = {
        'Views.csv': 'views', 'Reach.csv': 'reach', 'Interactions.csv': 'interactions',
        'Follows.csv': 'follows', 'Link clicks.csv': 'link_clicks', 'Visits.csv': 'visits',
        'Audience IG.csv': 'audience',
    }

    for src_dir, dst_dir, file_map, label in [
        (INBOX_OPT_FB, ARCHIVE_OPT_FB, fb_map, 'Facebook'),
        (INBOX_OPT_IG, ARCHIVE_OPT_IG, ig_map, 'Instagram'),
    ]:
        os.makedirs(dst_dir, exist_ok=True)
        for fname, key in file_map.items():
            src = os.path.join(src_dir, fname)
            if os.path.exists(src):
                dst = os.path.join(dst_dir, f'{today}_{key}.csv')
                shutil.copy2(src, dst)
                os.remove(src)
                print(f'  {label}/{fname} → {os.path.basename(dst)}')
                found = True

    return found


# ─── KEYE API ────────────────────────────────────────────────────────────────

def _keye_get(url):
    """Scarica un endpoint Keye con cookie auth, restituisce dict."""
    req = urllib.request.Request(url, headers={'Cookie': KEYE_COOKIE})
    with urllib.request.urlopen(req, timeout=30) as resp:
        return json.loads(resp.read())

def fetch_keye():
    """Scarica brand-summary + full export, li unisce e salva in archive/keye/."""
    os.makedirs(ARCHIVE_KEYE, exist_ok=True)
    today = datetime.now().strftime('%Y-%m-%d')
    dst   = os.path.join(ARCHIVE_KEYE, f'{today}_keye.json')

    print('  Fetch Keye brand-summary...')
    summary = _keye_get(KEYE_SUMMARY_URL)

    print('  Fetch Keye full export...')
    req_export = urllib.request.Request(
        KEYE_EXPORT_URL,
        headers={'X-Export-Key': KEYE_API_KEY}
    )
    with urllib.request.urlopen(req_export, timeout=30) as resp:
        export = json.loads(resp.read())

    # Merge: brand-summary primario per KPI/topic/competitor
    #        export per history/AI/SoS/channels/mentions/actors
    merged = dict(summary)
    for key in ('history', 'aiVisibility', 'shareOfSearch',
                'geoPresence', 'positionTargets', 'bestMentions',
                'thoughtLeaders', 'engagementActors', 'seoAudit'):
        if key in export:
            merged[key] = export[key]

    data = json.dumps(merged, ensure_ascii=False).encode('utf-8')
    with open(dst, 'wb') as f:
        f.write(data)

    print(f'  Salvato: {os.path.basename(dst)} ({len(data)//1024} KB)')
    return dst


def _score_to_grade(score):
    if score >= 90: return 'A+'
    if score >= 80: return 'A'
    if score >= 70: return 'B'
    if score >= 60: return 'C'
    if score >= 50: return 'D'
    return 'F'

def keye_from_export(d):
    """Costruisce KEYE_DATA dal nuovo endpoint /api/brand-summary."""
    kpi  = d.get('kpi', {})
    ci   = d.get('competitorIntelligence', {}) or {}

    # Overall score: media pesata SoV + sentiment
    sov       = kpi.get('shareOfVoice', 0)
    sentiment = kpi.get('sentimentScore', 0)
    sov_score = min(100, round(sov * 1.5))
    overall_score = round(sov_score * 0.4 + sentiment * 0.6)
    sov_delta = kpi.get('sovDelta', 0)
    verdict   = 'growing' if sov_delta > 2 else ('declining' if sov_delta < -2 else 'stable')

    # Competitors SOV
    comp_sov = [
        {'name': c.get('name'), 'share': c.get('share_pct', 0), 'color': ''}
        for c in (d.get('competitors') or [])
    ]

    # Topic SoV → formato atteso dal dashboard (rsShare, rsMentions, totalMentions, position)
    topic_sov = [
        {'topic':         t.get('topic', ''),
         'rsShare':       t.get('brand_share_pct', 0),
         'rsMentions':    0,
         'totalMentions': 0,
         'position':      t.get('brand_position', ''),
         'trend':         t.get('trend', 'stable'),
         'velocityRatio': t.get('velocity_ratio', 1.0),
         'leaderName':    t.get('leader_name', ''),
         'leaderShare':   t.get('leader_share_pct', 0)}
        for t in (d.get('topicSov') or [])
    ]

    # Signals: priority_actions dal competitor intelligence
    signals = [str(s) for s in (ci.get('priority_actions') or [])]

    # Competitor profiles → topActors placeholder (non disponibile in brand-summary)
    profiles = ci.get('profiles') or []
    top_actors = [
        {'name':          p.get('name', ''),
         'interactions':  0,
         'avgEr':         0,
         'sentiment':     '',
         'advocacyScore': round(100 - p.get('threat_score', 50))}
        for p in profiles[:5]
    ]

    # Dimensioni derivate dai dati disponibili
    pressure = ci.get('overall_competitive_pressure') or 50
    dims = [
        {'name': 'visibility',           'score': sov_score,               'grade': _score_to_grade(sov_score),    'delta': round(sov_delta), 'interpretation': f'SoV {sov}% (+{sov_delta}pp delta)'},
        {'name': 'sentimentHealth',      'score': sentiment,               'grade': _score_to_grade(sentiment),   'delta': None,             'interpretation': f'Sentiment {sentiment}/100'},
        {'name': 'competitivePosition',  'score': max(0, 100 - pressure),  'grade': _score_to_grade(max(0,100-pressure)), 'delta': None,     'interpretation': (ci.get('market_positioning_summary') or '')[:100]},
        {'name': 'engagementQuality',    'score': min(100, round(kpi.get('engagementRate', 0) * 10)), 'grade': _score_to_grade(min(100,round(kpi.get('engagementRate',0)*10))), 'delta': None, 'interpretation': f'ER {kpi.get("engagementRate",0)}%'},
        {'name': 'aiAuthority',          'score': 0, 'grade': 'F', 'delta': None, 'interpretation': 'Dato non disponibile'},
        {'name': 'seoAuthority',         'score': 0, 'grade': 'F', 'delta': None, 'interpretation': 'Dato non disponibile'},
    ]

    # ── Sezioni dal full export (mergiato da fetch_keye) ──────────────────
    hist_raw = (d.get('history') or {}).get('snapshots') or []
    history  = [
        {'date':           s.get('capturedAt', '')[:10],
         'totalMentions':  s.get('kpi', {}).get('totalMentions', 0),
         'sentimentScore': s.get('kpi', {}).get('sentimentScore', 0),
         'shareOfVoice':   s.get('kpi', {}).get('shareOfVoice', 0),
         'engagementRate': s.get('kpi', {}).get('engagementRate', 0)}
        for s in hist_raw
    ]

    ai = d.get('aiVisibility') or {}
    platforms = [
        {'platform':     p.get('platform'),    'model':       p.get('model'),
         'mentionRate':  p.get('mentionRate'),  'avgPosition': p.get('avgPosition'),
         'queriesRun':   p.get('queriesRun'),   'mentionCount':p.get('mentionCount'),
         'capturedAt':   p.get('capturedAt')}
        for p in (ai.get('platforms') or [])
    ]
    # aiAuthority/seoAuthority dal kpiPackage se disponibile
    kp   = d.get('kpiPackage') or {}
    dims_export = (kp.get('dimensions') or {})
    def _dim_ex(name):
        ex = dims_export.get(name, {})
        return ex.get('score', 0), ex.get('grade', 'F'), ex.get('delta'), ex.get('interpretation', '')
    ai_sc, ai_gr, ai_dl, ai_int = _dim_ex('aiAuthority')
    seo_sc, seo_gr, seo_dl, seo_int = _dim_ex('seoAuthority')
    dims[4] = {'name': 'aiAuthority',  'score': ai_sc,  'grade': ai_gr,  'delta': ai_dl,  'interpretation': ai_int  or 'Dato non disponibile'}
    dims[5] = {'name': 'seoAuthority', 'score': seo_sc, 'grade': seo_gr, 'delta': seo_dl, 'interpretation': seo_int or 'Dato non disponibile'}

    sos_raw = d.get('shareOfSearch') or {}
    share_of_search = {
        'window': sos_raw.get('window', ''), 'geo': sos_raw.get('geo', 'IT'),
        'computedAt': sos_raw.get('computedAt', ''),
        'results': [{'brand': r.get('brand'), 'sos_pct': r.get('sos_pct', 0),
                     'avg_interest': r.get('avg_interest', 0), 'trend': r.get('trend', 'stable')}
                    for r in (sos_raw.get('results') or [])]
    }

    geo = d.get('geoPresence') or {}
    channels = [
        {'channel': c.get('channel'), 'brandScore': c.get('brandScore', 0),
         'competitorGap': c.get('competitorGap', 0), 'trend': c.get('trend', 'stable'),
         'rationale': (c.get('rationale') or '')[:200]}
        for c in (geo.get('channelStrength') or [])
    ]

    pos_targets = [
        {'id': t.get('id'), 'label': t.get('label'), 'statement': (t.get('statement') or '')[:120],
         'status': t.get('status', ''), 'trend': t.get('trend', 'stable'),
         'supportingKeywords': (t.get('supportingKeywords') or [])[:6]}
        for t in (d.get('positionTargets') or [])
    ]

    best_mentions = [
        {'id': m.get('id'), 'source': m.get('source'),
         'content': (m.get('content') or '')[:300].replace('\n', ' '),
         'url': m.get('url'), 'publishedAt': m.get('publishedAt'),
         'sentiment': m.get('sentiment'), 'topics': m.get('topics', []),
         'engagementRate': m.get('engagementRate'), 'author': m.get('author')}
        for m in (d.get('bestMentions') or [])[:20]
    ]

    actors_raw = sorted(d.get('engagementActors') or [], key=lambda a: a.get('advocacy_score', 0), reverse=True)
    top_actors = [
        {'name':         a.get('name', ''),
         'interactions': a.get('interaction_count', 0),
         'avgEr':        round(a.get('avg_er_contribution', 0), 2),
         'sentiment':    a.get('sentiment', ''),
         'advocacyScore': a.get('advocacy_score', 0)}
        for a in actors_raw[:10]
    ]

    return {
        'meta': {
            'exportedAt':  d.get('generatedAt', ''),
            'brand':       d.get('brand', ''),
            'industry':    d.get('industry', ''),
            'period_from': (d.get('snapshotAt') or '')[:10],
            'period_to':   (d.get('generatedAt') or '')[:10],
        },
        'overall': {
            'score':        overall_score,
            'grade':        _score_to_grade(overall_score),
            'verdict':      verdict,
            'verdictReason': (ci.get('market_positioning_summary') or '')[:200],
            'delta':        round(sov_delta),
            'daysBetween':  7,
        },
        'dimensions':       dims,
        'snapshot': {
            'totalMentions':      kpi.get('totalMentions', 0),
            'mentionsDelta':      0,
            'sentimentScore':     sentiment,
            'sentimentBreakdown': kpi.get('sentiment', {'positive': 0, 'neutral': 0, 'negative': 0}),
            'shareOfVoice':       sov,
            'sovDelta':           sov_delta,
            'engagementRate':     kpi.get('engagementRate', 0),
            'topPerforming':      (d.get('topicClusters') or [{'name': '—'}])[0].get('name', '—'),
        },
        'signals':          signals,
        'topicClusters':    d.get('topicClusters', []),
        'competitorsSov':   comp_sov,
        'competitorDeltas': [],
        'sentimentByTopic': d.get('sentimentByTopic', []),
        'topicSov':         topic_sov,
        'history':          history,
        'aiVisibility': {
            'aiSignal':        ai.get('aiSignal', 0),
            'aiSov':           ai.get('aiSov', 0),
            'modelDivergence': ai.get('modelDivergence', 0),
            'platforms':       platforms,
        },
        'shareOfSearch':    share_of_search,
        'channels':         channels,
        'geoDistribution':  [],
        'positionTargets':  pos_targets,
        'industryContext':  '',
        'seo':              {'overall': 0, 'grade': '', 'categories': []},
        'thoughtLeaders':   [],
        'topActors':        top_actors,
        'bestMentions':     best_mentions,
    }


def update_dashboard(rs_data):
    """Sostituisce il blocco const RS_DATA = {...}; nel dashboard.html."""
    with open(DASHBOARD, 'r', encoding='utf-8') as f:
        html = f.read()

    start_marker = 'const RS_DATA = '
    end_marker   = '\n\nconst KEYE_DATA'

    start_idx = html.find(start_marker)
    end_idx   = html.find(end_marker, start_idx)

    if start_idx == -1 or end_idx == -1:
        print('ERRORE: marcatori RS_DATA / KEYE_DATA non trovati in dashboard.html')
        return False

    new_json = json.dumps(rs_data, ensure_ascii=False)
    new_html = html[:start_idx] + start_marker + new_json + ';' + html[end_idx:]

    with open(DASHBOARD, 'w', encoding='utf-8') as f:
        f.write(new_html)

    return True


def update_keye_data():
    """Legge il file keye più recente e aggiorna il blocco KEYE_DATA nel dashboard."""
    # Supporta sia il nuovo formato (????-??-??_keye.json) che il vecchio (_keye-export.json)
    keye_files = sorted(
        glob.glob(os.path.join(ARCHIVE_KEYE, '????-??-??_keye.json')) +
        glob.glob(os.path.join(ARCHIVE_KEYE, '????-??-??_keye-export.json'))
    )
    if not keye_files:
        print('  Nessun keye export in archivio — KEYE_DATA non aggiornato.')
        return False

    latest = keye_files[-1]
    print(f'  Keye export: {os.path.basename(latest)}')

    with open(latest, 'r', encoding='utf-8') as f:
        raw = json.load(f)

    keye_data = keye_from_export(raw)
    keye_json = json.dumps(keye_data, ensure_ascii=False)

    with open(DASHBOARD, 'r', encoding='utf-8') as f:
        html = f.read()

    start_marker = '\n\nconst KEYE_DATA = '
    end_marker   = ';\n\n/* ============ FORMATTERS'


    start_idx = html.find(start_marker)
    end_idx   = html.find(end_marker, start_idx)

    if start_idx == -1 or end_idx == -1:
        print('  ERRORE: marcatori KEYE_DATA non trovati.')
        return False

    new_html = html[:start_idx] + '\n\nconst KEYE_DATA = ' + keye_json + html[end_idx:]

    with open(DASHBOARD, 'w', encoding='utf-8') as f:
        f.write(new_html)

    return True


# ─── CAMPAIGN CSV ────────────────────────────────────────────────────────────

def _mdy(s):
    """M/D/YYYY → YYYY-MM-DD"""
    s = (s or '').strip().strip('"')
    if not s:
        return ''
    for fmt in ('%m/%d/%Y', '%Y-%m-%d'):
        try:
            return datetime.strptime(s, fmt).strftime('%Y-%m-%d')
        except ValueError:
            pass
    return ''

def _f(row, *cols):
    for col in cols:
        v = row.get(col, '').strip().strip('"').replace(',', '')
        if v:
            try:
                return float(v)
            except ValueError:
                pass
    return 0.0

def parse_campaign_csv(path):
    """Legge il campaign performance report CSV di LinkedIn Campaign Manager.
    Supporta: UTF-16 LE/BE, UTF-8 BOM, UTF-8. Separatore: tab o virgola.
    Gestisce sia report aggregati (1 riga/campagna) che giornalieri (N righe/campagna).
    """
    with open(path, 'rb') as f:
        raw = f.read()

    text = None
    for enc in ('utf-16', 'utf-8-sig', 'utf-8', 'latin-1'):
        try:
            text = raw.decode(enc)
            break
        except (UnicodeDecodeError, LookupError):
            pass
    if text is None:
        raise ValueError('Encoding CSV campagne non riconosciuto')

    lines = text.splitlines()

    def _parse_report_date(s):
        s = s.strip().strip('"')
        for fmt in ('%B %d, %Y, %I:%M %p', '%m/%d/%Y', '%Y-%m-%d', '%d/%m/%Y'):
            try:
                return datetime.strptime(s, fmt).strftime('%Y-%m-%d')
            except ValueError:
                pass
        return s[:10] if len(s) >= 10 else ''

    # Cerca report_start / report_end nelle prime 10 righe di metadata
    report_start, report_end, generated = '', '', ''
    for ml in lines[:10]:
        ml = ml.strip().strip('"')
        if ml.startswith('Report Start:'):
            report_start = _parse_report_date(ml.replace('Report Start:', '').strip())
        elif ml.startswith('Report End:'):
            report_end = _parse_report_date(ml.replace('Report End:', '').strip())
        elif ml.startswith('Date Generated:'):
            generated = _parse_report_date(ml.replace('Date Generated:', '').strip())

    # Rileva separatore (tab o virgola) e header row (prima riga con "Campaign Name")
    sep = '\t'
    header_idx = None
    for i, line in enumerate(lines):
        if 'Campaign Name' in line:
            sep = '\t' if '\t' in line else ','
            header_idx = i
            break
    if header_idx is None:
        raise ValueError('Header "Campaign Name" non trovato nel CSV campagne')

    headers = [h.strip().strip('"') for h in lines[header_idx].split(sep)]

    # ── Rilevamento formato: flat (91c/riga) vs multi-linea (23c prima riga + continuazione)
    # Nel formato multi-linea il testo creativo spanna più righe e le metriche
    # sono nell'ultima riga di ogni gruppo (offset +22 rispetto all'header).
    first_data_lines = [l for l in lines[header_idx+1:header_idx+5] if l.strip()]
    is_multiline = first_data_lines and len(first_data_lines[0].split(sep)) < 40

    if is_multiline:
        # Formato multi-linea: raggruppa per creative, estrai metriche dall'ultima riga
        creatives_raw = []
        current = None
        for l in lines[header_idx+1:]:
            cells = l.split(sep)
            cid = cells[3].strip().strip('"') if len(cells) > 3 else ''
            if len(cells) >= 20 and cid.isdigit():
                if current: creatives_raw.append(current)
                current = {'first': cells, 'last': None}
            elif current and len(cells) >= 50:
                current['last'] = cells
        if current: creatives_raw.append(current)

        def fvl(cells, i):
            try: return float(cells[i].strip().strip('"').replace(',','')) if i < len(cells) else 0.0
            except: return 0.0

        camps = defaultdict(lambda: {'id':'','name':'','objective':'','status':'','budget':0.0,'start':'','end':'','currency':'GBP'})
        daily_rows = []
        for c in creatives_raw:
            f, l = c['first'], c['last']
            if not l: continue
            cid  = f[3].strip().strip('"')
            name = f[4].strip().strip('"')
            c_obj= camps[name]
            c_obj['id']     = cid
            c_obj['name']   = name
            c_obj['status'] = f[5].strip().strip('"') if len(f)>5 else ''
            c_obj['currency'] = f[2].strip().strip('"') if len(f)>2 else 'GBP'
            try: c_obj['budget'] = float(f[8].strip().strip('"').replace(',','')) if len(f)>8 and f[8].strip() else c_obj['budget']
            except: pass
            # Metriche dalla riga finale (offset -22): spent=l[6], imp=l[7], clicks=l[8]...
            date_iso = _mdy(f[0]) or report_end or datetime.now().strftime('%Y-%m-%d')
            daily_rows.append({
                'date': date_iso, 'camp_id': cid, 'camp_name': name,
                'spent':       fvl(l, 6),
                'impressions': int(fvl(l, 7)),
                'clicks':      int(fvl(l, 8)),
                'reactions':   int(fvl(l, 12)),
                'comments':    int(fvl(l, 13)),
                'shares':      int(fvl(l, 14)),
                'follows':     int(fvl(l, 15)),
                'conversions': int(fvl(l, 27)),
                'leads':       int(fvl(l, 37)),
                'reach':       int(fvl(l, 42)),
                'video_views': 0,
            })

        camp_meta = {}
        for c in camps.values():
            camp_meta[c['id']] = {'id':c['id'],'name':c['name'],'objective':c.get('objective',''),
                                  'status':c['status'],'budget':round(c['budget'],2),
                                  'start':c.get('start',''),'end':c.get('end',''),'currency':c['currency']}
        return {
            'meta': {'report_start': report_start, 'report_end': report_end,
                     'generated': generated,
                     'currency': next(iter(camp_meta.values()))['currency'] if camp_meta else 'GBP'},
            'camp_meta': list(camp_meta.values()),
            'daily': daily_rows,
        }

    # ── Formato flat: logica originale ──────────────────────────────────────
    data_lines = [l for l in lines[header_idx + 1:] if l.strip()]

    # Nomi colonna alternativi usati da versioni diverse di Campaign Manager
    DATE_COLS    = ['Start Date (in UTC)', 'Date', 'Day', 'Start Date']
    SPENT_COLS   = ['Total Spent', 'Amount Spent (GBP)', 'Amount Spent', 'Spend']
    REACH_COLS   = ['Reach', 'Unique Impressions']
    FOLLOWS_COLS = ['Follows', 'New Followers', 'Follow']

    def _get(row, *keys):
        for k in keys:
            if k in row and row[k].strip():
                return row[k]
        return ''

    camps = defaultdict(lambda: {
        'id': '', 'name': '', 'objective': '', 'status': '',
        'budget': 0.0, 'start': '', 'end': '', 'currency': 'GBP',
    })
    daily_rows = []

    for line in data_lines:
        cells = line.split(sep)
        row = {headers[i]: cells[i].strip().strip('"') if i < len(cells) else '' for i in range(len(headers))}

        name = row.get('Campaign Name', '').strip()
        if not name:
            continue

        # Salta righe di continuazione del copy degli annunci:
        # Campaign ID deve essere numerico; se vuoto o URL → riga spuria
        camp_id_raw = row.get('Campaign ID', '').strip().strip('"')
        if not camp_id_raw.isdigit():
            continue

        # Data: prova colonne note, fallback a report_end (report aggregato)
        date_iso = ''
        for dc in DATE_COLS:
            date_iso = _mdy(row.get(dc, ''))
            if date_iso:
                break
        if not date_iso:
            date_iso = report_end or datetime.now().strftime('%Y-%m-%d')

        c = camps[name]
        c['id']        = row.get('Campaign ID', c['id'])
        c['name']      = name
        c['objective'] = row.get('Campaign Objective Type', c['objective']) or c['objective']
        c['status']    = row.get('Campaign Status', c['status'])
        c['currency']  = row.get('Currency', 'GBP')
        bud = row.get('Campaign Total Budget', '').replace(',', '').strip('"')
        if bud:
            try: c['budget'] = float(bud)
            except: pass
        s = _mdy(row.get('Campaign Start Date', ''))
        if s and (not c['start'] or s < c['start']): c['start'] = s
        e = _mdy(row.get('Campaign End Date', ''))
        if e and (not c['end'] or e > c['end']): c['end'] = e

        daily_rows.append({
            'date':        date_iso,
            'camp_id':     c['id'],
            'camp_name':   name,
            'spent':       round(_f(row, *SPENT_COLS), 4),
            'impressions': int(_f(row, 'Impressions')),
            'clicks':      int(_f(row, 'Clicks')),
            'reactions':   int(_f(row, 'Reactions')),
            'comments':    int(_f(row, 'Comments')),
            'shares':      int(_f(row, 'Shares')),
            'follows':     int(_f(row, *FOLLOWS_COLS)),
            'conversions': int(_f(row, 'Conversions')),
            'leads':       int(_f(row, 'Leads')),
            'reach':       int(_f(row, *REACH_COLS)),
            'video_views': int(_f(row, 'Video Views')),
        })

    # camp_meta: una entry per campagna (dati statici: budget, obiettivo, stato, date)
    camp_meta = {}
    for c in camps.values():
        camp_meta[c['id']] = {
            'id':        c['id'],
            'name':      c['name'],
            'objective': c['objective'],
            'status':    c['status'],
            'budget':    round(c['budget'], 2),
            'start':     c['start'],
            'end':       c['end'],
            'currency':  c['currency'],
        }

    return {
        'meta': {
            'report_start': report_start,
            'report_end':   report_end,
            'generated':    generated,
            'currency':     next(iter(camp_meta.values()))['currency'] if camp_meta else 'GBP',
        },
        'camp_meta': list(camp_meta.values()),
        'daily':     daily_rows,
    }


def parse_creative_csv(path):
    """Legge il Creative/Ad Performance Report di LinkedIn Campaign Manager.
    Il file ha righe multi-line: prima riga (23 celle) + testo ad + ultima riga (metriche).
    Restituisce lista di creative con metriche paid."""
    with open(path, 'rb') as f: raw = f.read()
    text = raw.decode('utf-16')
    lines = text.splitlines()

    header_idx = next((i for i, l in enumerate(lines) if 'Campaign Name' in l and 'Impressions' in l), None)
    if header_idx is None:
        raise ValueError('Header non trovato nel Creative Performance CSV')

    def clean(s): return s.strip().strip('"').strip()
    def fv(cells, i):
        try: return float(clean(cells[i]).replace(',','').replace('%','')) if i < len(cells) else 0.0
        except: return 0.0

    # Raggruppa linee per creative (prima riga = 23 celle con Campaign ID numerico)
    creatives, current = [], None
    for l in lines[header_idx+1:]:
        cells = l.split('\t')
        if len(cells) >= 23 and clean(cells[3]).isdigit():
            if current: creatives.append(current)
            current = {'first': cells, 'last': None}
        elif current and len(cells) >= 50:
            current['last'] = cells
    if current: creatives.append(current)

    rows = []
    for c in creatives:
        f, l = c['first'], c['last']
        if not l: continue
        # Offset: l'ultima riga inizia da col 22 dell'header (Ad Introduction Text)
        # → click_url=l[3] (header 25), spent=l[6] (28), imp=l[7] (29)...
        rows.append({
            'camp_id':    clean(f[3]),
            'camp_name':  clean(f[4]),
            'ad_name':    clean(f[19]) if len(f) > 19 else '',
            'click_url':  clean(l[3])  if len(l) > 3  else '',
            'spent':      fv(l, 6),
            'impressions':fv(l, 7),
            'clicks':     fv(l, 8),
            'reactions':  fv(l, 12),
            'comments':   fv(l, 13),
            'shares':     fv(l, 14),
            'er':         fv(l, 19),
            'reach':      fv(l, 42),
            'leads':      fv(l, 37),
        })
    return rows


def update_creative_data(creatives):
    """Sostituisce il blocco CREATIVE_DATA nel dashboard.html."""
    with open(DASHBOARD, 'r', encoding='utf-8') as f:
        html = f.read()
    marker = 'const CREATIVE_DATA = '
    end    = ';\n\nconst CAMP_DATA'
    idx = html.find(marker)
    if idx < 0:
        return False
    end_idx = html.find(end, idx)
    if end_idx < 0: return False
    html = html[:idx] + marker + json.dumps(creatives, ensure_ascii=False) + html[end_idx:]
    with open(DASHBOARD, 'w', encoding='utf-8') as f:
        f.write(html)
    return True


def update_camp_data(camp_data):
    """Aggiunge / sostituisce il blocco CAMP_DATA nel dashboard.html."""
    with open(DASHBOARD, 'r', encoding='utf-8') as f:
        html = f.read()

    new_json     = json.dumps(camp_data, ensure_ascii=False)
    start_marker = '\n\nconst CAMP_DATA = '
    end_marker   = ';\n\nconst KEYE_DATA'

    if start_marker in html:
        # Aggiorna il blocco esistente
        s = html.find(start_marker)
        e = html.find(end_marker, s)
        if e == -1:
            print('  ERRORE: fine CAMP_DATA non trovata.')
            return False
        new_html = html[:s] + start_marker + new_json + html[e:]
    else:
        # Prima volta: inserisce prima di KEYE_DATA
        keye_marker = '\n\nconst KEYE_DATA'
        idx = html.find(keye_marker)
        if idx == -1:
            print('  ERRORE: KEYE_DATA non trovato nel dashboard.')
            return False
        # Assicura che CREATIVE_DATA sia inserito prima di CAMP_DATA
        creative_marker = '\n\nconst CREATIVE_DATA = '
        if creative_marker not in html[:idx]:
            html = html[:idx] + creative_marker + '[];\n' + html[idx:]
            idx = html.find(keye_marker)
        new_html = html[:idx] + start_marker + new_json + ';' + html[idx:]

    with open(DASHBOARD, 'w', encoding='utf-8') as f:
        f.write(new_html)
    return True


# ─── INBOX ───────────────────────────────────────────────────────────────────

def _sniff_xls_kind(path):
    """Rileva se un XLS è 'content' o 'followers' leggendo i nomi dei fogli."""
    try:
        wb = xlrd.open_workbook(path)
        names = [s.lower() for s in wb.sheet_names()]
        if 'metriche' in names or 'tutti i post' in names:
            return 'content'
        if 'nuovi follower' in names:
            return 'followers'
    except Exception:
        pass
    return None

def _sniff_csv_kind(path):
    """Verifica se un CSV è un campaign performance report."""
    try:
        with open(path, 'rb') as f:
            header = f.read(200)
        for enc in ('utf-16', 'utf-8-sig', 'utf-8'):
            try:
                text = header.decode(enc)
                if 'campaign' in text.lower() or 'performance report' in text.lower():
                    return 'campaign'
            except Exception:
                pass
    except Exception:
        pass
    return None

def process_inbox():
    """
    Scansiona inbox/, identifica e archivia i file LinkedIn.
    Restituisce dict con i file trovati e dove sono stati archiviati.
    """
    os.makedirs(INBOX, exist_ok=True)
    found = {'content': None, 'followers': None, 'campaign': None, 'unknown': []}

    entries = [
        os.path.join(INBOX, f) for f in os.listdir(INBOX)
        if not f.startswith('.') and os.path.isfile(os.path.join(INBOX, f))
    ]

    if not entries:
        return found

    print(f'\nInbox: {len(entries)} file trovati')

    for path in entries:
        fname = os.path.basename(path)
        ext   = os.path.splitext(fname)[1].lower()

        if ext in ('.xls', '.xlsx'):
            kind = _sniff_xls_kind(path)
            if kind in ('content', 'followers'):
                dst = archive_new_file(path, kind)
                found[kind] = dst
                os.remove(path)
                print(f'  {fname} → {kind} → {os.path.basename(dst)}')
            else:
                print(f'  {fname} → XLS non riconosciuto (skip)')
                found['unknown'].append(fname)

        elif ext == '.csv':
            kind = _sniff_csv_kind(path)
            if kind == 'campaign':
                # Archivia il CSV campagne in archive/campagne/
                os.makedirs(ARCHIVE_CAMPAGNE, exist_ok=True)
                ts = datetime.now().strftime('%Y-%m-%d')
                dst = os.path.join(ARCHIVE_CAMPAGNE, f'{ts}_campagne.csv')
                if os.path.exists(dst):
                    dst = dst.replace('.csv', f'_{datetime.now().strftime("%H%M%S")}.csv')
                shutil.copy2(path, dst)
                os.remove(path)
                found['campaign'] = dst
                print(f'  {fname} → campaign CSV → {os.path.basename(dst)}')
            else:
                print(f'  {fname} → CSV non riconosciuto (skip)')
                found['unknown'].append(fname)
        else:
            print(f'  {fname} → formato non supportato (skip)')
            found['unknown'].append(fname)

    return found


# ─── MAIN ────────────────────────────────────────────────────────────────────

def main():
    for d in [ARCHIVE, ARCHIVE_KEYE, INBOX_RS, INBOX_OPT_FB, INBOX_OPT_IG,
              ARCHIVE_OPT_FB, ARCHIVE_OPT_IG]:
        os.makedirs(d, exist_ok=True)

    args         = sys.argv[1:]
    skip_keye    = '--no-keye'   in args
    skip_inbox   = '--no-inbox'  in args
    only_rs      = '--rs-only'   in args
    only_opt     = '--opt-only'  in args
    args = [a for a in args if a not in ('--no-keye','--no-inbox','--keye','--rs-only','--opt-only')]

    if args:
        print('Uso: python3 update.py [--no-keye] [--no-inbox] [--rs-only] [--opt-only]')
        sys.exit(1)

    run_rs  = not only_opt
    run_opt = not only_rs

    # ══════════════════════════════════════════════════════════════
    # CLIENTE 1 — RS ITALIA
    # ══════════════════════════════════════════════════════════════
    if run_rs:
        print('\n' + '═'*50)
        print('RS ITALIA')
        print('═'*50)

        # 1. Inbox
        if not skip_inbox:
            process_inbox()

        # 2. Keye
        if not skip_keye:
            print('\nScarico dati Keye...')
            try:
                fetch_keye()
            except Exception as e:
                print(f'  WARN Keye fetch fallito: {e}')
                print('  Continuo con l\'ultimo export in archivio...')

        # 3. LinkedIn XLS → RS_DATA
        pairs = find_week_pairs()
        if not pairs:
            print('\n⚠  FILE MANCANTE: LinkedIn Analytics XLS non trovato.')
            print('   → Scarica da LinkedIn Page Analytics (Contenuti + Follower)')
            print(f'   → Metti in: inbox/rs-italia/')
            if not run_opt:
                sys.exit(1)
        else:
            print(f'\nSettimane in archivio ({len(pairs)}):')
            for date, cf, ff in pairs:
                print(f'  {date}  {os.path.basename(cf)} + {os.path.basename(ff)}')

            rs_data = build_rs_data(pairs, n_weeks=4)
            m = rs_data['meta']
            print(f'  LinkedIn: {m["data_min"]} → {m["data_max"]} · {len(rs_data["metriche"])} giorni · fan base {m["fan_base_totale"]:,}')
            update_history(pairs)

            # 4. Dashboard RS
            print('\nAggiornamento dashboard.html...')
            if not update_dashboard(rs_data):
                print('✗ RS_DATA fallito.')
            else:
                print('  RS_DATA   ✓')

            update_keye_data()
            print('  KEYE_DATA ✓')

            camp_files = sorted(glob.glob(os.path.join(ARCHIVE_CAMPAGNE, '????-??-??_campagne.csv')))
            if not camp_files:
                camp_files = sorted(glob.glob(os.path.join(ARCHIVE_ROOT, '*campaign_performance_report*.csv')))
            if camp_files:
                try:
                    # Carica DIRETTAMENTE dal CSV più recente (unica fonte di verità)
                    camp_data = parse_campaign_csv(camp_files[-1])
                    camp_data = _apply_krein_filter(camp_data)

                    # Salva in dashboard.html
                    update_camp_data(camp_data)
                    print(f'  CAMP_DATA ✓  ({len(camp_data["camp_meta"])} campagne Krein · {camp_data["meta"]["report_start"]} → {camp_data["meta"]["report_end"]})')
                except Exception as e:
                    print(f'  CAMP_DATA WARN: {e}')
            else:
                print('  ⚠  FILE MANCANTE: Campaign Manager CSV')
                print('     → Scarica da LinkedIn Campaign Manager → Reports → Ad Performance')
                print(f'     → Metti in: inbox/rs-italia/')

            # Creative Performance (per-post paid data)
            creative_files = sorted(glob.glob(os.path.join(ARCHIVE_CAMPAGNE, '????-??-??_creative_performance.csv')))
            if creative_files:
                try:
                    all_creatives = parse_creative_csv(creative_files[-1])
                    krein_creatives = [r for r in all_creatives if r['camp_id'] in {'987808183','1056604124'}]
                    if update_creative_data(krein_creatives):
                        print(f'  CREATIVE_DATA ✓  ({len(krein_creatives)} creative Krein)')
                    else:
                        print('  CREATIVE_DATA WARN: update fallito')
                except Exception as e:
                    print(f'  CREATIVE_DATA WARN: {e}')
            else:
                print('  ⚠  FILE MANCANTE: Creative Performance CSV')
                print('     → Scarica da LinkedIn Campaign Manager → Reports → Creative Performance')
                print(f'     → Rinomina/metti in: inbox/')

    # ══════════════════════════════════════════════════════════════
    # CLIENTE 2 — OPTIMEDIA
    # ══════════════════════════════════════════════════════════════
    if run_opt:
        print('\n' + '═'*50)
        print('OPTIMEDIA')
        print('═'*50)

        # 1. Inbox Optimedia
        if not skip_inbox:
            print('\nInbox Optimedia...')
            opt_found = process_inbox_optimedia()
            if not opt_found:
                print('  Nessun file nuovo in inbox/optimedia/')

        # 2. Ricostruisci OPTIMEDIA_DATA dall'archivio
        fb_files = glob.glob(os.path.join(ARCHIVE_OPT_FB, '*.csv'))
        ig_files = glob.glob(os.path.join(ARCHIVE_OPT_IG, '*.csv'))

        if not fb_files:
            print('  ⚠  FILE MANCANTE: Facebook CSV')
            print('     → Scarica da Meta Business Suite → Insight → Esporta')
            print(f'     → Metti in: inbox/optimedia/Facebook/  (Views, Viewers, Interactions, Follows, Visits, Link clicks, Audience FB)')
        if not ig_files:
            print('  ⚠  FILE MANCANTE: Instagram CSV')
            print('     → Scarica da Meta Business Suite → Insight → Esporta')
            print(f'     → Metti in: inbox/optimedia/Instagram/  (Views, Reach, Interactions, Follows, Visits, Link clicks, Audience IG)')

        if fb_files or ig_files:
            print('\nAggiornamento optimedia.html...')
            try:
                opt_data = build_optimedia_data(ARCHIVE_OPT_FB, ARCHIVE_OPT_IG)
                m = opt_data['meta']
                fb_k = opt_data['facebook']['kpi']
                ig_k = opt_data['instagram']['kpi']
                print(f'  Periodo:   {m["data_min"]} → {m["data_max"]}')
                print(f'  Facebook:  {fb_k["views"]:,} views · {fb_k["interactions"]} int · ER {fb_k["er"]}%')
                print(f'  Instagram: {ig_k["reach"]:,} reach · {ig_k["interactions"]} int · ER {ig_k["er"]}%')
                if update_optimedia_dashboard(opt_data):
                    print('  OPTIMEDIA_DATA ✓')
                else:
                    print('  OPTIMEDIA_DATA WARN: aggiornamento fallito')
            except Exception as e:
                print(f'  OPTIMEDIA WARN: {e}')
        else:
            print('\n  Nessun file Optimedia in archivio — optimedia.html invariato')

    # ══════════════════════════════════════════════════════════════
    # DEPLOY
    # ══════════════════════════════════════════════════════════════
    print(f'\n✓ Fatto.')
    data_max = datetime.now().strftime('%Y-%m-%d')
    if run_rs and 'rs_data' in dir():
        data_max = rs_data['meta'].get('data_max', data_max)

    commit_msg = f'Update dati: {data_max}'
    print(f'\nDeploy automatico...')
    try:
        files_to_add = ['dashboard.html', 'rs_history.json', 'optimedia.html']
        subprocess.run(['git', 'add'] + files_to_add, cwd=BASE, check=True)
        result = subprocess.run(['git', 'diff', '--cached', '--quiet'], cwd=BASE)
        if result.returncode != 0:
            subprocess.run(['git', 'commit', '-m', commit_msg], cwd=BASE, check=True)
            subprocess.run(['git', 'push', 'origin', 'main'], cwd=BASE, check=True)
            print(f'  Deploy ✓  →  {commit_msg}')
        else:
            print('  Nessuna modifica da deployare.')
    except subprocess.CalledProcessError as e:
        print(f'  Deploy WARN: {e}')


if __name__ == '__main__':
    main()
