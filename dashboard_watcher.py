"""
PCMA Executive Dashboard - Live Data Watcher
=============================================
HOW TO USE:

1. Install dependencies (one time):
     pip install watchdog openpyxl

2. Put ALL THREE files in the SAME folder:
     - dashboard_watcher.py
     - pcma_executive_dashboard.html
     - 2026_CEO___Enterprise_Goals_-_IT_POC.xlsx

3. Run the watcher:
     python dashboard_watcher.py

4. In a SEPARATE terminal, from the same folder run:
     python -m http.server 8080

5. Open browser to:
     http://localhost:8080/pcma_executive_dashboard.html

Whenever you save the Excel file the dashboard updates within 30 seconds.
"""

import json, os, sys, time, logging
from datetime import datetime

try:
    import openpyxl
    from watchdog.observers import Observer
    from watchdog.events import FileSystemEventHandler
except ImportError:
    print("ERROR: Missing packages. Run:  pip install watchdog openpyxl")
    sys.exit(1)

EXCEL_FILE  = "2026_CEO___Enterprise_Goals_-_IT_POC.xlsx"
OUTPUT_JSON = "dashboard_data.json"

logging.basicConfig(level=logging.INFO, format="%(asctime)s  %(message)s", datefmt="%H:%M:%S")
log = logging.getLogger("pcma")

def safe_num(v):
    try:    return float(v)
    except: return None

def parse_excel(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    sections = []
    current  = None
    for row in rows[1:]:
        weight = row[0]
        desc   = str(row[1]).strip().lstrip('\u2022 ') if row[1] else ''
        goal_r = row[2]
        ytd_r  = row[3]
        source = str(row[5]).strip() if row[5] else ''
        num_weight = safe_num(weight)
        if num_weight and 0 < num_weight <= 1 and desc:
            current = {'id':'unknown','title':desc.strip(),'weight':int(round(num_weight*100)),'metrics':[]}
            tl = desc.lower()
            if   'financial'    in tl: current['id'] = 'financial'
            elif 'membership'   in tl: current['id'] = 'membership'
            elif 'endowment'    in tl: current['id'] = 'endowment'
            elif 'education'    in tl: current['id'] = 'education'
            elif 'internal'     in tl or 'organization' in tl: current['id'] = 'internal'
            sections.append(current)
            continue
        if current is None or not desc: continue
        goal_n = safe_num(goal_r)
        ytd_n  = safe_num(ytd_r)
        ytd_s  = str(ytd_r).strip() if ytd_r is not None else 'TBD'
        goal_s = str(goal_r).strip() if goal_r is not None else 'TBD'
        for bad in ('None','nan',''): 
            if ytd_s==bad:  ytd_s='TBD'
            if goal_s==bad: goal_s='TBD'
        fmt = 'number'
        dl = desc.lower()
        if '%' in desc or 'growth' in dl or 'turnover' in dl or 'engagement' in dl: fmt='percent'
        if '$' in desc or 'profit' in dl or 'raise' in dl: fmt='currency'
        if goal_n is None and ytd_n is None: fmt='qualitative'
        inverse = 'turnover' in dl
        m = {'name':desc,'goal':goal_n,'ytd':ytd_n,'goal_label':goal_s,'ytd_label':ytd_s,'format':fmt,'source':source}
        if inverse: m['inverse']=True
        if ytd_s in ('Quite Phase','Quiet Phase'): m['status']='Quiet Phase'; m['ytd_label']='Quiet Phase'
        if ytd_s in ('?','Pending'): m['status']='Pending'; m['ytd_label']='Pending'
        if ytd_s in ('TBD','?%'): m['status']='TBD'; m['ytd_label']='TBD'
        if ytd_n and goal_n and ytd_n > goal_n and not inverse: m['exceeds']=True
        current['metrics'].append(m)

    checklist=[
        {'label':'Dashboards & Reporting','done':True},
        {'label':'Standardized Data Management','done':False},
        {'label':'Process Mapping & Marketing','done':True},
        {'label':'System Simplification','done':False},
        {'label':'Training & Enablement','done':True},
        {'label':'AI & Automation','done':False}
    ]
    for sec in sections:
        for m in sec['metrics']:
            if 'data' in m['name'].lower() and 'technology' in m['name'].lower():
                m['checklist']=checklist

    return {'generated_at':datetime.now().strftime('%b %d, %Y %H:%M'),'year_progress':16.7,'sections':sections}

def write_json(data):
    tmp=OUTPUT_JSON+'.tmp'
    with open(tmp,'w',encoding='utf-8') as f: json.dump(data,f,indent=2,ensure_ascii=False)
    os.replace(tmp,OUTPUT_JSON)
    log.info('dashboard_data.json updated (%d sections)',len(data['sections']))

def refresh():
    if not os.path.exists(EXCEL_FILE): log.warning('Excel not found: %s',EXCEL_FILE); return
    try: write_json(parse_excel(EXCEL_FILE))
    except Exception as e: log.error('Parse error: %s',e)

class ExcelHandler(FileSystemEventHandler):
    def __init__(self): self._last=0
    def on_modified(self,event):
        if EXCEL_FILE in event.src_path:
            now=time.time()
            if now-self._last>2:
                self._last=now; time.sleep(0.5); log.info('Change detected - refreshing...'); refresh()
    on_created=on_modified

if __name__=='__main__':
    print('='*52)
    print('  PCMA Executive Dashboard - Live Watcher')
    print('='*52)
    print(f'  Watching : {EXCEL_FILE}')
    print(f'  Output   : {OUTPUT_JSON}')
    print()
    print('  To view the dashboard:')
    print('  1. Keep this running')
    print('  2. New terminal, same folder:  python -m http.server 8080')
    print('  3. Browser: http://localhost:8080/pcma_executive_dashboard.html')
    print()
    print('  Dashboard polls for updates every 30 seconds.')
    print('  Ctrl+C to stop.')
    print('='*52)
    refresh()
    observer=Observer()
    observer.schedule(ExcelHandler(),path='.',recursive=False)
    observer.start()
    log.info('Watching for changes...')
    try:
        while True: time.sleep(1)
    except KeyboardInterrupt:
        observer.stop(); log.info('Stopped.')
    observer.join()
