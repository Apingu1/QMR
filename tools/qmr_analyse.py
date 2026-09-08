import json,re
from collections import Counter,defaultdict
from datetime import datetime,date
from pathlib import Path

START=date(2026,4,1); END=date(2026,6,30)
EX=Path('analysis/extracted'); OUT=Path('analysis/qmr_analysis.json'); OUTLINE=Path('analysis/reference_outline.json')

def dt(v):
    if not v or not isinstance(v,str): return None
    try:return datetime.strptime(v[:10],'%Y-%m-%d').date()
    except:return None

def norm_status(v): return str(v or '').strip().lower().replace(' ','').replace('-','')
def is_done(v): return norm_status(v) in {'complete','completed','closed','yes','y'}
def base_ref(v): return re.sub(r'/[A-Z0-9]+$','',str(v or '').strip(),flags=re.I)

def load(name): return json.loads((EX/name).read_text(encoding='utf-8'))
def rows_from(name,sheet=0,header_row=None):
    d=load(name); sh=d['sheets'][sheet]; rows=sh['rows']
    if header_row is None:
        # choose row with most nonblank text cells among first 5 populated rows
        cand=max(rows[:5],key=lambda r:sum(1 for x in r['values'] if str(x or '').strip()))
        header_row=cand['row']
    headers=next(r['values'] for r in rows if r['row']==header_row)
    out=[]
    for r in rows:
        if r['row']<=header_row: continue
        vals=r['values']+['']*max(0,len(headers)-len(r['values']))
        rec={str(h).strip():vals[i] for i,h in enumerate(headers) if str(h).strip()}
        rec['_row']=r['row']; out.append(rec)
    return headers,out

def get(rec,*names):
    for n in names:
        if n in rec:return rec[n]
    return ''

def classify_records(records, cfg):
    cleaned=[]
    for r in records:
        ref=str(get(r,*cfg['ref'])).strip()
        if not ref: continue
        od=dt(get(r,*cfg['open']))
        if not od or od>END: continue
        cd=dt(get(r,*cfg.get('close',()))) if cfg.get('close') else None
        target=dt(get(r,*cfg.get('target',()))) if cfg.get('target') else None
        status=get(r,*cfg.get('status',()))
        close_known=bool(cfg.get('close'))
        if close_known:
            active=(cd is None or cd>=START)
            open_end=(cd is None or cd>END)
            closed_period=bool(cd and START<=cd<=END)
        else:
            # no closure date in source: period-raised records are reportable; older records only if still not marked complete
            active=(od>=START) or (not is_done(status))
            open_end=not is_done(status)
            closed_period=(od>=START and is_done(status))
        if not active: continue
        rr=dict(r)
        rr.update({'_ref':ref,'_base_ref':base_ref(ref),'_open':od.isoformat(),'_close':cd.isoformat() if cd else '', '_target':target.isoformat() if target else '', '_status_end':'Open' if open_end else 'Closed', '_new_period':START<=od<=END, '_closed_period':closed_period, '_carried':od<START, '_overdue_end':bool(open_end and target and target<END)})
        cleaned.append(rr)
    return cleaned

def summarize(name, records, cfg, dedupe=False):
    rs=classify_records(records,cfg)
    if dedupe:
        # metrics by parent reference; keep most informative row per parent for details
        groups=defaultdict(list)
        for r in rs: groups[r['_base_ref']].append(r)
        metric_rs=[]
        for k,g in groups.items():
            # merge flags and pick first row
            x=dict(g[0])
            x['_new_period']=any(z['_new_period'] for z in g); x['_closed_period']=all(z['_closed_period'] for z in g) if g else False
            x['_carried']=all(z['_carried'] for z in g)
            x['_status_end']='Open' if any(z['_status_end']=='Open' for z in g) else 'Closed'
            x['_overdue_end']=any(z['_overdue_end'] for z in g)
            metric_rs.append(x)
    else: metric_rs=rs
    sev=[str(get(r,*cfg.get('severity',()))).strip() for r in metric_rs if str(get(r,*cfg.get('severity',()))).strip()]
    status=[str(get(r,*cfg.get('status',()))).strip() for r in metric_rs if str(get(r,*cfg.get('status',()))).strip()]
    months=Counter()
    for r in metric_rs:
        if r['_new_period']: months[r['_open'][:7]]+=1
    significant=[]
    for r in rs:
        s=' '.join(str(get(r,*cfg.get('severity',()))).lower().split())
        reclass=' '.join(str(get(r,*cfg.get('reclass',()))).lower().split())
        if any(x in s or x in reclass for x in ('major','critical')):
            significant.append({k:v for k,v in {
                'ref':r['_ref'],'date':r['_open'],'severity':get(r,*cfg.get('severity',())),'reclassification':get(r,*cfg.get('reclass',())),
                'summary':get(r,*cfg.get('summary',())),'root_cause':get(r,*cfg.get('root',())),'status_at_jun':r['_status_end'],'close_date':r['_close'],'target':r['_target']}.items() if v not in ('',None)})
    return {
        'name':name,'counts':{'active_in_period':len(metric_rs),'brought_forward':sum(r['_carried'] for r in metric_rs),'new_in_period':sum(r['_new_period'] for r in metric_rs),'closed_in_period':sum(r['_closed_period'] for r in metric_rs),'open_at_30_jun':sum(r['_status_end']=='Open' for r in metric_rs),'overdue_open_at_30_jun':sum(r['_overdue_end'] for r in metric_rs)},
        'new_by_month':dict(months),'classification':dict(Counter(sev)),'current_source_status':dict(Counter(status)), 'significant_items':significant,
        'records':[{k:v for k,v in {'ref':r['_ref'],'base_ref':r['_base_ref'],'date_opened':r['_open'],'date_closed':r['_close'],'target':r['_target'],'status_at_30_jun':r['_status_end'],'carried_forward':r['_carried'],'new_in_period':r['_new_period'],'closed_in_period':r['_closed_period'],'overdue_at_30_jun':r['_overdue_end'],'severity':get(r,*cfg.get('severity',())),'department':get(r,*cfg.get('department',())),'area':get(r,*cfg.get('area',())),'summary':get(r,*cfg.get('summary',())),'root_cause':get(r,*cfg.get('root',()))}.items() if v not in ('',None,False)} for r in rs]
    }

configs={
'CAPA':('2026 - QMS DATA CAPA Version 1.1.json',2,{'ref':('CAPA Reference',),'open':('Date Logged',),'close':('Date Closed',),'target':('Target date',),'status':('Status',),'severity':('Category (Minor / Major / Critical)',),'summary':('Assignable Cause',),'root':('Assignable Cause',)}),
'Deviations':('2026 - QMS DATA DEVIATIONS Version 2.0 (to be Validated).json',2,{'ref':('Deviation Reference',),'open':('Date Deviation Raised',),'close':('Investigation Closed',),'target':('Target Date',),'status':('Status',),'severity':('Classification of Deviation',),'reclass':('Reclassification of Deviation',),'department':('Department/Area Affected',),'area':('Deviation Area (Processes and activities)',),'summary':('Brief  Summary','Brief Summary'),'root':('Root Cause(s) Identified',)}),
'Change Controls':('2026 - QMS DATA CHANGE CONTROLS Version 1.0.json',2,{'ref':('Change Control Number',),'open':('Date Logged',),'close':('Completion Date',),'target':('Target Date:',),'status':('Current Status',),'severity':('Classification',),'area':('Change Type',),'summary':('Change Summary',)}),
'Document Change Requests':('2026 - QMS DATA DOCUMENT CHANGE REQUEST Version 1.0.json',2,{'ref':('Change Request Number',),'open':('Date Raised',),'status':('Current Status',),'severity':('Change Type:',),'summary':('Details of change',)}),
'Effectiveness Checks':('2026 - QMS DATA EFFECTIVENESS CHECKS Version 1.0.json',2,{'ref':('CC Reference ',),'open':('Date CC Raised',),'close':('Date Closed',),'target':('Target date',),'status':('Status',),'summary':('Effectiveness check action(s)',)}),
'FMEA':('2026 - QMS DATA FMEA (to be Validated).json',1,{'ref':('FMEA Number',),'open':('Date Opened',),'close':('Date Closed',),'target':('Deadline',),'status':('Status',),'severity':('Risk Category',),'summary':('Title',),'root':('Potential failure mode risk',)}),
'Incidents':('2026 - QMS DATA Incident Management Log (to be Validated).json',1,{'ref':('Incident Reference',),'open':('Date Logged',),'close':('Incident report Completion Date',),'target':('Incident report Due Date',),'status':('Status',),'severity':('Criticality of Incident (Critical/ Major)',),'summary':('Details of Incident',)}),
'OOS-OOT':('2026 - QMS DATA OOS-OOT Version 1.0.json',2,{'ref':('Lab Investigation Number',),'open':('Date Raised',),'close':('Date Closed',),'status':('Status',),'severity':('OOS / OOT',),'summary':('Details of OOS / OOT',),'root':('Cause of OOS / OOT',)}),
'GMP Complaints':('2026 - QMS Data GMP COMPLAINT Version 2.0 (to be Validated).json',2,{'ref':('Complaint Reference',),'open':('Date Complaint Raised',),'close':('Investigation Closed Date',),'target':('Target Date ',),'status':('Status',),'severity':('Classification of Complaint ',),'reclass':('Reclassification of Complaint ',),'area':('Complaint Area (Processes and activities)',),'summary':('Brief Summary ','Brief  Summary '),'root':('Root Cause(s) Identified',)}),
'Supplier Complaints':('2026 - QMS DATA Complaints Against Suppliers (to be Validated).json',1,{'ref':('Complaint Reference',),'open':('Date Logged',),'close':('Date Closed',),'target':('Target Date',),'status':('Status',),'severity':('Nature of Complaint (Critical/ Major/ Minor)',),'summary':('Details of Complaint',)}),
}

analysis={'period':'April 26 - Jun 26','systems':{}}
for name,(fn,hr,cfg) in configs.items():
    _,rr=rows_from(fn,header_row=hr)
    dedupe=name in {'Incidents'}
    analysis['systems'][name]=summarize(name,rr,cfg,dedupe=dedupe)

# MDR: source is effectively empty; report zero if no usable refs
try:
    _,rr=rows_from('2026 - QMS DATA MDR Log (to be Validated).json',header_row=1)
    usable=[r for r in rr if any(str(v).strip() for k,v in r.items() if k!='_row')]
except Exception: usable=[]
analysis['systems']['MDR']={'name':'MDR','counts':{'active_in_period':0,'brought_forward':0,'new_in_period':0,'closed_in_period':0,'open_at_30_jun':0,'overdue_open_at_30_jun':0},'records':[],'note':'No MDR records identified in the source log.'}

# Root-cause trends from reportable deviations, split lines and common separators
roots=Counter(); depts=Counter(); areas=Counter()
for r in analysis['systems']['Deviations']['records']:
    if r.get('root_cause'):
        for x in re.split(r'[\n;/]+',str(r['root_cause'])):
            x=' '.join(x.split()).strip(' .-')
            if x: roots[x]+=1
    if r.get('department'): depts[str(r['department']).strip()]+=1
    if r.get('area'): areas[str(r['area']).strip()]+=1
analysis['deviation_trends']={'root_causes':dict(roots.most_common()),'departments':dict(depts.most_common()),'areas':dict(areas.most_common())}

# OOS causes and complaint areas
for sys,key in [('OOS-OOT','oos_causes'),('GMP Complaints','complaint_root_causes')]:
    c=Counter()
    for r in analysis['systems'][sys]['records']:
        if r.get('root_cause'):
            for x in re.split(r'[\n;/]+',str(r['root_cause'])):
                x=' '.join(x.split()).strip(' .-')
                if x:c[x]+=1
    analysis[key]=dict(c.most_common())

OUT.write_text(json.dumps(analysis,indent=2,ensure_ascii=False),encoding='utf-8')

# compact reference outline from the detailed inspection
try:
    ref=json.loads(Path('analysis/reference_structure.json').read_text(encoding='utf-8'))
    outline=[]
    def walk(sh, texts, tables, charts):
        if sh.get('text'): texts.append(sh['text'])
        if sh.get('table'): tables.append(sh['table'])
        if sh.get('chart'): charts.append(sh['chart'])
        for c in sh.get('children',[]): walk(c,texts,tables,charts)
    for sl in ref['slides']:
        texts=[]; tables=[]; charts=[]
        for sh in sl['shapes']: walk(sh,texts,tables,charts)
        outline.append({'slide':sl['slide'],'texts':texts,'tables':tables,'charts':charts})
    OUTLINE.write_text(json.dumps(outline,indent=2,ensure_ascii=False),encoding='utf-8')
except Exception as e:
    print('outline skipped',e)

print('wrote',OUT,'and',OUTLINE)
