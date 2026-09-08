import json,re
from collections import Counter,defaultdict
from datetime import datetime,date
from pathlib import Path

EX=Path('analysis/extracted'); OUT=Path('analysis/qmr_analysis.json'); OUTLINE=Path('analysis/reference_outline.json')
Q1=(date(2026,1,1),date(2026,3,31)); Q2=(date(2026,4,1),date(2026,6,30))

def dt(v):
    if not v or not isinstance(v,str): return None
    try:return datetime.strptime(v[:10],'%Y-%m-%d').date()
    except:return None

def norm_status(v): return str(v or '').strip().lower().replace(' ','').replace('-','')
def is_done(v): return norm_status(v) in {'complete','completed','closed','yes','y'}
def base_ref(v): return re.sub(r'/[A-Z0-9]+$','',str(v or '').strip(),flags=re.I)
def load(name): return json.loads((EX/name).read_text(encoding='utf-8'))

def rows_from(name,header_row):
    sh=load(name)['sheets'][0]; rows=sh['rows']; headers=next(r['values'] for r in rows if r['row']==header_row)
    out=[]
    for r in rows:
        if r['row']<=header_row: continue
        vals=r['values']+['']*max(0,len(headers)-len(r['values']))
        rec={str(h).strip():vals[i] for i,h in enumerate(headers) if str(h).strip()}; rec['_row']=r['row']; out.append(rec)
    return out

def get(rec,*names):
    for n in names:
        if n in rec:return rec[n]
    return ''

def classify(records,cfg,start,end):
    out=[]
    for r in records:
        ref=str(get(r,*cfg['ref'])).strip()
        if not ref: continue
        od=dt(get(r,*cfg['open']))
        if not od or od>end: continue
        cd=dt(get(r,*cfg.get('close',()))) if cfg.get('close') else None
        target=dt(get(r,*cfg.get('target',()))) if cfg.get('target') else None
        status=get(r,*cfg.get('status',()))
        if cfg.get('close'):
            active=(cd is None or cd>=start); open_end=(cd is None or cd>end); closed_period=bool(cd and start<=cd<=end)
        else:
            active=(od>=start) or (not is_done(status)); open_end=not is_done(status); closed_period=(od>=start and is_done(status))
        if not active: continue
        x=dict(r); x.update({'_ref':ref,'_base':base_ref(ref),'_od':od,'_cd':cd,'_target':target,'_new':start<=od<=end,'_closed':closed_period,'_open_end':open_end,'_carried':od<start,'_overdue':bool(open_end and target and target<end)})
        out.append(x)
    return out

def group_parent(rows,cfg):
    groups=defaultdict(list)
    for r in rows: groups[r['_base']].append(r)
    merged=[]
    for key,g in groups.items():
        x=dict(g[0])
        x['_ref']=key
        x['_od']=min(z['_od'] for z in g)
        cds=[z['_cd'] for z in g if z['_cd']]
        # parent closes only when every action is closed; then parent close date is the latest action close
        all_closed=all(z['_cd'] is not None for z in g) if cfg.get('close') else all(is_done(get(z,*cfg.get('status',()))) for z in g)
        x['_cd']=max(cds) if all_closed and cds else None
        x['_new']=any(z['_new'] for z in g)
        x['_carried']=all(z['_carried'] for z in g)
        x['_open_end']=any(z['_open_end'] for z in g)
        x['_closed']=not x['_open_end'] and any(z['_closed'] for z in g)
        x['_overdue']=any(z['_overdue'] for z in g)
        targets=[z['_target'] for z in g if z['_target']]
        x['_target']=max(targets) if targets else None
        # collapse descriptions/severities
        for field in ('summary','severity','department','area','root','reclass'):
            vals=[]
            for z in g:
                v=str(get(z,*cfg.get(field,()))).strip()
                if v and v not in vals: vals.append(v)
            x['_'+field]=' | '.join(vals)
        merged.append(x)
    return merged

def period_summary(records,cfg,start,end,dedupe=False):
    raw=classify(records,cfg,start,end); rs=group_parent(raw,cfg) if dedupe else raw
    def fval(r,field):
        if dedupe and '_'+field in r:return r['_'+field]
        return get(r,*cfg.get(field,()))
    months=Counter(r['_od'].strftime('%Y-%m') for r in rs if r['_new'])
    classes=Counter(str(fval(r,'severity')).strip() for r in rs if str(fval(r,'severity')).strip())
    new_rs=[r for r in rs if r['_new']]; closed_rs=[r for r in rs if r['_closed']]
    avg_close=None
    close_days=[]
    for r in closed_rs:
        if r['_cd']: close_days.append((r['_cd']-r['_od']).days)
    if close_days: avg_close=round(sum(close_days)/len(close_days),1)
    sig=[]
    for r in rs:
        s=(str(fval(r,'severity'))+' '+str(fval(r,'reclass'))).lower()
        if 'major' in s or 'critical' in s:
            sig.append({'ref':r['_ref'],'date':r['_od'].isoformat(),'summary':fval(r,'summary'),'classification':fval(r,'severity'),'reclassification':fval(r,'reclass'),'target':r['_target'].isoformat() if r['_target'] else '', 'completion':r['_cd'].isoformat() if r['_cd'] else '', 'status_end':'Open' if r['_open_end'] else 'Closed','root_cause':fval(r,'root')})
    return {'counts':{'active':len(rs),'brought_forward':sum(r['_carried'] for r in rs),'raised':len(new_rs),'closed':len(closed_rs),'open_end':sum(r['_open_end'] for r in rs),'overdue_open':sum(r['_overdue'] for r in rs),'major_critical_raised':sum(('major' in str(fval(r,'severity')).lower() or 'critical' in str(fval(r,'severity')).lower()) for r in new_rs),'major_critical_closed':sum(('major' in str(fval(r,'severity')).lower() or 'critical' in str(fval(r,'severity')).lower()) for r in closed_rs)},'avg_closure_days':avg_close,'new_by_month':dict(months),'classification':dict(classes),'significant':sig,'records':[{'ref':r['_ref'],'opened':r['_od'].isoformat(),'closed':r['_cd'].isoformat() if r['_cd'] else '', 'target':r['_target'].isoformat() if r['_target'] else '', 'status_end':'Open' if r['_open_end'] else 'Closed','summary':fval(r,'summary'),'classification':fval(r,'severity'),'department':fval(r,'department'),'area':fval(r,'area'),'root_cause':fval(r,'root'),'overdue':r['_overdue']} for r in rs]}

configs={
'CAPA':('2026 - QMS DATA CAPA Version 1.1.json',2,True,{'ref':('CAPA Reference',),'open':('Date Logged',),'close':('Date Closed',),'target':('Target date',),'status':('Status',),'severity':('Category (Minor / Major / Critical)',),'summary':('Assignable Cause',),'root':('Assignable Cause',)}),
'Deviations':('2026 - QMS DATA DEVIATIONS Version 2.0 (to be Validated).json',2,False,{'ref':('Deviation Reference',),'open':('Date Deviation Raised',),'close':('Investigation Closed',),'target':('Target Date',),'status':('Status',),'severity':('Classification of Deviation',),'reclass':('Reclassification of Deviation',),'department':('Department/Area Affected',),'area':('Deviation Area (Processes and activities)',),'summary':('Brief  Summary','Brief Summary'),'root':('Root Cause(s) Identified',)}),
'Change Controls':('2026 - QMS DATA CHANGE CONTROLS Version 1.0.json',2,True,{'ref':('Change Control Number',),'open':('Date Logged',),'close':('Completion Date',),'target':('Target Date:',),'status':('Current Status',),'severity':('Classification',),'area':('Change Type',),'summary':('Change Summary',)}),
'Document Change Requests':('2026 - QMS DATA DOCUMENT CHANGE REQUEST Version 1.0.json',2,False,{'ref':('Change Request Number',),'open':('Date Raised',),'status':('Current Status',),'severity':('Change Type:',),'summary':('Details of change',)}),
'Effectiveness Checks':('2026 - QMS DATA EFFECTIVENESS CHECKS Version 1.0.json',2,True,{'ref':('CC Reference ',),'open':('Date CC Raised',),'close':('Date Closed',),'target':('Target date',),'status':('Status',),'summary':('Effectiveness check action(s)',)}),
'FMEA':('2026 - QMS DATA FMEA (to be Validated).json',1,False,{'ref':('FMEA Number',),'open':('Date Opened',),'close':('Date Closed',),'target':('Deadline',),'status':('Status',),'severity':('Risk Category',),'summary':('Title',),'root':('Potential failure mode risk',)}),
'Incidents':('2026 - QMS DATA Incident Management Log (to be Validated).json',1,True,{'ref':('Incident Reference',),'open':('Date Logged',),'close':('Incident report Completion Date',),'target':('Incident report Due Date',),'status':('Status',),'severity':('Criticality of Incident (Critical/ Major)',),'summary':('Details of Incident',)}),
'OOS-OOT':('2026 - QMS DATA OOS-OOT Version 1.0.json',2,False,{'ref':('Lab Investigation Number',),'open':('Date Raised',),'close':('Date Closed',),'status':('Status',),'severity':('OOS / OOT',),'summary':('Details of OOS / OOT',),'root':('Cause of OOS / OOT',)}),
'GMP Complaints':('2026 - QMS Data GMP COMPLAINT Version 2.0 (to be Validated).json',2,False,{'ref':('Complaint Reference',),'open':('Date Complaint Raised',),'close':('Investigation Closed Date',),'target':('Target Date ',),'status':('Status',),'severity':('Classification of Complaint ',),'reclass':('Reclassification of Complaint ',),'area':('Complaint Area (Processes and activities)',),'summary':('Brief Summary ','Brief  Summary '),'root':('Root Cause(s) Identified',)}),
'Supplier Complaints':('2026 - QMS DATA Complaints Against Suppliers (to be Validated).json',1,False,{'ref':('Complaint Reference',),'open':('Date Logged',),'close':('Date Closed',),'target':('Target Date',),'status':('Status',),'severity':('Nature of Complaint (Critical/ Major/ Minor)',),'summary':('Details of Complaint',)}),
}
analysis={'period':'April 26 - Jun 26','systems':{}}
for name,(fn,hr,dedupe,cfg) in configs.items():
    records=rows_from(fn,hr)
    analysis['systems'][name]={'Q1':period_summary(records,cfg,*Q1,dedupe=dedupe),'Q2':period_summary(records,cfg,*Q2,dedupe=dedupe)}
analysis['systems']['MDR']={'Q1':{'counts':{'raised':0,'closed':0,'open_end':0}},'Q2':{'counts':{'raised':0,'closed':0,'open_end':0}},'note':'No MDR records identified.'}
# Q2 deviation trends
roots=Counter(); depts=Counter(); areas=Counter()
for r in analysis['systems']['Deviations']['Q2']['records']:
    for x in re.split(r'[\n;/]+',str(r.get('root_cause',''))):
        x=' '.join(x.split()).strip(' .-')
        if x: roots[x]+=1
    if r.get('department'): depts[r['department']]+=1
    if r.get('area'): areas[r['area']]+=1
analysis['deviation_trends_q2']={'root_causes':dict(roots.most_common()),'departments':dict(depts.most_common()),'areas':dict(areas.most_common())}
OUT.write_text(json.dumps(analysis,indent=2,ensure_ascii=False),encoding='utf-8')
# compact outline
try:
    ref=json.loads(Path('analysis/reference_structure.json').read_text(encoding='utf-8')); outline=[]
    def walk(sh,texts,tables,charts):
        if sh.get('text'):texts.append(sh['text'])
        if sh.get('table'):tables.append(sh['table'])
        if sh.get('chart'):charts.append(sh['chart'])
        for c in sh.get('children',[]):walk(c,texts,tables,charts)
    for sl in ref['slides']:
        texts=[];tables=[];charts=[]
        for sh in sl['shapes']:walk(sh,texts,tables,charts)
        outline.append({'slide':sl['slide'],'texts':texts,'tables':tables,'charts':charts})
    OUTLINE.write_text(json.dumps(outline,indent=2,ensure_ascii=False),encoding='utf-8')
except Exception as e: print('outline skipped',e)
print('wrote analysis')
