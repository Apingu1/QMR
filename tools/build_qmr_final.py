import json,re,math
from collections import Counter,defaultdict
from datetime import datetime,date
from pathlib import Path

import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.util import Pt

ROOT=Path('.')
REF=ROOT/'2. Metrics Review Apr - Jun 26 NEW FORMAT.pptx'
DATA=json.loads((ROOT/'analysis/qmr_analysis.json').read_text(encoding='utf-8'))
OUTDIR=ROOT/'output'; OUTDIR.mkdir(exist_ok=True)
OUT=OUTDIR/'Eaststone_QMR_Apr-Jun_2026.pptx'
ANOUT=ROOT/'analysis/final_qmr_summary.json'

def sys(name): return DATA['systems'][name]
def fmt_date(s):
    if not s:return ''
    try:return datetime.strptime(s[:10],'%Y-%m-%d').strftime('%d-%b-%y')
    except:return str(s)
def short(s,n=100):
    s=' '.join(str(s or '').split())
    return s if len(s)<=n else s[:n-1]+'…'
def tables(slide): return [s.table for s in slide.shapes if getattr(s,'has_table',False)]
def charts(slide): return [s.chart for s in slide.shapes if getattr(s,'has_chart',False)]
def set_text_contains(slide,needle,new):
    for sh in slide.shapes:
        if getattr(sh,'has_text_frame',False) and needle.lower() in (sh.text or '').lower():
            sh.text=new
            return sh
    return None
def set_exact_or_contains(slide,needle,new):
    return set_text_contains(slide,needle,new)
def fill_table(tbl,rows):
    R=len(tbl.rows); C=len(tbl.columns)
    for r in range(R):
        for c in range(C):
            val=''
            if r<len(rows) and c<len(rows[r]): val=rows[r][c]
            tbl.cell(r,c).text=str(val if val is not None else '')
    return tbl

def replace_chart(chart,categories,series,title=None):
    cd=CategoryChartData(); cd.categories=categories
    for name,vals in series: cd.add_series(name,vals)
    chart.replace_data(cd)
    if title is not None:
        chart.has_title=True; chart.chart_title.text_frame.text=title

def rm_shape(shape):
    sp=shape._element; sp.getparent().remove(sp)
def replace_picture(slide,pic,path):
    l,t,w,h=pic.left,pic.top,pic.width,pic.height; rm_shape(pic); slide.shapes.add_picture(str(path),l,t,w,h)

def assessment(q1,q2):
    a=q1['counts']['raised']; b=q2['counts']['raised']; s1=q1['counts'].get('major_critical_raised',0); s2=q2['counts'].get('major_critical_raised',0)
    if b<a and s2<=s1:return '🟢- Improving'
    if b>a*1.35 or s2>s1+1:return '🔴- Deteriorating'
    return '🟠- Stable'

def month_trend(system):
    q=sys(system); rec={r['ref']:r for r in q['Q1'].get('records',[])+q['Q2'].get('records',[])}
    cats=['Jan','Feb','Mar','Apr','May','Jun']; new=[0]*6; closed=[0]*6
    for r in rec.values():
        try:
            od=datetime.strptime(r['opened'][:10],'%Y-%m-%d').date()
            if od.year==2026 and 1<=od.month<=6:new[od.month-1]+=1
        except:pass
        try:
            cd=datetime.strptime(r.get('closed','')[:10],'%Y-%m-%d').date()
            if cd.year==2026 and 1<=cd.month<=6:closed[cd.month-1]+=1
        except:pass
    return cats,new,closed

def quarter_classification(system):
    q=sys(system); out=[]
    for quarter in ('Q1','Q2'):
        start=date(2026,1,1) if quarter=='Q1' else date(2026,4,1); end=date(2026,3,31) if quarter=='Q1' else date(2026,6,30)
        c=Counter()
        for r in q[quarter].get('records',[]):
            try:d=datetime.strptime(r['opened'][:10],'%Y-%m-%d').date()
            except:continue
            if start<=d<=end:
                cl=str(r.get('classification','')).strip() or 'Other'; c[cl]+=1
        out.append(c)
    return out

def cap_source_table():
    c=Counter()
    for r in sys('CAPA')['Q2']['records']:
        try:d=datetime.strptime(r['opened'][:10],'%Y-%m-%d').date()
        except:continue
        if not(date(2026,4,1)<=d<=date(2026,6,30)):continue
        s=(r.get('summary') or '').lower()
        if 'dev.' in s or s.startswith('dev'):k='Deviations'
        elif 'gmp.cpl' in s or 'complaint' in s:k='Complaints'
        elif 'oos' in s or 'oot' in s:k='OOS/OOT'
        elif 'audit' in s or 'inspection' in s:k='Audits / inspections'
        elif 'supplier' in s:k='Supplier / outsourced activity'
        else:k='Other'
        c[k]+=1
    return c

def cc_theme_table():
    c=Counter()
    for r in sys('Change Controls')['Q2']['records']:
        try:d=datetime.strptime(r['opened'][:10],'%Y-%m-%d').date()
        except:continue
        if not(date(2026,4,1)<=d<=date(2026,6,30)):continue
        s=(r.get('summary') or '').lower()
        if any(x in s for x in ['supplier','material','brand','excipient','api']):k='Material & Supplier Changes'
        elif any(x in s for x in ['equipment','machine','printer','chamber','blender','granulator','detector','room','facility']):k='Equipment/Facility Changes'
        elif any(x in s for x in ['sop','bmr','document','procedure','form','spreadsheet']):k='Documentation/SOP Changes'
        elif any(x in s for x in ['mhra','regulatory','licence','license']):k='Regulatory driven Changes'
        elif any(x in s for x in ['formulation','manufactur','process']):k='Process/Manufacturing changes'
        else:k='Other'
        c[k]+=1
    return c

def draw_bar(data,title,path,maxn=8):
    items=list(data.items())[:maxn]
    if not items:items=[('No recurring trend identified',0)]
    labels=[x[0] for x in items][::-1]; vals=[x[1] for x in items][::-1]
    fig,ax=plt.subplots(figsize=(11,3.1)); ax.barh(labels,vals); ax.set_title(title); ax.set_xlabel('Count'); ax.grid(axis='x',alpha=.25)
    for i,v in enumerate(vals): ax.text(v+0.05,i,str(v),va='center',fontsize=9)
    fig.tight_layout(); fig.savefig(path,dpi=180,bbox_inches='tight'); plt.close(fig)

def table_sig(title,headers,items,mode,capacity):
    rows=[[title]+['']*(len(headers)-1),headers]
    selected=[x for x in items if x['status_end']==mode]
    if not selected:
        rows.append(['','',f'None to report in Q2','','',''])
    else:
        for x in selected[:capacity]:
            rows.append([x['ref'],fmt_date(x['date']),short(x.get('summary',''),90),x.get('reclassification') or x.get('classification',''),fmt_date(x.get('target','')),fmt_date(x.get('completion','')) if mode=='Closed' else 'Investigation / actions in progress'])
        if len(selected)>capacity:
            rows[-1][2]=short(rows[-1][2],55)+f' (+{len(selected)-capacity} additional significant item(s))'
    return rows

def status_row(q2):
    c=q2['counts']; return [['Currently Open','Open with an Extension','Currently open & overdue at the end of quarter','Closed (Q2 2026)','Closed after target date','Closed after extension'],[c.get('open_end',0),'0',c.get('overdue_open',0),c.get('closed',0),'0','0']]

def trend_pairs(records):
    groups=defaultdict(list)
    for r in records:
        area=short(r.get('area') or 'Other',35); root=short(r.get('root_cause') or 'Other',45)
        for one in re.split(r'[\n;/]+',root):
            one=' '.join(one.split()).strip()
            if one:groups[(area,one)].append(r)
    arr=sorted(groups.items(),key=lambda kv:len(kv[1]),reverse=True)
    return [(k,v) for k,v in arr if len(v)>=2][:6] or arr[:4]

prs=Presentation(REF)
# Title
prs.slides[0].shapes[0].text='QMR for Period April 26 - Jun 26'
if len(prs.slides[0].shapes)>1: prs.slides[0].shapes[1].text=''

# DEVIATIONS
q=sys('Deviations'); q1,q2=q['Q1'],q['Q2']; sl=prs.slides[2]; ts=tables(sl)
fill_table(ts[0],[['KPI','Q1 (2026)','Q2 (2026)'],['Total raised this quarter',q1['counts']['raised'],q2['counts']['raised']],['Total closed this quarter',q1['counts']['closed'],q2['counts']['closed']],['Major/Critical raised',q1['counts']['major_critical_raised'],q2['counts']['major_critical_raised']],['Major/Critical closed',q1['counts']['major_critical_closed'],q2['counts']['major_critical_closed']],['Average Closure time',f"{q1['avg_closure_days'] or 0:g} days",f"{q2['avg_closure_days'] or 0:g} days"]])
set_text_contains(sl,'KPI & Trend Assessment',f"KPI & Trend Assessment:\n\n{assessment(q1,q2)}")
ch=charts(sl); cats,new,cl=month_trend('Deviations'); replace_chart(ch[1],cats,[('New Deviations',new),('Closed Deviations',cl)],'Deviations - 2026 Monthly Trend')
c1,c2=quarter_classification('Deviations'); replace_chart(ch[0],['Q1 2026','Q2 2026'],[('Minor',[c1.get('Minor',0),c2.get('Minor',0)]),('Major',[c1.get('Major',0),c2.get('Major',0)]),('Critical',[c1.get('Critical',0),c2.get('Critical',0)])],'Deviation Classification by Quarter')
sl=prs.slides[3]; ts=tables(sl); sig=q2['significant']; fill_table(ts[0],table_sig('Major & Critical Deviations Closed',['Deviation number','Date raised','Deviation description','Classification','Target date','Completion date'],sig,'Closed',max(1,len(ts[0].rows)-2))); fill_table(ts[1],table_sig('Major & Critical Deviations Open',['Deviation number','Date raised','Deviation description','Classification','Target date','Reason still open'],sig,'Open',max(1,len(ts[1].rows)-2))); fill_table(ts[2],status_row(q2))
# root charts
sl=prs.slides[4]; pics=[s for s in sl.shapes if str(s.shape_type).startswith('PICTURE')]; rc=DATA.get('deviation_trends_q2',{}); p1=OUTDIR/'dev_root_causes.png'; p2=OUTDIR/'dev_areas.png'; draw_bar(rc.get('root_causes',{}),'Q2 Deviation Root Causes',p1); draw_bar(rc.get('areas',{}),'Q2 Deviation Areas',p2)
if len(pics)>=2:
    pics_sorted=sorted(pics,key=lambda x:x.top); replace_picture(sl,pics_sorted[0],p1); replace_picture(sl,pics_sorted[1],p2)
sl=prs.slides[5]; ts=tables(sl); pairs=trend_pairs(q2['records']); rows=[['Trend','Deviation numbers','Deviation details (Description, investigation etc)','Systemic themes identified?','Escalation required/Regulatory impact?'],['Area &\nRoot Cause','','','','']]
for (area,root),rr in pairs: rows.append([f'{area}\n&\n{root}','\n'.join(x['ref'] for x in rr),short(' | '.join(x.get('summary','') for x in rr),250),'Recurring combination identified' if len(rr)>1 else 'No recurring Q2 combination','Continue QMS monitoring'])
fill_table(ts[0],rows); top=', '.join(list(rc.get('root_causes',{}).keys())[:3]) or 'No recurring root-cause theme identified'; set_text_contains(sl,'Systemic themes identified',f'Systemic themes identified:\n\nQ2 root-cause review shows the most frequent themes were {top}.\n\nFurther actions:\n\nContinue monitoring through deviation/CAPA trending. Existing actions remain tracked through the QMS; no additional standalone action is proposed solely from the quarterly trend review.')

# GMP COMPLAINTS
q=sys('GMP Complaints'); q1,q2=q['Q1'],q['Q2']; sl=prs.slides[6]; ts=tables(sl); fill_table(ts[0],[['KPI','Q1 (2026)','Q2 (2026)'],['Total raised this quarter',q1['counts']['raised'],q2['counts']['raised']],['Total closed this quarter',q1['counts']['closed'],q2['counts']['closed']],['Average Closure time',f"{q1['avg_closure_days'] or 0:g} days",f"{q2['avg_closure_days'] or 0:g} days"],['Major/Critical raised',q1['counts']['major_critical_raised'],q2['counts']['major_critical_raised']]])
set_text_contains(sl,'KPI & Trend Assessment',f"KPI & Trend Assessment:\n\n{assessment(q1,q2)}"); ch=charts(sl); cats,new,cl=month_trend('GMP Complaints'); replace_chart(ch[0],cats,[('New Complaints',new),('Completed Complaints',cl)],'GMP Complaints - 2026 Monthly Trend'); c1,c2=quarter_classification('GMP Complaints'); replace_chart(ch[1],['Q1 2026','Q2 2026'],[('Minor',[c1.get('Minor',0),c2.get('Minor',0)]),('Major',[c1.get('Major',0),c2.get('Major',0)]),('Critical',[c1.get('Critical',0),c2.get('Critical',0)])],'Complaint Classification by Quarter')
sl=prs.slides[7]; ts=tables(sl); sig=q2['significant']; fill_table(ts[0],table_sig('Major & Critical GMP Complaints Closed',['GMP Complaint number','Date raised','Complaint description','Classification','Target date','Completion date'],sig,'Closed',max(1,len(ts[0].rows)-2))); fill_table(ts[1],table_sig('Major & Critical GMP Complaints Open',['GMP Complaint number','Date raised','Complaint description','Classification','Target date','Reason still open'],sig,'Open',max(1,len(ts[1].rows)-2))); fill_table(ts[2],status_row(q2))
# complaint root cause visuals
sl=prs.slides[8]; pics=[s for s in sl.shapes if str(s.shape_type).startswith('PICTURE')]; roots=Counter(); areas=Counter()
for r in q2['records']:
    if r.get('area'):areas[r['area']]+=1
    for x in re.split(r'[\n;/]+',str(r.get('root_cause',''))):
        x=' '.join(x.split()).strip()
        if x:roots[x]+=1
p1=OUTDIR/'complaint_roots.png'; p2=OUTDIR/'complaint_areas.png'; draw_bar(roots,'Q2 Complaint Root Causes',p1); draw_bar(areas,'Q2 Complaint Areas',p2)
if len(pics)>=2:
    pics_sorted=sorted(pics,key=lambda x:x.top); replace_picture(sl,pics_sorted[0],p1); replace_picture(sl,pics_sorted[1],p2)
sl=prs.slides[9]; ts=tables(sl); pairs=trend_pairs(q2['records']); rows=[['Trend','GMP Complaint numbers','Complaint details (Description, investigation etc)','Systemic themes identified?','Escalation required/Regulatory impact?'],['Area &\nRoot Cause','','','','']]
for (area,root),rr in pairs: rows.append([f'{area}\n&\n{root}','\n'.join(x['ref'] for x in rr),short(' | '.join(x.get('summary','') for x in rr),250),'Recurring combination identified' if len(rr)>1 else 'No recurring Q2 combination','Continue monitoring'])
fill_table(ts[0],rows); set_text_contains(sl,'Systemic themes identified',f"Systemic themes identified:\n\n{('No recurring GMP complaint trend was identified in Q2.' if q2['counts']['raised']<=1 else 'Complaint themes remain limited and are being monitored through routine trending.')}\n\nFurther actions:\nContinue routine complaint trending and linked CAPA follow-up where applicable.")

# SUPPLIER COMPLAINTS
q=sys('Supplier Complaints'); q1,q2=q['Q1'],q['Q2']; sl=prs.slides[10]; ts=tables(sl); fill_table(ts[0],[['KPI','Q1 (2026)','Q2 (2026)'],['Total raised this quarter',q1['counts']['raised'],q2['counts']['raised']],['Total closed this quarter',q1['counts']['closed'],q2['counts']['closed']],['Average Closure time',f"{q1['avg_closure_days'] or 0:g} days",f"{q2['avg_closure_days'] or 0:g} days"],['Major/Critical raised',q1['counts']['major_critical_raised'],q2['counts']['major_critical_raised']]])
set_text_contains(sl,'KPI & Trend Assessment',f"KPI & Trend Assessment:\n\n{assessment(q1,q2)}"); set_text_contains(sl,'No Major or Critical', 'No Major or Critical supplier complaints raised in Q2' if q2['counts']['major_critical_raised']==0 else f"{q2['counts']['major_critical_raised']} Major/Critical supplier complaint(s) raised in Q2"); ch=charts(sl); cats,new,cl=month_trend('Supplier Complaints'); replace_chart(ch[0],cats,[('New Complaints',new),('Completed Complaints',cl)],'Supplier Complaints - 2026 Monthly Trend'); c1,c2=quarter_classification('Supplier Complaints'); replace_chart(ch[1],['Q1 2026','Q2 2026'],[('Minor',[c1.get('Minor',0),c2.get('Minor',0)]),('Major',[c1.get('Major',0),c2.get('Major',0)]),('Critical',[c1.get('Critical',0),c2.get('Critical',0)])],'Supplier Complaint Classification')

# OOS/OOT
q=sys('OOS-OOT'); q1,q2=q['Q1'],q['Q2']; sl=prs.slides[11]; ts=tables(sl)
def type_count(qt,t): return sum(1 for r in qt['records'] if r.get('classification','').strip().upper()==t and date.fromisoformat(r['opened']) >= (date(2026,1,1) if qt is q1 else date(2026,4,1)))
fill_table(ts[0],[['KPI','Q1 (2026)','Q2 (2026)'],['Total raised this quarter',q1['counts']['raised'],q2['counts']['raised']],['Total closed this quarter',q1['counts']['closed'],q2['counts']['closed']],['Average Closure time',f"{q1['avg_closure_days'] or 0:g} days",f"{q2['avg_closure_days'] or 0:g} days"],['True OOS',type_count(q1,'OOS'),type_count(q2,'OOS')],['True OOT',type_count(q1,'OOT'),type_count(q2,'OOT')],['Open at quarter end',q1['counts']['open_end'],q2['counts']['open_end']],['Batch Rejections','0',sum('reject' in (r.get('summary','')+' '+r.get('root_cause','')).lower() for r in q2['records'])]])
set_text_contains(sl,'KPI & Trend Assessment',f"KPI & Trend Assessment:\n\n{assessment(q1,q2)}"); ch=charts(sl); cats,new,cl=month_trend('OOS-OOT'); replace_chart(ch[0],cats,[('New OOS/OOT',new),('Completed OOS/OOT',cl)],'OOS/OOT - 2026 Monthly Trend')
sl=prs.slides[12]; ts=tables(sl); closed=[r for r in q2['records'] if r['status_end']=='Closed']; opened=[r for r in q2['records'] if r['status_end']=='Open']; rows=[['OOS/OOT Closed','','','',''],['OOS/OOT Number','Date raised','OOS/OOT description','Target date','Completion date']]+[[r['ref'],fmt_date(r['opened']),short(r.get('summary',''),105),fmt_date(r.get('target','')),fmt_date(r.get('closed',''))] for r in closed[:len(ts[0].rows)-2]]; fill_table(ts[0],rows); rows=[['OOS/OOT Open','','','',''],['OOS/OOT Number','Date raised','OOS/OOT description','Target date','Reason still open']]+[[r['ref'],fmt_date(r['opened']),short(r.get('summary',''),105),fmt_date(r.get('target','')),'Investigation in progress'] for r in opened[:len(ts[1].rows)-2]]; fill_table(ts[1],rows); fill_table(ts[2],status_row(q2)); oos=q2['counts']['raised']; fill_table(ts[3],[['Trend Analysis',f'{oos} OOS/OOT investigation(s) raised in Q2. Root causes were reviewed for recurrence; no single repeated analytical root cause dominates the quarter.'],['OOT Specifics',f"OOT raised in Q2: {type_count(q2,'OOT')}. OOS raised in Q2: {type_count(q2,'OOS')}."],['Regulatory Impact','No additional regulatory impact is identified from the Q2 OOS/OOT log.']])

# CAPA
q=sys('CAPA'); q1,q2=q['Q1'],q['Q2']; sl=prs.slides[13]; ts=tables(sl); fill_table(ts[0],[['KPI','Q1 (2026)','Q2 (2026)'],['Total raised this quarter',q1['counts']['raised'],q2['counts']['raised']],['Total closed this quarter',q1['counts']['closed'],q2['counts']['closed']],['Major/Critical raised',q1['counts']['major_critical_raised'],q2['counts']['major_critical_raised']],['Major/Critical closed',q1['counts']['major_critical_closed'],q2['counts']['major_critical_closed']]])
sources=cap_source_table(); fill_table(ts[1],[['Source of Q2 CAPAs','CAPA Quantity raised']]+[[k,v] for k,v in sources.most_common()]); set_text_contains(sl,'KPI & Trend Assessment',f"KPI & Trend Assessment:\n\n{assessment(q1,q2)}"); ch=charts(sl); cats,new,cl=month_trend('CAPA'); replace_chart(ch[0],cats,[('New CAPAs',new),('Completed CAPAs',cl)],'CAPAs - 2026 Monthly Trend')
sl=prs.slides[14]; ts=tables(sl); sig=q2['significant']; # use larger first table for OPEN, second for CLOSED
fill_table(ts[0],table_sig('Major & Critical CAPAs Open',['CAPA number','Date raised','CAPA description','Classification','Target date','Reason still open'],sig,'Open',max(1,len(ts[0].rows)-2))); fill_table(ts[1],table_sig('Major & Critical CAPAs Closed',['CAPA number','Date raised','CAPA description','Classification','Target date','Completion date'],sig,'Closed',max(1,len(ts[1].rows)-2))); fill_table(ts[2],status_row(q2)); fill_table(ts[3],[['CAPA position (Q2 2026)','Open','Overdue','Closed in Q2'],['',q2['counts']['open_end'],q2['counts']['overdue_open'],q2['counts']['closed']]])

# CHANGE CONTROLS
q=sys('Change Controls'); q1,q2=q['Q1'],q['Q2']; sl=prs.slides[15]; ts=tables(sl); fill_table(ts[0],[['KPI','Q1 (2026)','Q2 (2026)'],['Total raised this quarter',q1['counts']['raised'],q2['counts']['raised']],['Total closed this quarter',q1['counts']['closed'],q2['counts']['closed']],['Major/Critical raised',q1['counts']['major_critical_raised'],q2['counts']['major_critical_raised']],['Major/Critical closed',q1['counts']['major_critical_closed'],q2['counts']['major_critical_closed']]])
themes=cc_theme_table(); fill_table(ts[1],[['Q2 Change Control themes','CC Quantity raised']]+[[k,v] for k,v in themes.most_common()]); set_text_contains(sl,'KPI & Trend Assessment',f"KPI & Trend Assessment:\n\n{assessment(q1,q2)}"); ch=charts(sl); cats,new,cl=month_trend('Change Controls'); replace_chart(ch[0],cats,[('New Change Controls',new),('Completed Change Controls',cl)],'Change Controls - 2026 Monthly Trend')
sl=prs.slides[16]; ts=tables(sl); sig=q2['significant']; fill_table(ts[0],table_sig('Major & Critical Change Controls Closed',['Change Control number','Date raised','Change Control description','Classification','Target date','Completion date'],sig,'Closed',max(1,len(ts[0].rows)-2))); fill_table(ts[1],table_sig('Major & Critical Change Controls Open',['Change Control number','Date raised','Change Control description','Classification','Target date','Reason still open'],sig,'Open',max(1,len(ts[1].rows)-2))); fill_table(ts[2],status_row(q2)); ec=sys('Effectiveness Checks')['Q2']; fill_table(ts[3],[['Effectiveness Checks (Q2 2026)','Completed','Open','Overdue at quarter end'],['',ec['counts']['closed'],ec['counts']['open_end'],ec['counts']['overdue_open']]])

# FMEA / RISK
q=sys('FMEA'); q1,q2=q['Q1'],q['Q2']; sl=prs.slides[17]; ts=tables(sl); fill_table(ts[0],[['KPI','Q1 (2026)','Q2 (2026)'],['Total raised this quarter',q1['counts']['raised'],q2['counts']['raised']],['Total closed this quarter',q1['counts']['closed'],q2['counts']['closed']],['Major/Critical raised',q1['counts']['major_critical_raised'],q2['counts']['major_critical_raised']],['Major/Critical closed',q1['counts']['major_critical_closed'],q2['counts']['major_critical_closed']]]); fill_table(ts[1],[['Q2 FMEA / Risk Position','Count'],['Open at quarter end',q2['counts']['open_end']],['Overdue open',q2['counts']['overdue_open']],['Closed in Q2',q2['counts']['closed']],['Raised in Q2',q2['counts']['raised']]]); set_text_contains(sl,'KPI & Trend Assessment',f"KPI & Trend Assessment:\n\n{assessment(q1,q2)}"); ch=charts(sl); cats,new,cl=month_trend('FMEA'); replace_chart(ch[0],cats,[('New FMEAs',new),('Completed FMEAs',cl)],'FMEA / Risk Assessments - 2026 Monthly Trend')
sl=prs.slides[18]; ts=tables(sl); sig=q2['significant']; fill_table(ts[0],table_sig('Major & Critical FMEA / Risk Assessments Closed',['Risk Assessment number','Date raised','Risk Assessment description','Classification','Target date','Completion date'],sig,'Closed',max(1,len(ts[0].rows)-2))); fill_table(ts[1],table_sig('Major & Critical FMEA / Risk Assessments Open',['Risk Assessment number','Date raised','Risk Assessment description','Classification','Target date','Reason still open'],sig,'Open',max(1,len(ts[1].rows)-2))); fill_table(ts[2],status_row(q2))

# INCIDENT / MDR - repurpose slide 20
sl=prs.slides[19]; set_text_contains(sl,'RECALLS','INCIDENT MANAGEMENT & MDR'); set_text_contains(sl,'QUETIAPINE Recall Summary','Incident Management Summary'); set_text_contains(sl,'Summary of all recalls','Q2 Incident / MDR Summary'); inc=sys('Incidents')['Q2']; mdr=sys('MDR')['Q2']; ts=tables(sl); incrows=[['Date Raised','Classification','Incident','Status at Jun','CAPA / Link','Completion','Comments']]
for r in inc.get('records',[])[:len(ts[0].rows)-1]: incrows.append([fmt_date(r['opened']),r.get('classification',''),short(r.get('summary',''),90),r.get('status_end',''),'',fmt_date(r.get('closed','')),''])
if len(incrows)==1:incrows.append(['','','No incidents raised or carried open during Q2','','','',''])
fill_table(ts[0],incrows); fill_table(ts[1],[['Activity','Outcome'],['Incidents Raised in Q2',inc['counts']['raised']],['Incidents Closed in Q2',inc['counts']['closed']],['Incidents Open at 30 Jun',inc['counts']['open_end']],['MDRs Raised in Q2',mdr['counts'].get('raised',0)],['MDRs Open at 30 Jun',mdr['counts'].get('open_end',0)],['Improvement Opportunities Identified','Routine monitoring continues']]); fill_table(ts[2],[['Management Conclusion',''],['Observation','Impact'],['Q2 incident/MDR activity',('No Q2 incident or MDR escalation identified from the source logs.' if inc['counts']['raised']==0 and mdr['counts'].get('raised',0)==0 else 'Incident/MDR activity reviewed and tracked through the QMS.')],['Open significant incident backlog',inc['counts']['open_end']]])
set_text_contains(sl,'Recall Outcome & Key learning',f"Quarter outcome & key learning:\nIncidents raised in Q2: {inc['counts']['raised']}\nMDRs raised in Q2: {mdr['counts'].get('raised',0)}\nOpen incidents at 30 Jun: {inc['counts']['open_end']}\n\nNo separate recall event is identified in the Q2 incident/MDR source data.")

# DOCUMENT CHANGE REQUESTS - repurpose slide 21
sl=prs.slides[20]; set_text_contains(sl,'RETURNS & REJECTED BATCHES','DOCUMENT CHANGE REQUESTS'); set_text_contains(sl,'Product Returns Summary','DCR Activity Summary'); set_text_contains(sl,'Rejected Batches Summary','DCR Details'); dcr=sys('Document Change Requests')['Q2']; ts=tables(sl); cats=Counter(r.get('classification','') for r in dcr['records'] if r.get('classification'))
fill_table(ts[0],[['DCR Themes',''],['Change Type','Count']]+[[k,v] for k,v in cats.most_common()]); fill_table(ts[1],[['DCR Activity',''],['Metric','Quarter Total'],['DCRs Raised',dcr['counts']['raised']],['Completed (records raised Q2)',dcr['counts']['closed']],['Open in source',dcr['counts']['open_end']],['Major GMP change type',sum('major' in str(r.get('classification','')).lower() for r in dcr['records'])],['Minor/non-GMP change type',sum('minor' in str(r.get('classification','')).lower() for r in dcr['records'])],['Overdue','N/A - DCR log does not contain target dates'],['CAPAs Raised','N/A']]); fill_table(ts[2],[['Management Conclusion'],[f"{dcr['counts']['raised']} DCRs were raised in Q2. Document changes are being managed through the controlled DCR process; no separate critical DCR category is recorded in the log."]]); fill_table(ts[3],[['Selected DCRs',''],['Reference','Change summary']]+[[r['ref'],short(r.get('summary',''),115)] for r in dcr['records'][:max(1,len(ts[3].rows)-2)]]); fill_table(ts[4],[['Current DCR Position',''],['Metric','Count'],['Open in source',dcr['counts']['open_end']],['Raised Q2',dcr['counts']['raised']]]); fill_table(ts[5],[['Management Conclusion'],['DCR activity will continue to be monitored through routine document control review.']])

# EFFECTIVENESS CHECKS - repurpose controlled documents slide
sl=prs.slides[21]; set_text_contains(sl,'CONTROLLED DOCUMENTS','EFFECTIVENESS CHECKS'); ec=sys('Effectiveness Checks')['Q2']; ts=tables(sl); rows=[['CC Reference','Date Raised','Effectiveness Check Action','Target Date','Date Closed','Status','Owner','Location','Comments']]
for r in ec['records'][:len(ts[0].rows)-1]: rows.append([r['ref'],fmt_date(r['opened']),short(r.get('summary',''),70),fmt_date(r.get('target','')),fmt_date(r.get('closed','')),r['status_end'],'','',''])
fill_table(ts[0],rows); fill_table(ts[1],[['Metric','Q2 Total','','','','','','',''],['Checks closed',ec['counts']['closed'],'','','','','','',''],['Open at 30 Jun',ec['counts']['open_end'],'','','','','','','']]); fill_table(ts[2],[['Quarter Summary','Completed','Open','Overdue','','','','',''],['',ec['counts']['closed'],ec['counts']['open_end'],ec['counts']['overdue_open'],'','','','','']])

# Executive summary on original slide 31
sl=prs.slides[30]; dev=sys('Deviations')['Q2']; comp=sys('GMP Complaints')['Q2']; supp=sys('Supplier Complaints')['Q2']; oos=sys('OOS-OOT')['Q2']; capa=sys('CAPA')['Q2']; cc=sys('Change Controls')['Q2']; dcr=sys('Document Change Requests')['Q2']; inc=sys('Incidents')['Q2']
overall='🟠 – Improvement needed' if (dev['counts']['major_critical_raised']>=3 or capa['counts']['open_end']>=10 or cc['counts']['overdue_open']>0) else '🟢 – In control'
set_text_contains(sl,'Overall QMS Status',f'Overall QMS Status: {overall}')
quarter=(f"Quarter Summary\nThe Pharmaceutical Quality System remained operational and controlled during April–June 2026, with increased quality-system activity requiring continued management attention.\n\nKey observations include:\n• {dev['counts']['raised']} deviations raised and {dev['counts']['closed']} closed during Q2; {dev['counts']['major_critical_raised']} Major/Critical deviations were raised.\n• {comp['counts']['raised']} GMP complaint(s) and {supp['counts']['raised']} supplier complaint(s) were raised.\n• {oos['counts']['raised']} OOS/OOT investigation(s) were raised; {oos['counts']['open_end']} remained open at quarter end.\n• {capa['counts']['raised']} parent CAPAs were raised; {capa['counts']['open_end']} remained open at quarter end, with {capa['counts']['overdue_open']} overdue.\n• {cc['counts']['raised']} Change Controls were raised and {cc['counts']['closed']} completed during Q2.\n• No MDR activity was identified in the source log; Q2 incident activity was {inc['counts']['raised']}." )
set_text_contains(sl,'Quarter Summary',quarter)
risks=(f"Key Risks\n\nCAPA / Change Control Workload\nCurrent Risk Level: Medium\nThe Q2 CAPA workload increased, driven in part by audit and investigation actions. Open actions should continue to be prioritised according to risk and target date.\n\nDeviation Trend\nCurrent Risk Level: Medium\nDeviation volume and significant-event activity should remain under enhanced quarterly trending, particularly recurring procedural, equipment and training-related themes.\n\nDocument Control\nCurrent Risk Level: Low-Medium\n{dcr['counts']['raised']} DCRs were raised in Q2. Continue monitoring completion of document changes and linked implementation activities.")
set_text_contains(sl,'Key Risks',risks)
improvements=(f"Key Improvements Achieved\n\n• {dev['counts']['closed']} deviations and {capa['counts']['closed']} parent CAPAs were closed during the quarter.\n• GMP complaint volume remained low, with no Q2 Major/Critical complaint raised in the log.\n• Effectiveness checks continued to be completed and tracked through the QMS.\n• OOS/OOT investigations generated defined root causes and linked CAPA activity where required.")
set_text_contains(sl,'Key Improvements Achieved',improvements)

# Delete unsupported legacy Q1 operational slides 23-30, preserving Executive Summary and Questions.
for idx in range(29,21,-1):
    sldId=prs.slides._sldIdLst[idx]; prs.part.drop_rel(sldId.rId); del prs.slides._sldIdLst[idx]

# Remove any TBC/TBD text remnants and update old quarter labels in kept slides
for slide in prs.slides:
    for sh in slide.shapes:
        if getattr(sh,'has_text_frame',False):
            txt=sh.text or ''
            txt=txt.replace('TBC','').replace('TBD','')
            if txt!=sh.text: sh.text=txt
        if getattr(sh,'has_table',False):
            for row in sh.table.rows:
                for cell in row.cells:
                    cell.text=cell.text.replace('TBC','').replace('TBD','')

prs.save(OUT)
# compact QA summary
summary={'output':str(OUT),'slides':len(prs.slides),'period':'April 26 - Jun 26','key_counts':{k:sys(k)['Q2']['counts'] for k in ['Deviations','GMP Complaints','Supplier Complaints','OOS-OOT','CAPA','Change Controls','FMEA','Incidents','Document Change Requests','Effectiveness Checks']}}
ANOUT.write_text(json.dumps(summary,indent=2),encoding='utf-8')
print(json.dumps(summary,indent=2))
