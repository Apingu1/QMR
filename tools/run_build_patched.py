from pathlib import Path

p=Path('tools/build_qmr_final.py')
src=p.read_text(encoding='utf-8')
src=src.replace("def charts(slide): return [s.chart for s in slide.shapes if getattr(s,'has_chart',False)]", "def charts(slide): return [s for s in slide.shapes if getattr(s,'has_chart',False)]")
old="""def replace_chart(chart,categories,series,title=None):
    cd=CategoryChartData(); cd.categories=categories
    for name,vals in series: cd.add_series(name,vals)
    chart.replace_data(cd)
    if title is not None:
        chart.has_title=True; chart.chart_title.text_frame.text=title
"""
new="""def replace_chart(shape,categories,series,title=None):
    old_chart=shape.chart
    chart_type=old_chart.chart_type
    chart_style=old_chart.chart_style
    has_legend=old_chart.has_legend
    l,t,w,h=shape.left,shape.top,shape.width,shape.height
    parent=shape._parent
    cd=CategoryChartData(); cd.categories=categories
    for name,vals in series: cd.add_series(name,vals)
    sp=shape._element; sp.getparent().remove(sp)
    chart=parent.add_chart(chart_type,l,t,w,h,cd).chart
    if chart_style is not None:
        try: chart.chart_style=chart_style
        except Exception: pass
    chart.has_legend=has_legend
    if title is not None:
        chart.has_title=True; chart.chart_title.text_frame.text=title
"""
if old not in src:
    raise SystemExit('Expected replace_chart block not found')
src=src.replace(old,new)
code=compile(src,'tools/build_qmr_final.py','exec')
exec(code,{'__name__':'__main__','__file__':'tools/build_qmr_final.py'})
