"""ДЛЯ ТЕСТОВ. Проверяет неизменную приёмку дуговых корней и фактическое срабатывание нового сравнения прогнозов."""
from pathlib import Path
here=Path(__file__).resolve().parent;file=here/'Check-TestArc56Diagnostics.py';code=file.read_text('utf-8-sig')
code=code.replace('Cold893Arc56Diagnostics_20261126','Cold908ResidualArcPredictor_20261126').replace('Cold888StrictEnvelopeGrowth_20261126','Cold901ConsistentControlRefinement_20261126')
code=code.replace("audit['imported_this_pass']==['modTestCoupledElasticState']","audit['imported_this_pass']==['CCoupledArcContinuation','CCoupledNewtonSolver','modTestCoupledElasticState']")
old="    if path.is_file():assert path.read_bytes()==(root/path.relative_to(parent)).read_bytes()"
new="    if path.is_file() and path.name not in ('CCoupledArcContinuation.cls','CCoupledNewtonSolver.cls'):assert path.read_bytes()==(root/path.relative_to(parent)).read_bytes()"
assert code.count(old)==1;code=code.replace(old,new,1)
anchor="source=run_path(str(here/'Check-TestActualReverseCapture.py'))"
before="""rel='src/Solver/CCoupledNewtonSolver.cls'
assert (root/rel).read_text('utf-8-sig')==(parent/rel).read_text('utf-8-sig')+'\\n\\n'+(here/'CandidateArcPredictorDefects.cls.txt').read_text('utf-8-sig')
import re
def TestMethodMap(path):
    text=path.read_text('utf-8-sig')
    found=re.findall(r'(?ms)^(?:Private |Friend |Public )(?:Function|Sub|Property Get) ([A-Za-z_][A-Za-z_0-9]*)\\(.*?^End (?:Function|Sub|Property)',text)
    result={}
    for name in found:
        expression=r'(?ms)^(?:Private |Friend |Public )(?:Function|Sub|Property Get) '+re.escape(name)+r'\\(.*?^End (?:Function|Sub|Property)'
        result[name]=re.findall(expression,text)[0]
    return result
left=TestMethodMap(parent/'src/Solver/CCoupledArcContinuation.cls');right=TestMethodMap(root/'src/Solver/CCoupledArcContinuation.cls')
assert set(right)==set(left)|{'TestPredictorComparison'}
for name in left:
    if name!='TryStep':assert left[name]==right[name],name
"""
assert code.count(anchor)==1;code=code.replace(anchor,before+anchor,1)
for old,new in [("len(source['fields'])==56","len(source['fields'])==32"),('len(exported)==57','len(exported)==33'),
                ('exported[:56]','exported[:32]'),('TestH(phase,57)','TestH(phase,33)'),('source_states=56','source_states=32'),
                ('Arc56','Arc32'),('ARC56','ARC32'),('all_working_sources_exact888=True','only_predictor_comparison_and_selection_added=True')]:
    assert old in code,old;code=code.replace(old,new)
anchor='    effort=r.VectorTest();found=[]'
insert='''    comparison_header=r.VectorTest();assert len(comparison_header)==1
    comparison=r.VectorTest() if comparison_header[0] else None
    if comparison is not None:
        assert len(comparison)==6 and min(comparison[1:])>=0 and comparison[3]>0
        expected=not(comparison[1]+comparison[3]<comparison[2] and comparison[1]<.25*comparison[2])
        assert bool(comparison[0])==expected
        assert abs(comparison[3]-TestBound(source['lambdas'][-1]))<1e-20
        assert max(comparison[4:])<=1/32+1e-17 and min(comparison[4:])>0
'''
assert code.count(anchor)==1;code=code.replace(anchor,insert+anchor,1)
anchor='        H_guide_called=bool(guides[\'H\'])'
assert code.count(anchor)==1
code=code.replace(anchor,"        predictor_comparison=(comparison.tolist() if comparison is not None else None),\n"+anchor,1)
code=code.replace("root/'Arc32DiagnosticsProof.json'","root/'ResidualArcPredictorProof.json'")
code+='\nprint("PREDICTOR COMPARISONS",[(x["case"],x["predictor_comparison"]) for x in cases])\n'
exec(compile(code,str(file),'exec'),{'__file__':str(file),'__name__':'residual_arc_predictor_test'})
