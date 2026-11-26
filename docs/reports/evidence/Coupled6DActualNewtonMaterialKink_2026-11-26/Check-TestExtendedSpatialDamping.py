"""ДЛЯ ТЕСТОВ. Независимо проверяет все найденные корни расширенного демпфирования и сохранность исходных восьми случаев."""
from pathlib import Path
import json
here=Path(__file__).resolve().parent;file=here/'Check-TestArc56Diagnostics.py';code=file.read_text('utf-8-sig')
for old,new in [('Cold893Arc56Diagnostics_20261126','Cold912ExtendedSpatialDamping_20261126'),
 ('Cold888StrictEnvelopeGrowth_20261126','Cold911NewtonRejectedCapture_20261126'),
 ("audit['imported_this_pass']==['modTestCoupledElasticState']","audit['imported_this_pass']==['CCoupledNewtonSolver']"),
 ('len(source[\'fields\'])==56','len(source[\'fields\'])==32'),('len(exported)==57','len(exported)==33'),
 ('exported[:56]','exported[:32]'),('TestH(phase,57)','TestH(phase,33)'),('source_states=56','source_states=32'),
 ('Arc56','Arc32'),('ARC56','ARC32'),('all_working_sources_exact888=True','only_spatial_damping_floor_changed=True')]:
    assert old in code,old;code=code.replace(old,new)
old="    if path.is_file():assert path.read_bytes()==(root/path.relative_to(parent)).read_bytes()"
new="    if path.is_file() and path.name!='CCoupledNewtonSolver.cls':assert path.read_bytes()==(root/path.relative_to(parent)).read_bytes()"
assert code.count(old)==1;code=code.replace(old,new,1)
anchor="source=run_path(str(here/'Check-TestActualReverseCapture.py'))"
insert='''rel='src/Solver/CCoupledNewtonSolver.cls'
old='        If mSteelBranchBuilds > 0 Then minimumFraction = STEEL_MIN_FRACTION'
new="        If mSteelBranchBuilds > 0 Or (mControlActive And mUseSpatialHistory) Then minimumFraction = STEEL_MIN_FRACTION ' Меньшие пробы нужны у близкой смены пространственной ветви; приёмка равновесия не меняется."
assert (parent/rel).read_text('utf-8-sig').count(old)==1
assert (root/rel).read_text('utf-8-sig')==(parent/rel).read_text('utf-8-sig').replace(old,new,1)
'''
assert code.count(anchor)==1;code=code.replace(anchor,insert+anchor,1)
anchor='    effort=r.VectorTest();found=[]'
assert code.count(anchor)==1
code=code.replace(anchor,"    comparison_header=r.VectorTest();assert comparison_header.tolist()==[0.]\n"+anchor,1)
code=code.replace("root/'Arc32DiagnosticsProof.json'","root/'ExtendedSpatialDampingProof.json'")
exec(compile(code,str(file),'exec'),{'__file__':str(file),'__name__':'extended_spatial_damping_roots_test'})
base=Path('D:/CodexScratch/RC_Section_NDM_TQ');root=base/'Cold912ExtendedSpatialDamping_20261126'
before=json.loads((base/'Cold909ResidualArcPredictorBaseline_20261126/ResidualArcPredictorBaselineProof.json').read_bytes())
path=root/'ExtendedSpatialDampingProof.json';after=json.loads(path.read_bytes())
lost=[a['case'] for a,b in zip(before['cases'],after['cases']) if a['solved'] and not b['solved']]
gained=[a['case'] for a,b in zip(before['cases'],after['cases']) if not a['solved'] and b['solved']]
after.update(lost_original_cases=lost,gained_original_cases=gained,original_budget_and_force_acceptance_unchanged=True,
    directed_candidate_acceptable=not lost,cold_and_regression_verification_still_required=True)
path.write_text(json.dumps(after,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('BASELINE LOSSES',lost,'GAINS',gained)
assert not lost,'Меньшие пробы потеряли исходный найденный корень; вариант не принят.'
