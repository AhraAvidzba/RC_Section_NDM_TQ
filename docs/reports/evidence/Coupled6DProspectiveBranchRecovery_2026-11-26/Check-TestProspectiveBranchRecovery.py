"""ДЛЯ ТЕСТОВ. Независимо проверяет корни рабочего восстановления, исходные дуговые случаи и точный состав изменений."""
from pathlib import Path
import json
here=Path(__file__).resolve().parent;file=here/'Check-TestArc56Diagnostics.py';code=file.read_text('utf-8-sig')
for old,new in [('Cold893Arc56Diagnostics_20261126','Cold915ProspectiveBranchRecovery_20261126'),
 ('Cold888StrictEnvelopeGrowth_20261126','Cold911NewtonRejectedCapture_20261126'),
 ("audit['imported_this_pass']==['modTestCoupledElasticState']","audit['imported_this_pass']==['CCoupledNewtonSolver']"),
 ('len(source[\'fields\'])==56','len(source[\'fields\'])==32'),('len(exported)==57','len(exported)==33'),
 ('exported[:56]','exported[:32]'),('TestH(phase,57)','TestH(phase,33)'),('source_states=56','source_states=32'),
 ('Arc56','Arc32'),('ARC56','ARC32'),('all_working_sources_exact888=True','only_prospective_branch_recovery_added=True')]:
    assert old in code,old;code=code.replace(old,new)
old="    if path.is_file():assert path.read_bytes()==(root/path.relative_to(parent)).read_bytes()"
new="    if path.is_file() and path.name!='CCoupledNewtonSolver.cls':assert path.read_bytes()==(root/path.relative_to(parent)).read_bytes()"
assert code.count(old)==1;code=code.replace(old,new,1)
anchor='    effort=r.VectorTest();found=[]';assert code.count(anchor)==1
code=code.replace(anchor,"    comparison_header=r.VectorTest();assert comparison_header.tolist()==[0.]\n"+anchor,1)
code=code.replace("root/'Arc32DiagnosticsProof.json'","root/'ProspectiveBranchRecoveryProof.json'")
namespace={'__file__':str(file),'__name__':'prospective_branch_recovery_roots_test'}
exec(compile(code,str(file),'exec'),namespace)
root=namespace['root'];parent=namespace['parent']
stager=(here/'Stage-TestProspectiveBranchRecovery.py').read_text('utf-8-sig')
start=stager.index("name='src/Solver/CCoupledNewtonSolver.cls';code=");end=stager.index('values[name]=code',start)
expected={'root':parent,'here':here};exec(compile(stager[start:end],str(here/'Stage-TestProspectiveBranchRecovery.py'),'exec'),expected)
assert expected['code']==(root/'src/Solver/CCoupledNewtonSolver.cls').read_text('utf-8-sig')
before=json.loads((root.parent/'Cold909ResidualArcPredictorBaseline_20261126/ResidualArcPredictorBaselineProof.json').read_bytes())
path=root/'ProspectiveBranchRecoveryProof.json';after=json.loads(path.read_bytes())
lost=[a['case'] for a,b in zip(before['cases'],after['cases']) if a['solved'] and not b['solved']]
gained=[a['case'] for a,b in zip(before['cases'],after['cases']) if not a['solved'] and b['solved']]
text=(root/'Native/Arc32Diagnostics.txt').read_text('cp1251')
after.update(lost_original_cases=lost,gained_original_cases=gained,
    numerical_source_exactly_reconstructed_from_stager=True,
    actual_prospective_diagnostic_records=text.count('проб касательной за сменой ветви:'),
    original_budget_and_force_acceptance_unchanged=True,directed_nonregression_passed=not lost,
    physical_capacity_certified=False,cold_and_regression_verification_still_required=True)
path.write_text(json.dumps(after,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('BASELINE LOSSES',lost,'GAINS',gained,'ACTUAL DIAGNOSTIC RECORDS',after['actual_prospective_diagnostic_records'])
assert not lost,'Рабочее восстановление потеряло исходный принятый случай; вариант не принят.'
