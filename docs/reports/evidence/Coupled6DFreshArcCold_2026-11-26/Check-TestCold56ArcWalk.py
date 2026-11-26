"""ДЛЯ ТЕСТОВ. Независимо проверяет каждую принятую дуговую пару продолжения после исходного холодного состояния56."""
from pathlib import Path
from runpy import run_path
import json
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ');root=base/'Cold894Cold56ArcWalk_20261126';parent=base/'Cold888StrictEnvelopeGrowth_20261126'
assert (root/'docs/regression/Diagnostic.xlsm').read_bytes()==(parent/'docs/regression/Diagnostic.xlsm').read_bytes()
for path in (parent/'src').rglob('*'):
    if path.is_file():assert path.read_bytes()==(root/path.relative_to(parent)).read_bytes()
assert (root/'tests/modTestCoupledElasticState.bas').read_bytes()==(parent/'tests/modTestCoupledElasticState.bas').read_bytes()
scope=json.loads((root/'Preparation.json').read_bytes())['directed_scope'];assert scope['actual_imports_this_replay']==[]
file=here/'Check-TestFixedMetricKernel.py';code=file.read_text('utf-8-sig')
for old,new in [
 ("default=base/'Cold867FixedMetricKernel_20261126'","default=base/'Cold894Cold56ArcWalk_20261126'"),
 ("audit['imported_this_pass']==['CCoupledArcContinuation','CCoupledDeformationContinuation','CCoupledNewtonSolver','modTestCoupledElasticState']","audit['imported_this_pass']==['CCoupledDeformationContinuation']"),
 ("len(audit['other_actual_components_exact_parent_except_identifier_case'])==252","len(audit['other_actual_components_exact_parent_except_identifier_case'])==255"),
 ("scope['case']=='GeneralFixedMetricKernelFromCold19'","scope['case']=='ExactCold56FixedMetricArcWalk'"),
 ("len(source['fields'])==19","len(source['fields'])==56"),
 ("(base/'Cold866Cold19ReferenceMetricPairs_20261126/Input/Reverse1.bin')","(base/'Cold893Arc56Diagnostics_20261126/Input/Reverse1.bin')"),
 ('working_kernel_changed_and_verified=True','working_kernel_changed_this_replay=False'),
 ('source19_true_history=True','source56_true_cold_prefix_then_separate_follow_on_budget=True'),
 ("root/'FixedMetricKernelProof.json'","root/'Cold56ArcWalkProof.json'")]:
    assert old in code,old;code=code.replace(old,new)
import sys
old='    return float(np.sqrt(max(0.,delta@M@delta)))'
new="    strain=np.einsum('spcd,d->spc',op,delta)\n    return float(np.sqrt(np.einsum('spc,spc,p,c->',strain,strain,coords[:,2],np.array([1,1,1,.5,.5,.5]))/(3*sum(coords[:,2]))))"
assert code.count(old)==1;code=code.replace(old,new,1)
old='        expected_arc=float(np.sqrt(TestLength(movement)**2+alpha**2*tm[1]**2))'
new="        rounded_motion=(reference+movement)-reference\n        projected_spacing=float((M@movement)@rounded_motion/TestLength(movement))\n        expected_arc=float(np.sqrt(projected_spacing**2+alpha**2*tm[1]**2))"
assert code.count(old)==1;code=code.replace(old,new,1)
anchor='        assert abs(info[12]-alpha**2*tm[1]/expected_arc)<1e-18'
assert code.count(anchor)==1
code=code.replace(anchor,"        if abs(info[12]-alpha**2*tm[1]/expected_arc)>=1e-18: print('Q_ARITHMETIC_DIAG',attempt,info[12],alpha**2*tm[1]/expected_arc,alpha,info[9],tm.tolist(),expected_arc)\n"+anchor,1)
saved=list(sys.argv);sys.argv=[str(file)]
try:exec(compile(code,str(file),'exec'),{'__file__':str(file),'__name__':'cold56_arc_walk_evidence_test'})
finally:sys.argv=saved
