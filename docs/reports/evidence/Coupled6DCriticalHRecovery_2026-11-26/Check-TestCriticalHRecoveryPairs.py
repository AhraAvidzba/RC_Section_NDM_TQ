"""ДЛЯ ТЕСТОВ. Проверяет критическую H-пару878 и точное сохранение доказанного решения876 при изменении только области применения."""
from pathlib import Path
import json
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ');root=base/'Cold878CriticalHRecoveryPairs_20261126';parent=base/'Cold876VerifiedHCornerPairs20_20261126'
for name,old,new in (
 ('src/Solver/CCoupledDeformationContinuation.cls','                trial.Lambda, projection, mAbsTol, mRelTol, mMaxIterations, probes, True) Then','                trial.Lambda, projection, mAbsTol, mRelTol, mMaxIterations, probes, True, True) Then'),
 ('src/Solver/CCoupledArcContinuation.cls','        projection, mAbsolute, mRelative, mMaximumIterations \\ 2, probes, False)','        projection, mAbsolute, mRelative, mMaximumIterations \\ 2, probes, False, True)'),
 ('tests/modTestCoupledElasticState.bas','            factor + ratio * (factor - previousFactor), target, returned, projection, absTol, relTol, 30, probes, False)','            factor + ratio * (factor - previousFactor), target, returned, projection, absTol, relTol, 30, probes, False, True)')):
    before=(parent/name).read_text('utf-8-sig');after=(root/name).read_text('utf-8-sig');assert after.count(new)==1 and after.replace(new,old,1)==before,name
source_file=here/'Check-TestVerifiedHCornerPairs.py';text=source_file.read_text('utf-8-sig').replace('Cold876VerifiedHCornerPairs20_20261126','Cold878CriticalHRecoveryPairs_20261126')
text=text.replace("audit['imported_this_pass']==['CCoupledNewtonSolver','modTestCoupledElasticState']","audit['imported_this_pass']==['CCoupledArcContinuation','CCoupledDeformationContinuation','CCoupledNewtonSolver','modTestCoupledElasticState']")
text=text.replace("==254\")","==252\")")
anchor="namespace={'__file__':str(source_file),'__name__':'verified_H_pair_source_and_analytic_test'}"
assert text.count(anchor)==1
insert='''code=code.replace("f.name not in ('CCoupledArcContinuation.cls','CCoupledNewtonSolver.cls')", "f.name not in ('CCoupledArcContinuation.cls','CCoupledNewtonSolver.cls','CCoupledDeformationContinuation.cls')")
'''
text=text.replace(anchor,insert+anchor,1)
exec(compile(text,str(source_file),'exec'),{'__file__':str(source_file),'__name__':'critical_H_pair_scope_test'})
proof=json.loads((root/'VerifiedHCornerPairsProof.json').read_bytes());assert proof['verified_working_H_pairs']==1
for role in ('Full','Half'):
    for suffix in ('.csv','.Operators.csv'):
        name='Native/VerifiedH.Case1'+role+'.Sequence'+suffix
        assert (root/name).read_bytes()==(parent/name).read_bytes(),name
proof.update(H_pair_required_at_recovery_and_turn=True,ordinary_H_steps_preserve_original_scope=True,
    exactly_one_argument_changed_in_each_controller_arc_and_test_call=True,
    verified_case1_full_half_stresses_and_operators_bit_identical876=True)
(root/'CriticalHRecoveryPairsProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('CRITICAL H PAIR SAME TWO ROOTS/STRESSES/OPERATORS AS876;NORMAL H SCOPE RESTORED;NO PHYSICAL LIMIT')
