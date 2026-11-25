"""ДЛЯ ТЕСТОВ. Даёт полной H-пробе исходный бюджет30 и отличает непробованную половину от неудачного решения до компиляции."""
from pathlib import Path
from hashlib import sha256
import json
root=Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold874SharedCurveHControl20_20261126');assert not (root/'CompilePoint.log').exists()
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
old='            Dim portion As Double: portion = 1#: If role = 1 Then portion = 0.5';new='            Dim portion As Double: portion = 1#: If role = 2 Then portion = 0.5';assert code.count(old)==1;code=code.replace(old,new,1)
start=code.index('            budget = 30 - used: If role = 1 Then budget = budget \\ 2')
end=code.index('\n            TestWriteReverseBinaryVector handle, mReverseSolverTest.ReadPolishEffortTest',start)
new='''            budget = 30 - used: solved = False: probes = 0: returned = 0#
            Dim iterations As Long: iterations = 0
            If budget > 0 Then
                solved = mReverseSolverTest.TryEnvelopeStep(mReverseBaseTest, mReverseDirectionTest, previous, factor, seed, _
                    factor + portion * ratio * (factor - previousFactor), target, returned, projection, absTol, relTol, budget, probes, False)
                iterations = mReverseSolverTest.IterationCount
            End If
            used = used + iterations: energy = 0#
            TestWriteReverseBinaryVector handle, Array(caseIndex, role, ratio, portion, target, Abs(CLng(solved)), returned, probes, iterations, used, revision, mReverseSolverTest.AcceptedRevision, budget)'''
code=code[:start]+new+code[end:]
prep=json.loads((root/'Preparation.json').read_bytes());data=code.replace('\n','\r\n').encode('utf-8-sig');(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
for f in (Path(__file__),Path(__file__).with_name('Stage-TestSharedCurveHControl.py')):
    (root/f.name).write_bytes(f.read_bytes());prep['inputs'][f.name]=sha256(f.read_bytes()).hexdigest().upper()
prep['directed_scope'].update(full_H_probe_gets_original30_budget=True,half_uses_remaining_pair_budget_only=True,skipped_half_recorded_with_zero_budget=True)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('FULL H GETS ORIGINAL30;PAIR MAX30;ZERO-REMAINING HALF EXPLICITLY NOT ATTEMPTED')
