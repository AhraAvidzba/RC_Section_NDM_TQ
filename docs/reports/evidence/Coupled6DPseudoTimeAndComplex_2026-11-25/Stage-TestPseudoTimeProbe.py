"""ДЛЯ ТЕСТОВ. Проверяет собственный ограниченный численный резерв, переиспользуя штатные LU, метрический контроль и проверку НДС."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,re

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold739PseudoTimeProbe_20261125';parent=base/'Cold738TailFullDifferences_20261125'
proof=json.loads((parent/'FullDifferenceProof.json').read_bytes())
assert proof['sample_count']==247 and proof['assembled_vs_pure_residual_bitidentical']
assert all(p['central']['rank']==17 and not p['unavailable'] for p in proof['steps'])
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'PseudoTimeCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
values={};prep=json.loads((root/'Preparation.json').read_bytes())
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
anchor='Option Explicit\n';assert code.count(anchor)==1
code=code.replace(anchor,anchor+'''Private mPtcBestFieldTest() As Double, mPtcBestFactorTest As Double, mPtcTargetTest As Double
Private mPtcInitialNormTest As Double, mPtcBestNormTest As Double, mPtcBestReadyTest As Boolean
Private mPtcRootTest As String, mPtcSuppressTest As Boolean
''',1)
anchor='        If ready Then TestWriteTailDifferenceSamples equations, zero, residual\n';assert code.count(anchor)==1
code=code.replace(anchor,'        If ready Then TestProbePseudoTimeDirections equations, zero, residual, matrix, accepted, seed\n',1)
anchor='    If index = mReverseFieldsTest.Count Then TestWriteFrozenTailMatrices field, factor\n';assert code.count(anchor)==1
code=code.replace(anchor,'    If index = mReverseFieldsTest.Count And Not mPtcSuppressTest Then TestWriteFrozenTailMatrices field, factor\n',1)
method,=re.findall(r'(?ms)^Private Sub TestExportFastFrozenReplay\(.*?^End Sub',code)
anchor='    Close #handle\n    Exit Sub\n';assert method.count(anchor)==1
changed=method.replace(anchor,'    Close #handle\n    If Not mPtcSuppressTest Then TestVerifyPseudoTimeCandidate\n    Exit Sub\n',1)
code=code.replace(method,changed,1)+'\n'+(here/'CandidateTestPseudoTimeProbe.bas.txt').read_text('utf-8-sig')
assert max(map(len,code.splitlines()))<1024
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
values['CandidateTestPseudoTimeProbe.bas.txt']=(here/'CandidateTestPseudoTimeProbe.bas.txt').read_bytes()
values['ParentCodeUpdate.json']=(parent/'docs/regression/CodeUpdate.json').read_bytes()
values['ParentCompileClosure.json']=(parent/'CompilePointClosure.json').read_bytes()
values[Path(__file__).name]=Path(__file__).read_bytes()
for name,b in values.items():(root/name).write_bytes(b);prep['inputs'][name]=sha256(b).hexdigest().upper()
prep['directed_scope'].update(case='BoundedPseudoTimeDirectionsThenOrdinaryNDS',actual_import_this_pass=['modTestCoupledElasticState'],
    native_numeric_kernel_exact738=True,original70_history_exact=True,
    preconditioner='fresh native accepted-field Jacobian in same immutable context',
    trial_mu=[.1,1,10,100],trial_fractions=[1,.5,.25],maximum_readonly_trials=12,
    field_movement_radius='native SpatialFieldDistance from accepted field to forward prediction',
    all_linear_solves_use_existing_CCoupledLinearSystem=True,final_root_uses_original_tolerances_and30_iterations=True,
    only_successful_root_replayed_as71states=True,no_automatic_capacity_search=True,
    numerical_candidate_not_production=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('739 NATIVE BOUNDED PSEUDO-TIME PROPOSALS; SAME CORE/HISTORY; FINAL ORDINARY30-ITERATION NDS CHECK')
