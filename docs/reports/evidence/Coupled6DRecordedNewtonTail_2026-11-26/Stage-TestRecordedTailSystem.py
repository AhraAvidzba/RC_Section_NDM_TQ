"""ДЛЯ ТЕСТОВ. Готовит точный большой отказ54/55 с его принятым префиксом25 для отдельной проверки без нового холодного поиска."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,struct
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
source=base/'Cold923RecordedNewtonCold96_20261126';parent=base/'Cold922ColdNewtonCapture_20261126'
root=base/'Cold925RecordedTailSystem_20261126'
proof=json.loads((source/'ColdRejectedContextsProof.json').read_bytes())
record=next(x for x in proof['records'] if x['serial']==54)
assert record['governing_row']==55 and record['arc_load_projection']==0
assert len(record['accepted_history_prefix_matches'])==1 and record['accepted_history_prefix_matches'][0]['state']==25
read=run_path(str(here/'Read-TestColdRawCapture.py'));frames,_=read['TestHistory'](source/'Native/Stock.RawHistory.bin');frame=frames[-1]
assert len(frame['fields'])==25
prefix=sha256(frame['fields'].tobytes()+frame['fractions'].tobytes()+frame['lambdas'].tobytes()).hexdigest().upper()
assert prefix==record['accepted_history_prefix_matches'][0]['raw_prefix_sha256']
def TestEnds(raw):
    """Находит границы исходных двоичных векторов, не преобразуя Double в десятичную запись."""
    positions=[0]
    for _ in range(67):
        n,=struct.unpack_from('<i',raw,positions[-1]);positions.append(positions[-1]+4+8*n)
    return positions
a=(source/'Native/NewtonRejected.00054.bin').read_bytes();b=(source/'Native/NewtonRejected.00055.bin').read_bytes()
pa,pb=TestEnds(a),TestEnds(b)
for i in (1,2,3,4):assert a[pa[i]:pa[i+1]]==b[pb[i]:pb[i+1]]
fixture=a[:pa[-1]]+b[pb[6]:pb[-1]]
binary=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](source/'Input/Reverse1.bin')
trajectory=run_path(str(here/'Stage-TestColdRawReplay.py'))['TestTrajectoryBytes'](binary,frame)
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'TailSystemCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
entry=(here/'TestReadRecordedTailSystem.bas.txt').read_text('utf-8-sig').rstrip()
prefixComment="' FOR TESTS: baseline entry point only; existing methods retained.\n"
anchor=prefixComment*8+'\n\nPublic Sub TestNewtonRejectedCapture('
assert code.count(anchor)==1
values[name]=code.replace(anchor,prefixComment*9+'\n\n'+entry+'\n\n\nPublic Sub TestNewtonRejectedCapture(',1).replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
values['Compile-TestExternalPoint.ps1']=compiler.replace("-ModuleNames 'modTestCoupledElasticState'","-ModuleNames 'modTestCoupledElasticState' -EntryPointNames 'TestReadRecordedTailSystem'",1).replace('\n','\r\n').encode('utf-8-sig')
name='tools/build_workbook/Run-Audit02ModeTests.ps1';runner=(base/'Cold857SpatialDeclineRecovery_20261125'/name).read_text('utf-8-sig')
anchor='        else { $result = [string]$excel.Run($macroName) }';assert runner.count(anchor)==1
values[name]=runner.replace(anchor,anchor+'''
        if ($macro -ne 'modTestCoupledElasticState.RunTestReversePolish') {throw 'Unexpected recorded tail setup macro.'}
        $lines.Add([string]$excel.Run("'RC_Section_NDM.xlsm'!modTestCoupledElasticState.TestReadRecordedTailSystem"))
        Save-Progress
''',1).replace('\n','\r\n').encode('utf-8-sig')
values.update({'Input/FrozenNewton.bin':fixture,'Input/Reverse1.bin':trajectory,
 'Input/SourceColdRawHistory.bin':(source/'Native/Stock.RawHistory.bin').read_bytes(),
 'SourceRejectedContextsProof.json':(source/'ColdRejectedContextsProof.json').read_bytes(),
 'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
 'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for name,data in values.items():
    if name.endswith('.bas'):assert max(map(len,data.decode('utf-8-sig').splitlines()))<1024
    (root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='ExactRecordedTailSystemAndRecoveryModes',actual_import_this_pass=['modTestCoupledElasticState'],
 actual_component_count=257,all_numerical_classes_exact922=True,source_record_pair=[54,55],source_history_prefix_states=25,
 native_base_residual_and_refined_matrix_guard_before_new_attempts=True,
 compare_reference_lambda_modes=True,comparison_changes_entire_recovery_policy_not_only_lambda_guard=True,
 material_history_not_accepted=True,original_force_tolerance_and30_iterations=True,
 no_cold_or_physical_capacity_claim=True,no_release_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'EXACT54/55 CURRENT/HISTORY/RHS;REFINED MATRIX GUARD;TWO REFERENCE MODES;SAME30')
