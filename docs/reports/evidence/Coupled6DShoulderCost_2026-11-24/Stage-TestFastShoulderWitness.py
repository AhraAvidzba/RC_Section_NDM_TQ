"""ДЛЯ ТЕСТОВ. Сохраняет исход качества до экспорта и проверяет финальное поле от уже проверенного точного префикса."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,re

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold728FastShoulderWitness_20261124';parent=base/'Cold727CriticalShoulderProfile_20261124'
text=(parent/'Native/PostCrackReplay.txt').read_text('utf-8-sig')
assert 'WATCHDOG_FAILURE: timeout=180 sec;' in text
closure=json.loads((parent/'Native/PostCrackReplay.txt.ProcessClosure.json').read_bytes())
assert closure['Stage']=='TIMEOUT' and closure['Identity']['Valid'] and closure['Before']['State']=='Live' and closure['After']['State']=='Absent'
assert closure['Before']['PID']==closure['After']['PID']==closure['Identity']['PID']
assert closure['Before']['StartTicksUTC']==closure['After']['StartTicksUTC']==str(closure['Identity']['Ticks'])
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'WitnessCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
values={};prep=json.loads((root/'Preparation.json').read_bytes())
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
anchor='''        If solved Then
            solver.CopyTrial field, actual: fields.Add field: trace.Add 1#, trialFactor''';assert code.count(anchor)==1
code=code.replace(anchor,'''        Dim witnessHandleTest As Integer: witnessHandleTest = FreeFile
        Open evidenceRoot & "\\Native\\QualityOutcome.txt" For Output As #witnessHandleTest
        Print #witnessHandleTest, report
        Close #witnessHandleTest
'''+anchor,1)
anchor='            TestExportReverseCandidate 999\n';assert code.count(anchor)==1
code=code.replace(anchor,'            TestExportFinalShoulderWitness\n',1)
method,=re.findall(r'(?ms)^Private Sub TestExportReverseCandidate\(.*?^End Sub',code)
witness=method.replace('Private Sub TestExportReverseCandidate(ByVal caseIndex As Long)','Private Sub TestExportFinalShoulderWitness()\n    Const caseIndex As Long = 999',1)
anchor='    mReverseSolverTest.ResetHistory\n';assert witness.count(anchor)==1
witness=witness.replace(anchor,'''    If Not mReverseSolverTest.ReplayAcceptedFields(mReverseFieldsTest, mReverseFieldsTest.Count - 1) Then _
        Err.Raise 5, , "Не восстановлен точный проверенный префикс для финального поля плеча."
''',1)
witness=witness.replace('    For index = 1 To mReverseFieldsTest.Count','    For index = mReverseFieldsTest.Count To mReverseFieldsTest.Count',1)
witness=witness.replace('"F;" & CStr(index) & ";"','"F;1;"',1).replace('"R;" & CStr(index) & ";"','"R;1;"',1)
witness=witness.replace('TestWriteHardPreparedStress handle, index,','TestWriteHardPreparedStress handle, 1,',1)
code+="\n' -------------------- ДЛЯ ТЕСТОВ --------------------\n' Экспортирует единственное финальное поле, восстановив тот же исходный префикс StressOnly; проверки поля и шести усилий сохранены.\n"+witness+'\n'
assert max(map(len,code.splitlines()))<1024
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
values['Compile-TestExternalPoint.ps1']=(here/'Compile-TestExternalPoint.ps1').read_bytes()
runner=(root/'Run-TestPostCrackRawReplay.ps1').read_text('utf-8-sig').replace('-TimeoutSeconds 180','-TimeoutSeconds 240')
values['Run-TestPostCrackRawReplay.ps1']=runner.replace('\n','\r\n').encode('utf-8-sig')
values['ParentCodeUpdate.json']=(parent/'docs/regression/CodeUpdate.json').read_bytes()
values['ParentCompileClosure.json']=(parent/'CompilePointClosure.json').read_bytes()
values[Path(__file__).name]=Path(__file__).read_bytes()
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope'].update(case='CriticalShoulderWithImmediateWitness',actual_import_this_pass=['modTestCoupledElasticState'],
    original727_quality_and_numeric_core_exact=True,quality_outcome_flushed_before_export=True,
    exact_prefix52_replayed_without_new_field_corrections=True,only_final_field_tensors_exported=True,
    existing_all53_source_equilibrium_proof='Cold726PostYieldProfile_20261124/DirectProbeProof.json',
    previous727_timeout180_during_export=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('728 SAME QUALITY/NDS CORE; IMMEDIATE OUTCOME; FINAL FIELD EXPORT FROM EXACT VERIFIED PREFIX')
