"""ДЛЯ ТЕСТОВ. После собственной посттрещинной кривой пробует точно заданные моменты обычным TryStep без готовых полей извне."""
from pathlib import Path
from hashlib import sha256
import json,shutil

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ');parent=base/'Cold1011PureBendingASDCurvature_20261126'
root=base/'Cold1013PureBendingPostCurveForce_20261126';assert not root.exists()
compiled=json.loads((parent/'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
assert compiled['Compiled'] and sha256((parent/'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper()==compiled['WorkbookSHA256']
root.mkdir()
for directory in ('src','tests','tools','Input','workbook'):shutil.copytree(parent/directory,root/directory)
for directory in ('Native','Temp','docs/regression'):(root/directory).mkdir(parents=True,exist_ok=True)
for name in ('Guard-IsolatedOffice.ps1','Run-TestPureBendingCurvature.ps1'):shutil.copyfile(parent/name,root/name)
shutil.copyfile(parent/'docs/regression/Diagnostic.xlsm',root/'docs/regression/Diagnostic.xlsm')
shutil.copyfile(parent/'docs/regression/CodeUpdate.json',root/'ParentCodeUpdate.json')
shutil.copyfile(here/'Compile-TestExternalPoint.ps1',root/'Compile-TestExternalPoint.ps1')
source=(parent/'tests/modTestCoupledElasticState.bas').read_text('utf-8-sig')
anchor='    TestRunPureBendingCurvature = "CURVATURE_TRACE:';assert source.count(anchor)==1
extra='''    ' -------------------- ДЛЯ ТЕСТОВ --------------------
    ' Нагрузка вновь управляется силой; принятая история получена только что этим запуском из нулевого состояния.
    handle = FreeFile
    Open folder & "\\Native\\PostCurveForce.bin" For Binary Access Write As #handle
    magic = "NDMBENDPOST00001": Put #handle, , magic
    TestWriteReverseBinaryVector handle, Array(options(2), factor, 6#)
    Dim forceRatio As Variant, probeSolved As Boolean
    For Each forceRatio In Array(1#, 1.02, 1.05, 1.1, 1.2, 1.3)
        probeSolved = solver.TryStep(0#, CDbl(forceRatio) * CDbl(options(2)), 0#, 0#, 0#, 0#, 0.0000005, 0.0000005, 30)
        effort = solver.ReadPolishEffortTest
        TestWriteReverseBinaryVector handle, Array(CDbl(forceRatio), Abs(CLng(probeSolved)), effort(0), effort(1), effort(7))
        If probeSolved Then
            solver.CopyTrial nextField, actual
            TestWriteReverseBinaryVector handle, nextField: TestWriteReverseBinaryVector handle, actual
            TestWriteASDColdStressTest handle, nextField
            solver.AcceptStep
        Else
            solver.DiscardStep
        End If
    Next forceRatio
    Close #handle: handle = 0
'''
source=source.replace(anchor,extra+anchor,1)
(root/'tests/modTestCoupledElasticState.bas').write_bytes(source.replace('\n','\r\n').encode('utf-8-sig'))
shutil.copyfile(Path(__file__),root/Path(__file__).name)
(root/'CompiledSourceReuse.json').write_text(json.dumps(dict(parent=str(parent)),indent=2)+'\n',encoding='utf-8')
prep=json.loads((parent/'Preparation.json').read_bytes());prep['inputs']={p.relative_to(root).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in root.rglob('*')
    if p.is_file() and p.relative_to(root).as_posix()!='docs/regression/Diagnostic.xlsm'}
prep['directed_scope'].update(case='PureBendingPostCurveForce',actual_import_this_pass=['modTestCoupledElasticState'],
    all256_numerical_components_exact1011=True,post_curve_force_ratios=[1,1.02,1.05,1.1,1.2,1.3],
    force_probes_use_this_run_accepted_history=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('PREPARED1013; EXISTING FORCE SOLVER AFTER THIS RUN CURVATURE PATH; NUMERICAL CLASSES UNCHANGED')
