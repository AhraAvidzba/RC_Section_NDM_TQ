"""ДЛЯ ТЕСТОВ. Ограничивает специальный диагностический вызов его исходным macro; сохраняет сбой маршрутизации и повторяет все регрессии."""
from pathlib import Path
from hashlib import sha256
import json

root=Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold960ResolvedFieldProgress_20261126')
closure=json.loads((root/'Native/WorkingGuideRegressions.txt.ProcessClosure.json').read_bytes())
assert closure['ChildHasExited'] and closure['ChildExitCode']==1 and closure['Closure']['FinalState']['State']=='Absent'
name='tools/build_workbook/Run-Audit02ModeTests.ps1';path=root/name;code=path.read_text('utf-8-sig')
old='''        if ($macro -ne 'modTestCoupledElasticState.RunTestReversePolish') {throw 'Unexpected recorded tail setup macro.'}
        $lines.Add([string]$excel.Run("'RC_Section_NDM.xlsm'!modTestCoupledElasticState.TestRunFieldProgressSteps"))
        Save-Progress'''
new='''        if ($macro -eq 'modTestCoupledElasticState.RunTestReversePolish') {
            $lines.Add([string]$excel.Run("'RC_Section_NDM.xlsm'!modTestCoupledElasticState.TestRunFieldProgressSteps"))
            Save-Progress
        }'''
assert code.count(old)==1
out=root/'Native/RegressionDispatchSetupFailure';out.mkdir()
(out/'Run-Audit02ModeTests.BeforeCorrection.ps1').write_bytes(path.read_bytes())
for item in ('WorkingGuideRegressions.txt','WorkingGuideRegressions.txt.ProcessClosure.json'):(root/'Native'/item).rename(out/item)
path.write_bytes(code.replace(old,new,1).replace('\n','\r\n').encode('utf-8-sig'))
prep=json.loads((root/'Preparation.json').read_bytes());prep['inputs'][name]=sha256(path.read_bytes()).hexdigest().upper()
helper=Path(__file__);(root/helper.name).write_bytes(helper.read_bytes());prep['inputs'][helper.name]=sha256(helper.read_bytes()).hexdigest().upper()
prep['directed_scope']['regression_dispatch_setup_failure_preserved_and_corrected']=True
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
compiled=json.loads((root/'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
assert sha256((root/'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper()==compiled['WorkbookSHA256']
print('DISPATCH SCOPED TO ITS ORIGINAL SETUP MACRO;FAILED LOGS PRESERVED;SAVED BOOK UNCHANGED;ALL3 REGRESSIONS REQUIRE RERUN')
