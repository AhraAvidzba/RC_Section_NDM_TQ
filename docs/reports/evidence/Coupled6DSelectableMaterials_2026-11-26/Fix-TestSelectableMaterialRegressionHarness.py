"""ДЛЯ ТЕСТОВ. Исправляет тип аргумента короткого watchdog до запуска Excel: его Macro принимает строку с разделением запятой."""
from pathlib import Path
import json

root=Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold1018SelectableMaterialCore_20261126')
path=root/'Run-TestSelectableMaterialRegressions.ps1';text=path.read_text('utf-8-sig')
assert not (root/'Native/MaterialRegressions.txt').exists()
needle="$taskMacros=@('modTestCoupledElasticState.RunTestCoupledEnvelope','modTestCoupledElasticState.RunTestCoupledConcreteResponse')"
assert text.count(needle)==1
(root/'MaterialRegressionHarnessBeforeFix.ps1.txt').write_bytes(path.read_bytes())
text=text.replace(needle,"$taskMacros='modTestCoupledElasticState.RunTestCoupledEnvelope,modTestCoupledElasticState.RunTestCoupledConcreteResponse'")
path.write_bytes(text.replace('\n','\r\n').encode('utf-8-sig'))
(root/'MaterialRegressionHarnessFix.json').write_text(json.dumps(dict(error='Macro parameter expects string, not array',
    numerical_calculation_not_started_at_failure=True,source_workbook_unchanged=True),indent=2)+'\n',encoding='utf-8')
print('MATERIAL REGRESSION WATCHDOG ARGUMENT FIXED; NO EXCEL STARTED AT ORIGINAL PARAMETER FAILURE')
