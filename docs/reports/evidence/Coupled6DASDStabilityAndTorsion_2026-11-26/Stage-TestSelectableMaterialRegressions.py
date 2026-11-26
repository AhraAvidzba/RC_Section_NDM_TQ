"""ДЛЯ ТЕСТОВ. Готовит две короткие нелинейные проверки CDP после выделения общей калибровки, без пересборки и полного набора."""
from pathlib import Path
from hashlib import sha256
import json

root=Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold1018SelectableMaterialCore_20261126')
compiled=json.loads((root/'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
assert sha256((root/'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper()==compiled['WorkbookSHA256']
source=(root/'Run-TestSelectableMaterialCore.ps1').read_text('utf-8-sig')
source=source.replace("$taskMacros='modTestSelectableMaterialCore.TestRunSelectableMaterialCore'",
    "$taskMacros=@('modTestCoupledElasticState.RunTestCoupledEnvelope','modTestCoupledElasticState.RunTestCoupledConcreteResponse')")
source=source.replace('Native/RealColdReplay.txt','Native/MaterialRegressions.txt').replace('-MacroArgument1 $Root ','')
assert 'TestRunSelectableMaterialCore' not in source and '-MacroArgument1' not in source
path=root/'Run-TestSelectableMaterialRegressions.ps1';assert not path.exists()
path.write_bytes(source.replace('\n','\r\n').encode('utf-8-sig'))
print('PREPARED TWO SHORT EXISTING MATERIAL REGRESSIONS;SAME ACTUAL WORKING BOOK')
