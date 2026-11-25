"""ДЛЯ ТЕСТОВ. Проверяет обязательную H-пару на неизменной книге876; при необходимости тот же комплект допускает отдельный холодный запуск."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ');parent=base/'Cold876VerifiedHCornerPairs20_20261126';root=base/'Cold877VerifiedHRegressions_20261126'
audit=json.loads((parent/'PointTraceSourceComposition.json').read_bytes());assert audit['components']==256 and len(audit['imported_this_pass'])==2
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
prep=json.loads((root/'Preparation.json').read_bytes())
for name,data in {'tools/build_workbook/Run-Audit02ModeTests.ps1':(base/'Cold857SpatialDeclineRecovery_20261125/tools/build_workbook/Run-Audit02ModeTests.ps1').read_bytes(),
 'Run-TestOperationalArcRegressions.ps1':(here/'Run-TestOperationalArcRegressions.ps1').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()}.items():
    (root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope'].update(case='VerifiedHDirectedRegressions',actual_import_this_pass=audit['imported_this_pass'],actual_imports_this_replay=[],exact_compiled_source876=True,
    full_acceptance_suite=False,no_cold_or_capacity_claim=True)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
assert (root/'docs/regression/Diagnostic.xlsm').read_bytes()==(parent/'docs/regression/Diagnostic.xlsm').read_bytes()
print(root,'EXACT COMPILED876;164DIRECTED;NO FULL ACCEPTANCE CLAIM')
