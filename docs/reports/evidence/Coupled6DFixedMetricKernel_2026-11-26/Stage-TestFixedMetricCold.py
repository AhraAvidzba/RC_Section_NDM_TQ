"""ДЛЯ ТЕСТОВ. Использует неизменную скомпилированную книгу867 для регрессии и настоящего холодного96-запуска без передачи тёплых полей."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold867FixedMetricKernel_20261126';root=base/'Cold868FixedMetricCold_20261126'
audit=json.loads((parent/'PointTraceSourceComposition.json').read_bytes());assert audit['components']==256 and len(audit['imported_this_pass'])==4
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
prep=json.loads((root/'Preparation.json').read_bytes())
for name,data in {'tools/build_workbook/Run-Audit02ModeTests.ps1':(base/'Cold857SpatialDeclineRecovery_20261125/tools/build_workbook/Run-Audit02ModeTests.ps1').read_bytes(),
 'Run-TestOperationalArcRegressions.ps1':(here/'Run-TestOperationalArcRegressions.ps1').read_bytes(),
 Path(__file__).name:Path(__file__).read_bytes()}.items():
    (root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope'].update(case='FixedMetricKernelTrueCold96',same_compiled_working_code_as867=True,
    actual_import_this_pass=audit['imported_this_pass'],actual_imports_this_replay=[],
    full_cold_from_zero=True,warm_input_fields_not_used_by_stock_macro=True,
    primary_iterations_per_phase30_unchanged=True,global_probe_limit96_unchanged=True,
    full_suite=False,timeout_seconds600=True,saved_reopened_results_required=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
assert (root/'docs/regression/Diagnostic.xlsm').read_bytes()==(parent/'docs/regression/Diagnostic.xlsm').read_bytes()
print(root,'EXACT COMPILED867;SOURCE/HISTORY NOT REUSED;TRUE COLD96 AND TARGETED164 READY')
