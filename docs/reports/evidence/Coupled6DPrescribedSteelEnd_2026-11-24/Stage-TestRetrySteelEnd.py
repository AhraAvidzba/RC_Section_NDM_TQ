"""ДЛЯ ТЕСТОВ. Исправляет только отсутствовавший каталог отчёта724; исходный ошибочный запуск сохраняется."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,re

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold725SteelEndDirectoryRetry_20261124';parent=base/'Cold724PriorBranchSteelEnd_20261124'
text=(parent/'Native/CurveMatrix.txt').read_text('utf-8-sig')
closure=json.loads((parent/'Native/CurveMatrix.txt.ProcessClosure.json').read_bytes())
assert 'RUNTIME ERROR: VBAProject Path not found' in text and 'SOURCE_UNCHANGED: True' in text
assert not re.findall(r'(?m)^TOTAL',text)
assert closure['ChildHasExited'] and closure['Passed'] and closure['ChildExitCode']==1
identity=closure['Closure']['Identity'];final=closure['Closure']['FinalState']
assert identity['Valid'] and final['State']=='Absent' and final['PID']==identity['PID']
assert str(identity['Ticks'])==final['StartTicksUTC']
native=dict(passed=0,failed=1,macros=0,runtime_error_before_calculation=True,own_process_closed=True)
assert not (parent/'Matrix/Native/PostCrackTension.Curve.csv').exists()
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
(root/'Matrix/Native').mkdir(parents=True,exist_ok=True)
for name in ('tests/modTestCoupledElasticState.bas','src/Materials/CCoupledConcreteResponse.cls','docs/regression/Diagnostic.xlsm'):
    assert (root/name).read_bytes()==(parent/name).read_bytes()
prep=json.loads((root/'Preparation.json').read_bytes())
name=Path(__file__).name;(root/name).write_bytes(Path(__file__).read_bytes())
prep['inputs'][name]=sha256((root/name).read_bytes()).hexdigest().upper()
prep['directed_scope'].update(case='SteelEndReportDirectoryRetry',previous_failure_before_calculation=True,
    original724_source_and_compiled_book_exact=True,no_new_compilation=True,only_report_directory_created=True)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
(root/'PreparationFailure724.json').write_text(json.dumps(dict(native=native,
    reason='Missing Matrix/Native directory; no curve point or solver call completed',
    no_material_or_solver_failure_claimed=True),indent=2)+'\n',encoding='utf-8')
print('725 SAME COMPILED724 BOOK; ONLY MATRIX REPORT DIRECTORY CREATED; ORIGINAL FAILURE PRESERVED')
