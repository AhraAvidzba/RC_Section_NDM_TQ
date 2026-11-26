"""ДЛЯ ТЕСТОВ. Исправляет выделение выходных массивов тестового драйвера; сохраняет отрицательный запуск 1015 без изменения его входов."""
from pathlib import Path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ');failed=base/'Cold1015PureBendingSteelLimit_20261126'
report=(failed/'Native/RealColdReplay.txt').read_text('utf-8-sig');assert 'RUNTIME ERROR: VBAProject Subscript out of range' in report and 'SOURCE_UNCHANGED: True' in report
closure=json.loads((failed/'Native/RealColdReplay.txt.ProcessClosure.json').read_text('utf-8-sig'))
assert closure['ChildHasExited'] and closure['ChildExitCode']==1
assert closure['Closure']['Passed'] and closure['Closure']['Identity']['Valid'] and closure['Closure']['FinalState']['State']=='Absent'
owner=json.loads((failed/'Native/RealColdReplay.txt.Ownership.json').read_text('utf-8-sig'))
assert closure['Closure']['Identity']['PID']==owner['PID'] and int(closure['Closure']['Identity']['Ticks'])==int(owner['StartTicksUTC'])
proof=dict(test_driver_failed_before_steel_control=True,reason='PointStrain and AxialGradientAtPoint require six allocated output slots.',
    numerical_material_classes_unchanged=True,source_unchanged=True,own_exact_excel_closed=True,
    no_equilibrium_or_physical_failure_assigned=True,whole_run_passed=False,
    source_book_sha256=sha256((failed/'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper(),corrected_repeat='1016')
(failed/'SteelLimitDriverFailureProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
source=here/'Stage-TestPureBendingSteelLimit.py';code=source.read_text('utf-8-sig')
assert code.count('Dim strain() As Double, gradient() As Double')==1
code=code.replace('Dim strain() As Double, gradient() As Double','Dim strain(1 To 6) As Double, gradient(1 To 6) As Double')
code=code.replace('Cold1015PureBendingSteelLimit_20261126','Cold1016PureBendingSteelLimitFixed_20261126').replace('PREPARED1015','PREPARED1016')
exec(compile(code,str(source),'exec'),{'__file__':str(Path(__file__).resolve()),'__name__':'fixed_steel_limit_stage_test'})
