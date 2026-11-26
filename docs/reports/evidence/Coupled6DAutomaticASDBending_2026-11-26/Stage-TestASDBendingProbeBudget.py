"""ДЛЯ ТЕСТОВ. Изолирует влияние бюджета продолжения: 192 пробы вместо 96, без изменения материалов, шагов и допусков."""
from pathlib import Path
from hashlib import sha256
import json,shutil

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold1021ASDBendingMainCapacityRetry_20261126';root=base/'Cold1022ASDBendingProbeBudget_20261126'
assert not root.exists()
compiled=json.loads((parent/'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
assert sha256((parent/'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper()==compiled['WorkbookSHA256']
assert json.loads((parent/'MainCapacitySourceProof.json').read_bytes())['all255_other_components_unchanged_except_identifier_case']
native=(parent/'Native/RealColdReplay.txt').read_text('utf-8-sig')
assert 'WATCHDOG_COMPLETED: exit=0; source unchanged=True' in native and 'status=NumFail' in native
root.mkdir()
for directory in ('src','tests','tools','workbook'):shutil.copytree(parent/directory,root/directory)
for directory in ('Native','Temp','docs/regression'):(root/directory).mkdir(parents=True,exist_ok=True)
shutil.copyfile(parent/'docs/regression/Diagnostic.xlsm',root/'docs/regression/Diagnostic.xlsm')
shutil.copyfile(parent/'docs/regression/CodeUpdate.json',root/'ParentCodeUpdate.json')
shutil.copyfile(parent/'Guard-IsolatedOffice.ps1',root/'Guard-IsolatedOffice.ps1')
name='src/Solver/CCoupledDeformationContinuation.cls';body=(root/name).read_text('utf-8-sig')
needle='Private Const PROBE_LIMIT As Long = 96';assert body.count(needle)==1
(root/name).write_bytes(body.replace(needle,'Private Const PROBE_LIMIT As Long = 192').replace('\n','\r\n').encode('utf-8-sig'))
compile_script=(parent/'Compile-TestSelectableMaterialCore.ps1').read_text('utf-8-sig')
compile_script=compile_script.replace("-ModuleNames 'modTestSelectableMaterialCore'","-ModuleNames 'CCoupledDeformationContinuation'")
(root/'Compile-TestASDProbeBudget.ps1').write_bytes(compile_script.replace('\n','\r\n').encode('utf-8-sig'))
runner=(parent/'Run-TestASDBendingMainCapacity.ps1').read_text('utf-8-sig')
runner=runner.replace('$compiled.ProductionChanged -or',"!$compiled.ProductionChanged -or @($compiled.Modules).Count -ne 1 -or $compiled.Modules[0] -cne 'CCoupledDeformationContinuation' -or")
(root/'Run-TestASDBendingMainCapacity.ps1').write_bytes(runner.replace('\n','\r\n').encode('utf-8-sig'))
(root/'CompiledSourceReuse.json').write_text(json.dumps(dict(parent=str(parent)),indent=2)+'\n',encoding='utf-8')
shutil.copyfile(Path(__file__),root/Path(__file__).name)
prep=dict(directed_scope=dict(case='ASD_main_capacity_budget_only',private_controller_change='PROBE_LIMIT96to192',
    material_and_equilibrium_criteria_unchanged=True,reference_ASD_strain_limit_kNm=4.658444129476362),
    inputs={p.relative_to(root).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in root.rglob('*')
        if p.is_file() and p.relative_to(root).as_posix()!='docs/regression/Diagnostic.xlsm'})
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('PREPARED1022; ONLY CONTROLLER PROBE BUDGET CHANGES96to192')
