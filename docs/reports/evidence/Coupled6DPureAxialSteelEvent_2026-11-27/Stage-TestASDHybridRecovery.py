"""ДЛЯ ТЕСТОВ. Проверяет существующий разностный гибридный резерв на ASD без ослабления материальной или силовой приёмки."""
from pathlib import Path
from hashlib import sha256
import json,shutil

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold1030ASDScale1M_20261126';root=base/'Cold1033ASDHybridRecovery_20261126';assert not root.exists()
compiled=json.loads((parent/'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
assert sha256((parent/'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper()==compiled['WorkbookSHA256']
native=(parent/'Native/RealColdReplay.txt').read_text('utf-8-sig');assert 'WATCHDOG_COMPLETED: exit=0; source unchanged=True' in native
root.mkdir()
for directory in ('src','tests','tools','workbook','Input'):shutil.copytree(parent/directory,root/directory)
for directory in ('Native','Temp','docs/regression'):(root/directory).mkdir(parents=True,exist_ok=True)
for name in ('Guard-IsolatedOffice.ps1','Run-TestASDBendingMainCapacity.ps1'):shutil.copyfile(parent/name,root/name)
shutil.copyfile(parent/'docs/regression/Diagnostic.xlsm',root/'docs/regression/Diagnostic.xlsm')
shutil.copyfile(parent/'docs/regression/CodeUpdate.json',root/'ParentCodeUpdate.json')
name='src/Solver/CCoupledNewtonSolver.cls';code=(root/name).read_text('utf-8-sig')
needle='''    SupportsHybridRecovery = Not (HybridResponseHasPrincipalBoundaries(mResponse) Or _
        HybridResponseHasPrincipalBoundaries(mPlusResponse) Or HybridResponseHasPrincipalBoundaries(mMinusResponse))'''
assert code.count(needle)==1
code=code.replace(needle,'    SupportsHybridRecovery = True',1)
old_comment="' Разрешает контроллеру запросить полный разностный резерв только для пространственной истории без спектральных границ."
assert code.count(old_comment)==1
code=code.replace(old_comment,"' Разрешает полный разностный резерв пространственной истории; материальные разрывы остаются в проверяемой системе.",1)
(root/name).write_bytes(code.replace('\n','\r\n').encode('utf-8-sig'))
compiler=(base/'Cold1029ASDTwoPointStart_20261126/Compile-TestASDTwoPointStart.ps1').read_text('utf-8-sig')
compiler=compiler.replace("-ModuleNames 'CCoupledLoadContinuation'","-ModuleNames 'CCoupledNewtonSolver'")
(root/'Compile-TestASDHybridRecovery.ps1').write_bytes(compiler.replace('\n','\r\n').encode('utf-8-sig'))
(root/'CompiledSourceReuse.json').write_text(json.dumps(dict(parent=str(parent)),indent=2)+'\n',encoding='utf-8')
shutil.copyfile(Path(__file__),root/Path(__file__).name)
prep=dict(directed_scope=dict(case='ASD_existing_hybrid_reserve_enabled',base_Mx_Nmm=1000000.,
    source_change='SupportsHybridRecovery gate and its comment only',existing_hybrid_code_reused=True,
    physical_materials_and_equilibrium_tolerances_unchanged=True,control_probe_budget=96,physical_capacity_not_assumed=True),
    inputs={p.relative_to(root).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in root.rglob('*')
        if p.is_file() and p.relative_to(root).as_posix()!='docs/regression/Diagnostic.xlsm'})
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('PREPARED1033; EXISTING HYBRID FD RESERVE ENABLED FOR ASD;SAME MATERIALS AND EQUILIBRIUM CRITERIA')
