"""ДЛЯ ТЕСТОВ. Проверяет тот же луч изгиба при трёх масштабах исходного момента с побайтно одинаковым ядром и последовательными Excel-процессами."""
from pathlib import Path
from hashlib import sha256
import json,shutil,struct

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ');parent=base/'Cold1029ASDTwoPointStart_20261126'
compiled=json.loads((parent/'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
book=(parent/'docs/regression/Diagnostic.xlsm').read_bytes();assert sha256(book).hexdigest().upper()==compiled['WorkbookSHA256']
verified=json.loads((parent/'Native/OpenSeesMainCapacityEquilibrium.json').read_bytes());assert verified['all56plus4_equations_and_six_resultants_verified']
proof=json.loads((parent/'EarlySourceProof.json').read_bytes());assert proof['one_updated_controller_exact']
cases=[(1030,'Cold1030ASDScale1M_20261126',1000000.),(1031,'Cold1031ASDScale4M_20261126',4000000.),(1032,'Cold1032ASDScale10M_20261126',10000000.)]
for n,folder,moment in cases:
    root=base/folder;assert not root.exists();root.mkdir()
    for directory in ('src','tests','tools','workbook','Input'):shutil.copytree(parent/directory,root/directory)
    for directory in ('Native','Temp','docs/regression'):(root/directory).mkdir(parents=True,exist_ok=True)
    for name in ('Guard-IsolatedOffice.ps1','Run-TestASDBendingMainCapacity.ps1'):shutil.copyfile(parent/name,root/name)
    (root/'docs/regression/Diagnostic.xlsm').write_bytes(book)
    shutil.copyfile(parent/'docs/regression/CodeUpdate.json',root/'docs/regression/CodeUpdate.json')
    (root/'Input/CapacityBase.bin').write_bytes(struct.pack('<id',1,moment))
    inherited=dict(proof,book_reused_bitexact_from_parent=str(parent),all256_compiled_components_reused_exact=True,no_new_compilation_required=True)
    (root/'EarlySourceProof.json').write_text(json.dumps(inherited,sort_keys=True,indent=2)+'\n',encoding='utf-8')
    (root/'CompiledSourceReuse.json').write_text(json.dumps(dict(parent=str(parent)),indent=2)+'\n',encoding='utf-8')
    shutil.copyfile(Path(__file__),root/Path(__file__).name)
    prep=dict(directed_scope=dict(case='ASD_same_bending_ray_scaled_base',base_Mx_Nmm=moment,
        exact_parent_book_reused=True,all_materials_and_limits_and_equilibrium_tolerances_unchanged=True),
        inputs={p.relative_to(root).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in root.rglob('*')
            if p.is_file() and p.relative_to(root).as_posix() not in ('docs/regression/Diagnostic.xlsm','docs/regression/CodeUpdate.json')})
    (root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
runner=base/cases[0][1]/'Run-TestASDCapacityScaleBundle.ps1'
roots='\n'.join("    '"+str(base/folder).replace('\\','/')+"'" for _,folder,_ in cases)
body='''# ДЛЯ ТЕСТОВ. Последовательно проверяет три масштаба одного луча; каждый дочерний запуск самостоятельно проверяет ресурсы и защищает чужие Excel.
$ErrorActionPreference='Stop'
$roots=@(
ROOTS
)
foreach($taskRoot in $roots) {
    & (Join-Path $taskRoot 'Run-TestASDBendingMainCapacity.ps1') -Root $taskRoot
    if($LASTEXITCODE -ne 0){throw ('Case failed before completion: '+$taskRoot)}
}
'''.replace('ROOTS',roots)
runner.write_bytes(body.replace('\n','\r\n').encode('utf-8-sig'))
print('PREPARED1030-1032;BASE MOMENTS1/4/10KNM;SAME COMPILED KERNEL;SEQUENTIAL GUARDED RUNS')
