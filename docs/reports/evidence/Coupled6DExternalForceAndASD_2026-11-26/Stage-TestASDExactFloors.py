"""ДЛЯ ТЕСТОВ. Исправляет подготовку численного нижнего уровня таблиц по исходнику ASD; узлы диаграмм, решатель и силовые допуски неизменны."""
from pathlib import Path
from hashlib import sha256
import json,shutil

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold1001ASDNativeUpdate_20261126';root=base/'Cold1002ASDExactFloors_20261126';assert not root.exists()
cause=json.loads((parent/'Native/ASDNativeSourceComparison.json').read_bytes())
assert cause['prepared_native']['native_error_MPa']<1e-10 and cause['OpenSees_source']['OpenSees_error_MPa']<1e-10
root.mkdir()
for directory in ('src','tests','tools','Input','workbook'):shutil.copytree(parent/directory,root/directory)
for directory in ('Native','Temp','docs/regression'):(root/directory).mkdir(parents=True,exist_ok=True)
for name in ('Guard-IsolatedOffice.ps1','Run-TestASDHardShearCold.ps1','PointTraceSourceComposition.json','CompilePointClosure.json','CompileClosure.json'):
    shutil.copyfile(parent/name,root/name)
for name in ('Diagnostic.xlsm','CodeUpdate.json'):shutil.copyfile(parent/'docs/regression'/name,root/'docs/regression'/name)
contract=json.loads((parent/'Input/MaterialContract.json').read_bytes());floors={}
for kind,rows in contract['uniaxial_tables'].items():
    positive=[abs(b[1]-a[1]) for a,b in zip(rows,rows[1:]) if b[1]!=a[1]]
    floors[kind]=min(positive)*1e-6
assert floors==cause['OpenSees_source']['floors']
lines=(root/'Input/matched.csv').read_text('utf-8').splitlines()
for i,line in enumerate(lines):
    if line.startswith('BRANCH;'):
        fields=line.split(';');fields[3]=format(floors[fields[1]],'.17g');lines[i]=';'.join(fields)
(root/'Input/matched.csv').write_bytes(('\r\n'.join(lines)+'\r\n').encode())
contract.update(prepared_stress_floors_MPa=floors,
    stress_floor_rule='1e-6 times the minimum strictly positive absolute difference of adjacent nominal stress ordinates, as ASDConcrete3D HardeningLaw',
    all_curve_knots_strains_nominal_stresses_and_damage_unchanged=True)
(root/'Input/MaterialContract.json').write_text(json.dumps(contract,indent=2,sort_keys=True)+'\n',encoding='utf-8')
(root/'SourceFloorCauseComparison.json').write_bytes((parent/'Native/ASDNativeSourceComparison.json').read_bytes())
compiled=json.loads((parent/'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
assert sha256((root/'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper()==compiled['WorkbookSHA256']
reuse=dict(parent=str(parent),all257_actual_components_and_saved_workbook_bitexact=True,actual_import_this_pass=[],
           only_prepared_stress_floors_changed=True,source_workbook_sha256=compiled['WorkbookSHA256'])
(root/'CompiledSourceReuse.json').write_text(json.dumps(reuse,indent=2,sort_keys=True)+'\n',encoding='utf-8')
(root/'Replay-TestASDNativeMaterial.py').write_bytes((here/'Replay-TestASDNativeMaterial.py').read_bytes())
(root/Path(__file__).name).write_bytes(Path(__file__).read_bytes())
prep=json.loads((parent/'Preparation.json').read_bytes())
prep['inputs']={p.relative_to(root).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in root.rglob('*')
    if p.is_file() and p.relative_to(root).as_posix() not in ('docs/regression/Diagnostic.xlsm','docs/regression/CodeUpdate.json')}
prep['directed_scope'].update(case='ASDExactSourceFloorsCold',actual_import_this_pass=[],
    all257_actual_components_and_saved_book_bitexact1001=True,no_new_compile=True,only_prepared_stress_floors_changed=True)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('PREPARED1002 EXACT SOURCE FLOORS',floors,'ALL257 COMPONENTS/BOOK EXACT1001;CURVE KNOTS/TOLERANCES UNCHANGED')
