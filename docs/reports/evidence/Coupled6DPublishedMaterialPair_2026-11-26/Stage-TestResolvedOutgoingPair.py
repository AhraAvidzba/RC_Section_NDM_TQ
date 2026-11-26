"""ДЛЯ ТЕСТОВ. Увеличивает только шаг материального управления для проверки разрешимого разгрузочного продвижения."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold974MidpointMaterialPair_20261126';root=base/'Cold975ResolvedOutgoingPair_20261126'
proof=json.loads((parent/'MidpointMaterialPairProof.json').read_bytes())
assert proof['pair']['passes_existing_independent_half_limits'] and not proof['pair']['resolved_spatial_field_advance']
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];r=R((parent/'Input/RecordedYieldSeed.bin').read_bytes());r.position=0
meta,current,accepted,residual,projection=[r.VectorTest() for _ in range(5)];assert r.position==len(r.raw)
accepted_value=json.loads((base/'Cold970PrefixSurfaceMatrices_20261126/SurfaceMatricesProof.json').read_bytes())['point21_plus_loading_value_at_accepted25_MPa']
factor=16.;old_full=float(meta[2]);meta[2]=accepted_value+factor*(old_full-accepted_value);meta[9]=.5*(accepted_value+meta[2])
vector=run_path(str(here/'Stage-TestColdRawReplay.py'))['TestVector']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
updates={'Input/RecordedYieldSeed.bin':b''.join(vector(v) for v in (meta,current,accepted,residual,projection)),
    Path(__file__).name:Path(__file__).read_bytes(),
    'MaterialPrefixOrigin.json':(json.dumps(dict(point=21,side=1,station_source='actual968 nearest unloading crossing',
        accepted_coordinate_MPa=accepted_value,full_target_MPa=float(meta[2]),half_target_MPa=float(meta[9]),
        numerical_control_increment_factor_relative974=factor,source974_full_target_MPa=old_full,
        numerical_coordinate_step_only_not_material_parameter=True,no_future_equilibrium_field_in_consumed_fixture=True,
        half_seed_native_midpoint_and_actual_shared30_preserved=True),sort_keys=True,indent=2)+'\n').encode()}
prep=json.loads((root/'Preparation.json').read_bytes())
for name,data in updates.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope'].update(case='ResolvedOutgoingMaterialPairWith16TimesNumericalIncrement',actual_import_this_pass=[],
    all257_actual_components_and_saved_book_bitexact974=True,no_new_numerical_code_or_compilation=True,
    numerical_material_coordinate_increment_only_changed=True,original_materials_force_path_tolerances_and_shared30_unchanged=True)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'EXACT974 BOOK;NUMERICAL CONTROL INCREMENTx16 ONLY;SAME MATERIALS/FORCES/TOLERANCES/SHARED30')
