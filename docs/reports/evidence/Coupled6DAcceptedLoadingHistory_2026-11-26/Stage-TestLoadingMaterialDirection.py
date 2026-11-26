"""ДЛЯ ТЕСТОВ. Проверяет нагружающую сторону активной материальной координаты после подтверждённой упругой разгрузки977."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold977MaterialDirectionAudit_20261126';root=base/'Cold978LoadingMaterialDirection_20261126'
proof=json.loads((parent/'MaterialDirectionProof.json').read_bytes())
assert proof['ordinary_elastic_unloading_cases']==2 and proof['equilibrated_cases_with_resolved_H_growth']==0
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];r=R((parent/'Input/RecordedYieldSeed.bin').read_bytes());r.position=0
meta,current,accepted,residual,projection=[r.VectorTest() for _ in range(5)];assert r.position==len(r.raw)
origin=json.loads((parent/'MaterialPrefixOrigin.json').read_bytes());value=origin['accepted_coordinate_MPa']
old=float(meta[2]);assert old<value;meta[2]=value-(old-value);meta[9]=.5*(value+meta[2])
vector=run_path(str(here/'Stage-TestColdRawReplay.py'))['TestVector']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
updates={'Input/RecordedYieldSeed.bin':b''.join(vector(v) for v in (meta,current,accepted,residual,projection)),
    Path(__file__).name:Path(__file__).read_bytes(),
    'MaterialPrefixOrigin.json':(json.dumps(dict(point=int(meta[0]),side=int(meta[1]),
        station_source='actual968 nearest rejected-Newton material switch',accepted_coordinate_MPa=value,
        previous_unloading_full_target_MPa=old,full_target_MPa=float(meta[2]),half_target_MPa=float(meta[9]),
        target_increment_sign_reversed_to_loading=True,numerical_control_target_only_not_material_parameter=True,
        no_future_equilibrium_field_input=True,existing_H_steel_spatial_guards_preserved=True,
        half_seed_native_midpoint_and_actual_shared30_preserved=True),sort_keys=True,indent=2)+'\n').encode()}
prep=json.loads((root/'Preparation.json').read_bytes())
for name,data in updates.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='PositiveActiveMaterialLoadingCoordinateWithExistingDirectionGuards',actual_import_this_pass=[],actual_component_count=257,
    all257_actual_components_and_saved_book_bitexact977=True,no_new_numerical_code_or_compilation=True,
    only_internal_material_target_increment_sign_reversed=True,all_materials_geometry_force_path_tolerances_and_shared30_preserved=True,
    original_failed_seed968_and_no_future_equilibrium_input=True,native_H_steel_and_spatial_orientation_guards_preserved=True,
    no_material_history_acceptance=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'EXACT977 BOOK;LOADING SIGN OF INTERNAL TARGET ONLY;MATERIALS/FORCES/TOLERANCES/SHARED30/DIRECTION GUARDS UNCHANGED')
