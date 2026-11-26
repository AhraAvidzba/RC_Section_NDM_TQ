"""ДЛЯ ТЕСТОВ. Проверяет разгрузочный материальный шаг21/+1 и половину без изменения собранного VBA."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold972ActiveUnloadingCoordinate_20261126';root=base/'Cold973OutgoingMaterialPair_20261126'
proof=json.loads((parent/'MaterialPrefixCoordinateProof.json').read_bytes())
assert proof['cases'][2]['independently_equilibrated'] and proof['cases'][3]['independently_equilibrated']
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];r=R((parent/'Input/RecordedYieldSeed.bin').read_bytes());r.position=0
context,current,accepted,residual,projection=[r.VectorTest() for _ in range(5)];assert r.position==len(r.raw)
assert context[0]==21 and context[1]==1 and context[2]==0 and context[9]<0
# При принятом поле25 поверхность21/+1 равна нулю с точностью арифметики; точное значение сохранено независимой970.
accepted_value=json.loads((base/'Cold970PrefixSurfaceMatrices_20261126/SurfaceMatricesProof.json').read_bytes())['point21_plus_loading_value_at_accepted25_MPa']
full=context[9];half=.5*(accepted_value+full);context[2]=full;context[9]=half
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
vector=run_path(str(here/'Stage-TestColdRawReplay.py'))['TestVector']
updates={'Input/RecordedYieldSeed.bin':b''.join(vector(v) for v in (context,current,accepted,residual,projection)),
    Path(__file__).name:Path(__file__).read_bytes(),
    'MaterialPrefixOrigin.json':(json.dumps(dict(point=21,side=1,station_source='actual968 nearest unloading crossing',
        accepted_coordinate_MPa=accepted_value,full_target_MPa=float(full),half_target_MPa=float(half),
        full_target_from_actual_already_evaluated_Newton_trial=True,half_target_of_original_material_increment=True,
        no_future_equilibrium_field_in_consumed_fixture=True,starting_field_is_actual_failed968_iteration=True,
        each_independent_case30_steps_not_joint_native_pair_acceptance=True),sort_keys=True,indent=2)+'\n').encode()}
prep=json.loads((root/'Preparation.json').read_bytes())
for name,data in updates.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope'].update(case='Actual21PlusOutgoingMaterialFullAndHalf',all257_actual_components_and_saved_book_bitexact972=True,
    no_new_numerical_code_or_compilation=True,original_force_path_materials_and30_per_independent_case=True,
    independent_boundary_and_unloading_targets_not_full_half_pair=False,full_and_half_targets_of_same_material_increment=True,
    no_material_history_acceptance=True,not_joint_native_full_half_recovery_acceptance=True)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'EXACT972 BOOK;21/+1 OUTGOING FULL/HALF;MATERIAL AND FORCE PATH UNCHANGED;INDEPENDENT30 EACH')
