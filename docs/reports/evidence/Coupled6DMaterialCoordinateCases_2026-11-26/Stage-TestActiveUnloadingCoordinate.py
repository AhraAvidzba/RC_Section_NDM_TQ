"""ДЛЯ ТЕСТОВ. Проверяет действительную разгрузочную поверхность21/+1 тем же собранным решателем без изменения VBA."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold971MaterialPrefixCoordinate_20261126';root=base/'Cold972ActiveUnloadingCoordinate_20261126'
proof=json.loads((parent/'MaterialPrefixCoordinateProof.json').read_bytes())
assert proof['independently_verified_material_coordinate_equilibria']==0
origin=base/'Cold968PrefixNewtonCapture_20261126';rejected=json.loads((origin/'PrefixNewtonCaptureProof.json').read_bytes())['rejected_newton_records'][1]
nearest=rejected['probes'][-1]['elastic_predictor_loading_surface_sign_changes']
assert len(nearest)==1
station=next(s for s in rejected['probes'][0]['elastic_predictor_loading_surface_sign_changes']
    if (s['point'],s['side'])==(nearest[0]['point'],nearest[0]['side']))
assert station['from_value_MPa']>0>station['to_value_MPa']
point,side=station['point'],station['side'];assert (point,side)==(21,1)
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];r=R((parent/'Input/RecordedYieldSeed.bin').read_bytes());r.position=0
context,current,accepted,residual,projection=[r.VectorTest() for _ in range(5)];assert r.position==len(r.raw)
context[0]=point;context[1]=side;context[2]=0.;context[9]=station['to_value_MPa']
vector=run_path(str(here/'Stage-TestColdRawReplay.py'))['TestVector']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
updates={'Input/RecordedYieldSeed.bin':b''.join(vector(v) for v in (context,current,accepted,residual,projection)),
    Path(__file__).name:Path(__file__).read_bytes()}
updates['MaterialPrefixOrigin.json']=(json.dumps(dict(station_source='actual968 rejected Newton record2;full fraction1 sign crossing',
    point=point,side=side,boundary_target_MPa=0.,unloading_target_MPa=station['to_value_MPa'],
    source_current_trial_loading_value_MPa=station['from_value_MPa'],starting_field_is_actual_failed968_not_equilibrium=True,
    no_future_equilibrium_field_in_consumed_fixture=True,not_a_full_half_pair=True,
    targets_derived_from_actual_already_evaluated_Newton_trial=True,original_materials_and_force_path_preserved=True),sort_keys=True,indent=2)+'\n').encode()
prep=json.loads((root/'Preparation.json').read_bytes())
for name,data in updates.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope'].update(case='ActualRecordedActive21PlusBoundaryAndUnloadingCoordinates',
    all257_actual_components_and_saved_book_bitexact971=True,no_new_numerical_code_or_compilation=True,
    station_selected_from_actual_sign_crossing_not_fixed_by_engineering_input=True,
    independent_boundary_and_unloading_targets_not_full_half_pair=True,no_material_history_acceptance=True)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'EXACT971 BOOK;ACTUAL21/+1 CROSSING;BOUNDARY AND UNLOADING;NO NEW VBA OR MATERIAL CHANGE;INDEPENDENT30 EACH')
