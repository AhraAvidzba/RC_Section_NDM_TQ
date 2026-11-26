"""ДЛЯ ТЕСТОВ. Проверяет восемь рабочих НДС, отдельные наборы характеристик и таблицы нового ASD перед внешним материальным повторением."""
from pathlib import Path
from hashlib import sha256
from runpy import run_path
import json,numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ');root=base/'Cold1018SelectableMaterialCore_20261126'
native,text=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')
assert native==dict(passed=1,failed=0,macros=1,timeout600=False,own_process_closed=True)
assert 'SELECTABLE_CORE: both laws;' in text
source=json.loads((root/'SelectableCoreSourceProof.json').read_bytes())
assert sha256((root/'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper()==source['saved_book_sha256']
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];r=R((root/'Native/SelectableCore.bin').read_bytes())
assert r.raw[:16]==b'NDMLAWSEL0000001'
header=r.VectorTest();assert header.tolist()==[30000.,.2,16.,30.,3.]
contract=dict(concrete_E_MPa=30000.,poisson=.2,integration_mode='native_single_update_working',uniaxial_tables={})
branches={}
for expected in (0,1):
    code,count,floor=r.VectorTest();assert code==expected and int(count)==count
    rows=np.array([r.VectorTest() for _ in range(int(count))]);assert rows.shape==(int(count),3)
    # Экспорт вызывает Stresses даже в нуле: он возвращает численный floor.
    # Исходный нулевой узел таблицы обязателен по контракту Initialize; отдельно проверяем фактическое ограничение ординат.
    np.testing.assert_array_equal(rows[0],[0.,floor,floor])
    rows[0]=0.
    assert np.all(np.diff(rows[:,0])>0) and np.all(rows[:,2]>=rows[:,1])
    jumps=abs(np.diff(rows[:,1]));expected_floor=1e-6*min(jumps[jumps>0])
    assert abs(floor/expected_floor-1)<1e-10
    damage=np.zeros(int(count));damage[1:]=1-rows[1:,1]/rows[1:,2]
    table=np.column_stack([rows[:,:2],damage]);name='T' if code==0 else 'C'
    contract['uniaxial_tables'][name]=table.tolist();branches[name]=dict(point_count=int(count),stress_floor_MPa=floor,origin_numeric_floor_verified=True)
    assert abs(max(rows[:,1])-(3 if code==0 else 30))<1e-12
frames_count,=r.VectorTest();assert frames_count==9
frames=[]
for _ in range(int(frames_count)):
    strain=r.VectorTest();stress=r.VectorTest();assert len(strain)==len(stress)==6
    frames.append(dict(strain=[strain.tolist()],expected=[stress.tolist()]))
case_count,=r.VectorTest();assert case_count==8
cases=[];fields={}
for model in (1,2):
    for index in (1,2,3,4):
        observed_model,observed_index,seconds=r.VectorTest();assert (observed_model,observed_index)==(model,index) and seconds>=0
        target=r.VectorTest();actual=r.VectorTest();field=r.VectorTest();assert len(field)==60
        expected=np.zeros(6)
        if index in (1,4):expected[1]=500000.
        elif index==2:expected[3]=100000.
        else:expected[:]=[1000.,500000.,100000.,100000.,200.,-100.]
        np.testing.assert_array_equal(target,expected)
        assert max(abs((actual-target)/[1,180,180,180,1,1]))<.1
        cases.append(dict(model='CDP' if model==1 else 'ASD',case=index,value_set='SLS(II)' if index==4 else 'ULS(I)',
            target_N_Nmm=target.tolist(),actual_N_Nmm=actual.tolist(),seconds=seconds,
            maximum_resultant_error_N=float(max(abs((actual-target)/[1,180,180,180,1,1])))))
        fields[model,index]=field
assert r.position==len(r.raw)
gauge_differences={}
for index in (1,2,3,4):
    # Последние четыре координаты — множители служебных связей, не деформации.
    # Сохраняем их отличие отдельно; исходный допуск сравнения всех 56 физических координат остаётся прежним.
    np.testing.assert_allclose(fields[1,index][:56],fields[2,index][:56],atol=1e-10,rtol=0)
    gauge_differences[str(index)]=(fields[1,index][56:]-fields[2,index][56:]).tolist()
# В провайдере Eb общий для ULS/SLS; шестой аргумент Initialize — Ebt, а не второй Eb.
# Меняются сопротивления 30/3 -> 40/4; малый упругий изгиб при том же Eb обязан сохранить физическое поле.
np.testing.assert_allclose(fields[2,1][:56],fields[2,4][:56],atol=1e-10,rtol=0)
proof=dict(native=native,all_eight_working_core_cases_completed=True,source_unchanged=True,
    force_resultants_checked=True,CDP_ASD_56_physical_coordinates_agree=True,alternate_strength_value_set_completed=True,
    common_Eb_for_ULS_and_SLS_verified=True,alternate_strengths_not_compared_at_nonlinear_load=True,
    gauge_multipliers_are_not_strain_coordinates=True,gauge_multiplier_differences=gauge_differences,
    material_branches=branches,material_history_frames=len(frames),cases=cases,
    physical_capacity_not_tested=True,native_dropdown_UI_not_accepted=True,whole_goal_completed=False)
(root/'SelectableMaterialCoreProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
wire=dict(material_contract=contract,frames=frames,no_equilibrium_solve_in_external_material_replay=True,
    native_binary_sha256=sha256(r.raw).hexdigest().upper())
(root/'Native/SelectableCoreMaterialReplayInput.json').write_text(json.dumps(wire,separators=(',',':'))+'\n',encoding='utf-8')
print('WORKING SELECTABLE MATERIAL CORE VERIFIED',json.dumps(proof,sort_keys=True))
