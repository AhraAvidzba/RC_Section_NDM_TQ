"""ДЛЯ ТЕСТОВ. Переиспользует собранный вход936 для материальных префиксов точки18, выбранной автоматической дугой962."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold968PrefixNewtonCapture_20261126';root=base/'Cold971MaterialPrefixCoordinate_20261126'
origin=json.loads((parent/'PrefixNewtonCaptureProof.json').read_bytes());record=origin['rejected_newton_records'][1]
raw=(parent/'Native'/record['file']).read_bytes();assert sha256(raw).hexdigest().upper()==record['sha256']
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];r=R(raw);r.position=0
meta,current,accepted,residual,rhs,delta=[r.VectorTest() for _ in range(6)]
matrix=np.array([r.VectorTest() for _ in range(61)]);projection=matrix[-1,:60]/meta[4]
source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](parent/'Input/Reverse1.bin')
np.testing.assert_array_equal(accepted,source['fields'][-1])
guide_proof=json.loads((base/'Cold962CurrentOperationalArc_20261126/CurrentOperationalArcProof.json').read_bytes())
assert guide_proof['no_supplied_future_guide_or_field'];event=guide_proof['event'];point,side=event['point'],event['side']
assert (point,side)==(18,0)
g,L,w,v,coords,axis,B,gradient=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](parent/'Native/EnvelopeEnergy.Accepted.Operators.csv',60,43)
op=np.array([B+s*.001*gradient for s in (-1,0,1)])
E,nu=source['calibration'][:2];G=E/(2*(1+nu));off=E*nu/((1+nu)*(1-2*nu))
D=np.diag([2*G]*3+[G]*3);D[:3,:3]+=off
r=R((parent/'SourceAcceptedSurfaceStates.bin').read_bytes());states={}
for _ in range(129):
    p,s=r.VectorTest();states[int(p),int(s)]=r.VectorTest()
assert r.position==len(r.raw)
helper=(here/'Check-TestCold32OutgoingStressScan.py').read_text('utf-8-sig');begin=helper.index('def TestYield(');end=helper.index('def TestAssemble(',begin)
exec(compile(helper[begin:end],str(here/'Check-TestCold32OutgoingStressScan.py'),'exec'),globals())
initial=TestYield(point,side,accepted);assert initial<0
fraction=1/16;full=initial*(1-fraction);half=.5*(initial+full)
event_meta=[float(point),float(side),full,meta[2],meta[5],source['options'][3],source['options'][4],source['options'][2],meta[4],half]
vector=run_path(str(here/'Stage-TestColdRawReplay.py'))['TestVector']
fixture=b''.join(vector(v) for v in (event_meta,current,accepted,residual,projection))
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
updates={'Input/RecordedYieldSeed.bin':fixture,'SourceActualPrefixRejected.bin':raw,
    'SourceAutomatic962ArcProof.json':(base/'Cold962CurrentOperationalArc_20261126/CurrentOperationalArcProof.json').read_bytes(),
    Path(__file__).name:Path(__file__).read_bytes()}
name='tools/build_workbook/Run-Audit02ModeTests.ps1';runner=(root/name).read_text('utf-8-sig')
old='            $null = $excel.Run("\'RC_Section_NDM.xlsm\'!modTestColdNewtonCapture.TestSetCaptureRoot", $MacroArgument1)\n';assert runner.count(old)==1
runner=runner.replace(old,'',1);old='modTestCoupledElasticState.TestRunControlledSecantPrefix';assert runner.count(old)==1
updates[name]=runner.replace(old,'modTestCoupledElasticState.TestSolveRecordedYieldCoordinate',1).replace('\n','\r\n').encode('utf-8-sig')
updates['MaterialPrefixOrigin.json']=(json.dumps(dict(automatic_station_source='962 existing OperationalArc event',point=point,side=side,
    accepted_coordinate_MPa=initial,fraction_to_zero=fraction,full_target_MPa=full,half_target_MPa=half,
    starting_field_is_actual_rejected968_iteration_not_equilibrium=True,source_record=record['file'],
    no_future_equilibrium_field_in_consumed_fixture=True,original_force_path_and_materials_preserved=True,
    each_independent_full_or_half_case_limited30_not_jointly_accepted=True),sort_keys=True,indent=2)+'\n').encode()
prep=json.loads((root/'Preparation.json').read_bytes())
for name,data in updates.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='ExistingMaterialCoordinatePrefixAndHalfOfProgramSelected18',actual_import_this_pass=[],actual_component_count=257,
    all257_actual_components_and_saved_book_bitexact968=True,no_new_numerical_code_or_compilation=True,
    automatic_point18_selection_reused_as_provenance_not_rerun_inside_this_entry=True,
    no_future_equilibrium_field_in_consumed_fixture=True,original25_history_and_force_path_and_materials_preserved=True,
    original30_per_independent_case=True,not_a_shared_full_half30_acceptance_operation=True,
    no_material_history_acceptance=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'EXACT968 BOOK;EXISTING MATERIAL ENTRY;POINT18 FULL/HALF;NO FUTURE ROOT;INDEPENDENT30 EACH;NOT CAPACITY')
