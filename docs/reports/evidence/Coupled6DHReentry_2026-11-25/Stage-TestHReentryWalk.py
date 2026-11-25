"""ДЛЯ ТЕСТОВ. Проверяет первые восемь автоматических шагов после независимо проверенной H-направленной границы нагружения."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold841HGuideMaterialEvent_20261125';root=base/'Cold842HReentryWalk_20261125'
proof=json.loads((parent/'HGuideMaterialEventProof.json').read_bytes())
assert proof['verified_local_bridge_to_loading_boundary'] and proof['zero_iteration_exact_restore_verified']
source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](parent/'Input/Reverse1.bin')
Reader=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];r=Reader((parent/'Native/HGuideEvent.bin').read_bytes())
record,info,state,face,half,candidate,restore,effort=[r.VectorTest() for _ in range(8)]
assert record[0]==1 and record[5]==2 and restore[0]==1
np.testing.assert_array_equal(face,candidate)
fields=np.vstack([source['fields'],candidate]);factors=np.append(source['lambdas'],record[8]);assert len(fields)==71
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'HReentryWalkCompilationAncestor';archive.mkdir()
for n in ('CompilePoint.log','CompilePointClosure.json'):(root/n).rename(archive/n)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
assert code.count('    TestHGuideMaterialEvent\n')==1
code=code.replace('    TestHGuideMaterialEvent\n','    TestRecentLoadingDirectionGuards\n    TestRunAutomaticLoadingWalk\n    TestRecentLoadingResetGuard\n',1)
assert code.count('If newPoints >= 32 Or totalCalls + 3 > 96')==1
code=code.replace('If newPoints >= 32 Or totalCalls + 3 > 96','If newPoints >= 8 Or totalCalls + 3 > 96',1)
code=code.replace("' Проверяет до32 новых равновесий общей рабочей операцией;", "' Проверяет до восьми первых равновесий после проверенной границы нагружения;",1)
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
values['Input/Reverse1.bin']=run_path(str(here/'Stage-TestColdRawReplay.py'))['TestTrajectoryBytes'](source,
    dict(fields=fields,fractions=np.ones(len(fields)),lambdas=factors))
values.update({'Compile-TestExternalPoint.ps1':(here/'Compile-TestExternalPoint.ps1').read_bytes(),
    'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for n,data in values.items():
    if n.endswith(('.bas','.cls')):assert max(map(len,data.decode('utf-8-sig').splitlines()))<1024
    (root/n).write_bytes(data);prep['inputs'][n]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='AutomaticDepartureFromVerifiedHGuideLoadingFace',actual_import_this_pass=['modTestCoupledElasticState'],actual_component_count=256,
    all_working_math_and_materials_exact841=True,exact70_prefix_plus_verified841_event=True,
    first_eight_new_points_diagnostic_scope=True,maximum_calls96=True,maximum_attempts32=True,
    original_full_half_and_lambda_and_three_station_steel_criteria=True,no_future_field_prescribed=True,
    no_cold_capacity_claim=True,physical_capacity_certified=False,no_working_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print(root)
