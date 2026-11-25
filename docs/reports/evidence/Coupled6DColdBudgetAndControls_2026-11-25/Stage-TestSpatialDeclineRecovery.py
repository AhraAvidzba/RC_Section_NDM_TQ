"""ДЛЯ ТЕСТОВ. Разрешает существующий возврат неподтверждённого спада для пространственных обобщённых мер, включая закручивание."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold855BalancedInitialControl_20261125';root=base/'Cold857SpatialDeclineRecovery_20261125'
proof=json.loads((parent/'BalancedInitialColdProof.json').read_bytes())
assert proof['cold']['capacity']=='NumFail' and proof['cold']['native']['own_process_closed']
log=(parent/'Native/HBackendStock.txt').read_text('utf-8-sig')
assert 'Суммарный спад найден, но его максимум не подтверждён повтором плеч.' in log
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'SpatialDeclineCompilationAncestor';archive.mkdir()
for n in ('CompilePoint.log','CompilePointClosure.json'):(root/n).rename(archive/n)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
name='src/Solver/CCoupledDeformationContinuation.cls';code=(root/name).read_text('utf-8-sig')
old='ElseIf reinforced And mSolver.UsesSpatialHistory And (mUsesTransverseControl Or mUsesFieldControl) Then';assert code.count(old)==1
code=code.replace(old,'ElseIf reinforced And mSolver.UsesSpatialHistory Then \' Пространственный возврат спада применим и к обобщённому закручиванию, без подтверждения самого пика.',1)
old='If reinforced And (mUsesTransverseControl Or mUsesFieldControl) Then';assert code.count(old)==2
code=code.replace(old,'If reinforced And (mSolver.UsesSpatialHistory Or mUsesTransverseControl Or mUsesFieldControl) Then')
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig');old="-ModuleNames 'modTestCoupledElasticState'";assert compiler.count(old)==1
values['Compile-TestExternalPoint.ps1']=compiler.replace(old,"-ModuleNames 'CCoupledDeformationContinuation'",1).replace('\n','\r\n').encode('utf-8-sig')
values.update({'NegativeInitialControlColdProof.json':(parent/'BalancedInitialColdProof.json').read_bytes(),
    'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for n,data in values.items():(root/n).write_bytes(data);prep['inputs'][n]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='SpatialGeneralizedControlDeclineRecovery',actual_import_this_pass=['CCoupledDeformationContinuation'],actual_component_count=256,
    only_three_Trace_routing_conditions_changed=True,existing_ReplayUnverifiedDecline_method_exact=True,
    all_Newton_arc_material_and_acceptance_math_exact855=True,original96_and30_limits=True,
    old_transverse_and_field_paths_preserved=True,spatial_twist_path_can_use_existing_unverified_decline_replay=True,
    unverified_peak_not_promoted_to_limit=True,cold_workflow_not_yet_run=True,physical_capacity_certified=False,no_working_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print(root,'THREE ROUTING CONDITIONS;EXISTING REPLAY;SAME PHYSICS/96/30;NO PEAK CERTIFICATE')
