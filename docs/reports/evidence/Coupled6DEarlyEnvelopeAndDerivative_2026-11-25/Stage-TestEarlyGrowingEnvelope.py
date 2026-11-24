"""ДЛЯ ТЕСТОВ. Разрешает начальное H-управление пространственного CDP с арматурой только при уже проверенном росте H."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold745EarlyGrowingEnvelope_20261125';parent=base/'Cold743UnverifiedCumulativeRecovery_20261125'
proof=json.loads((parent/'CumulativeRecoveryColdProof.json').read_bytes())
independent=json.loads((base/'Cold744RecoveredColdFrozenReplay_20261125/RecoveredColdFrozenReplayProof.json').read_bytes())
assert proof['cold743']['native']['own_process_closed'] and independent['states']==proof['accepted_states']==32
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'EarlyEnvelopeCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
name='src/Solver/CCoupledDeformationContinuation.cls';code=(root/name).read_text('utf-8-sig')
old="""    If Not mReinforcement Is Nothing Then
        If mReinforcement.HasSteel Then Exit Sub ' Между изменениями бетона сталь может воспринимать нагрузку при практически постоянном H; начальная геометрическая мера сохраняется.
    End If"""
new="""    If Not mReinforcement Is Nothing Then
        If mReinforcement.HasSteel Then
            If Not mSolver.UsesSpatialHistory Or Not mSolver.SupportsHybridRecovery Then Exit Sub ' Начальная H-параметризация с арматурой проверяется только для пространственного CDP.
        End If
    End If"""
assert code.count(old)==1
code=code.replace(old,new,1)
assert code.replace(new,old,1)==(parent/name).read_text('utf-8-sig')
assert max(map(len,code.splitlines()))<1024
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig').replace("-ModuleNames 'modTestCoupledElasticState'","-ModuleNames 'CCoupledDeformationContinuation'")
values['Compile-TestExternalPoint.ps1']=compiler.replace('\n','\r\n').encode('utf-8-sig')
values['ParentCodeUpdate.json']=(parent/'docs/regression/CodeUpdate.json').read_bytes()
values['ParentCompileClosure.json']=(parent/'CompilePointClosure.json').read_bytes()
values[Path(__file__).name]=Path(__file__).read_bytes()
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope'].update(case='EarlyGrowingEnvelopeRC',actual_import_this_pass=['CCoupledDeformationContinuation'],
    only_initial_H_eligibility_changed=True,spatial_CDP_only_for_steel=True,
    original_positive_and_distinguishable_H_growth_required=True,
    original_TryLeaveEnvelopeControl_on_plateau_preserved=True,
    original96probes30iterations_and02_H_step_cap_preserved=True,
    material_Newton_and_test_module_exact743=True,no_working_promotion=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('745 EARLY GROWING H ELIGIBILITY; SAME MODEL/TOLERANCES/BUDGETS/PLATEAU FALLBACK; ONE CONTROLLER IMPORT')
