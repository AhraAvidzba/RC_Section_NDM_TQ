"""ДЛЯ ТЕСТОВ. Исключает из выбора начала пластического нагружения события, меньшие уже обязательного разрешения половинного подхода."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,re

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold789RejectedOperationalArc_20261125';root=base/'Cold791ResolvedLoadingEvent_20261125'
native,_=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](parent,'RealColdReplay');assert native['failed']==0
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'ResolvedEventCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
prep=json.loads((root/'Preparation.json').read_bytes());name='src/Solver/CCoupledNewtonSolver.cls';code=(root/name).read_text('utf-8-sig')
method,=re.findall(r'(?ms)^Friend Function TryLoadingYieldBoundary\(.*?^End Function',code)
anchor='    Dim point As Long, side As Long: nearest = 1#\n';assert method.count(anchor)==1
updated=method.replace(anchor,anchor+'''    Dim eventRadius As Double: eventRadius = SpatialFieldDistance(mAccepted, ends)
    Dim resolution As Double
    resolution = 16# * GeomMax(ControlledLoadTolerance(baseLoad, direction, fromLambda, absoluteTolerance, relativeTolerance), _
        ControlledLoadTolerance(baseLoad, direction, toLambda, absoluteTolerance, relativeTolerance)) ' Половинный подход должен превышать восемь фактических допусков равновесия; неразличимая исходная граница не заслоняет остальные события.
''',1)
old='''                        If Not selected Or fraction < nearest Then station = candidate: nearest = fraction: selected = True
'''
assert updated.count(old)==1
updated=updated.replace(old,'''                        If fraction * eventRadius * mReferenceModulus > resolution Then
                            If Not selected Or fraction < nearest Then station = candidate: nearest = fraction: selected = True
                        End If
''',1)
updated=updated.replace('    radius = SpatialFieldDistance(mAccepted, ends)\n','    radius = eventRadius\n',1)
code=code.replace(method,updated,1)
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
values={name:code.replace('\n','\r\n').encode('utf-8-sig'),
    'Compile-TestExternalPoint.ps1':compiler.replace("-ModuleNames 'modTestCoupledElasticState'","-ModuleNames 'CCoupledNewtonSolver'",1).replace('\n','\r\n').encode('utf-8-sig'),
    'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()}
for n,data in values.items():(root/n).write_bytes(data);prep['inputs'][n]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='ResolvedLoadingEventSelection',actual_import_this_pass=['CCoupledNewtonSolver'],actual_component_count=256,
    exact_cold_prefix10_and_ratio1_unchanged=True,only_operational_loading_event_selection_changed=True,
    event_selection_resolution_matches_existing_half_approach_guard=True,
    original_equilibrium_material_force_gauge_and_half_error_acceptance_unchanged=True,
    material_and_arc_class_and_controller_exact789=True,no_working_promotion=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('791 SAME COLD10/RATIO1; SKIP UNRESOLVED LOADING EVENTS BEFORE NEAREST CHOICE; ALL ACCEPTANCE/BUDGETS SAME')
