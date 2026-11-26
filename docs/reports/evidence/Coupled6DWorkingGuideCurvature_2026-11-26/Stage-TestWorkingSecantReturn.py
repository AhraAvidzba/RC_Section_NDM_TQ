"""ДЛЯ ТЕСТОВ. Добавляет общий прогноз по секущей в рабочий возврат и проверяет его с исходными половинными/бюджетными ограничителями."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold941OriginalControlReturn_20261126';source=base/'Cold938RecordedYieldAdvance_20261126'
root=base/'Cold947WorkingSecantReturn_20261126'
proof=json.loads((source/'RecordedYieldCoordinateProof.json').read_bytes())['cases'][2];assert proof['independently_equilibrated']
case=run_path(str(here/'Read-TestRecordedYieldCoordinate.py'))['TestRead'](source/'Native/RecordedYieldCoordinate.bin')[2]
assert case['info'][10]==17
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
ancestor=root/'WorkingSecantCompilationAncestor';ancestor.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(ancestor/name)
updates={};name='src/Solver/CCoupledNewtonSolver.cls';code=(root/name).read_text('utf-8-sig')
begin=code.index('Private Function TryReturnOriginalMaterialControl(');end=code.index('\nEnd Function',begin)
body=code[begin:end]
anchor='    On Error GoTo InterruptedReturn\n    DiscardStep';assert body.count(anchor)==1
change='''    On Error GoTo InterruptedReturn
    Dim acceptedCoordinate As Double, guideCoordinate As Double, fraction As Double
    For i = 1 To mGaugeStart - 1
        acceptedCoordinate = acceptedCoordinate + projection(i) * mAccepted(i)
        guideCoordinate = guideCoordinate + projection(i) * guide(i)
    Next i
    If guideCoordinate <> acceptedCoordinate Then
        fraction = (target - acceptedCoordinate) / (guideCoordinate - acceptedCoordinate)
        If fraction > 0# And fraction < 1# Then
            For i = 1 To mDegreeCount: seed(i) = mAccepted(i) + fraction * (guide(i) - mAccepted(i)): Next i
            details = details & "; прогноз исходной цели по секущей: доля " & CStr(fraction)
        End If
    End If
    DiscardStep'''
body=body.replace(anchor,change,1);code=code[:begin]+body+code[end:]
code=code.rstrip()+'\n\n'+(here/'TestMaterialGuideOriginalReturn.cls.txt').read_text('utf-8-sig').rstrip()+'\n'
updates[name]=code.replace('\n','\r\n').encode('utf-8-sig')
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
prefix="' FOR TESTS: baseline entry point only; existing methods retained.\n"
anchor=prefix*12+'\n\nPublic Function TestSolveOriginalControlReturn()';assert code.count(anchor)==1
entry=(here/'TestRunWorkingGuideReturn.bas.txt').read_text('utf-8-sig').rstrip()
updates[name]=code.replace(anchor,prefix*13+'\n\n'+entry+'\n\n\nPublic Function TestSolveOriginalControlReturn()',1).replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
updates['Compile-TestExternalPoint.ps1']=compiler.replace("-ModuleNames 'modTestCoupledElasticState'", "-ModuleNames 'CCoupledNewtonSolver','modTestCoupledElasticState' -EntryPointNames 'TestRunWorkingGuideReturn'",1).replace('\n','\r\n').encode('utf-8-sig')
name='tools/build_workbook/Run-Audit02ModeTests.ps1';runner=(root/name).read_text('utf-8-sig')
old='modTestCoupledElasticState.TestSolveOriginalControlReturn';assert runner.count(old)==1
updates[name]=runner.replace(old,'modTestCoupledElasticState.TestRunWorkingGuideReturn',1).replace('\n','\r\n').encode('utf-8-sig')
vector=run_path(str(here/'Stage-TestColdRawReplay.py'))['TestVector']
updates.update({'Input/RecordedYieldSeed.bin':(base/'Cold942OriginalControlFarReturn_20261126/Input/RecordedYieldSeed.bin').read_bytes(),
    'Input/GuideReturnEffort.bin':vector(case['info'][10:13]),'SourceMaterialRoot.bin':(source/'Native/RecordedYieldCoordinate.bin').read_bytes(),
    'VerifiedGuide938.json':json.dumps(proof,sort_keys=True,indent=2).encode(),
    'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for helper in ('TestMaterialGuideOriginalReturn.cls.txt','TestRunWorkingGuideReturn.bas.txt'):updates[helper]=(here/helper).read_bytes()
prep=json.loads((root/'Preparation.json').read_bytes())
for name,data in updates.items():
    if name.endswith(('.bas','.cls')):assert max(map(len,data.decode('utf-8-sig').splitlines()))<1024
    (root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='WorkingMaterialGuideReturnWithSecantInitialisation',
    actual_import_this_pass=['CCoupledNewtonSolver','modTestCoupledElasticState'],actual_component_count=257,
    generic_secant_field_only_no_hardcoded_point_or_material_level=True,guide_lambda_kept_as_initial_guess=True,
    working_half_uncertainty_direction_and_steel_guards_unchanged=True,
    recorded_guide17_steps_charged_before_return=True,total30_steps_enforced_by_working_method=True,
    case2_synthetic_exhausted30_step_guard=True,earlier_failed_primary_effort_not_included=True,
    no_history_acceptance=True,physical_capacity_certified=False,no_release_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'WORKING RETURN + SECANT;HALF/UNCERTAINTY GUARDS;17 GUIDE STEPS CHARGED')
