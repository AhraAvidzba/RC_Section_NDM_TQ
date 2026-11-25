"""ДЛЯ ТЕСТОВ. Сохраняет внутреннюю точность дугового восстановления в последующем геометрическом продолжении; H и материалы не меняются."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,re

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold797ArcGrowthControl_20261125';root=base/'Cold800PostArcAccuracy_20261125'
prior=json.loads((base/'Cold798AdaptiveColdReplay_20261125/OperationalColdReplayProof.json').read_bytes())
assert prior['accepted_states']==67 and prior['states_above_arc_inner_tolerance']==57
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'PostArcAccuracyCompilationAncestor';archive.mkdir()
for n in ('CompilePoint.log','CompilePointClosure.json'):(root/n).rename(archive/n)
prep=json.loads((root/'Preparation.json').read_bytes());name='src/Solver/CCoupledDeformationContinuation.cls';code=(root/name).read_text('utf-8-sig')
anchor='Private mSignedArcProbes As Long, mSignedArcRecovered As Long\n';assert code.count(anchor)==1
code=code.replace(anchor,anchor+'Private mPostArcRefinedProbes As Long\n',1)
anchor='    mSignedArcRevision = -1: mSignedArcAttempts = 0: mSignedArcRatio = 1#: mSignedArcProbes = 0: mSignedArcRecovered = 0\n';assert code.count(anchor)==1
code=code.replace(anchor,anchor+'    mPostArcRefinedProbes = 0\n',1)
method,=re.findall(r'(?ms)^Private Function Probe\(.*?^End Function',code)
anchor='    If mProbeCount >= PROBE_LIMIT Then Exit Function\n';assert method.count(anchor)==1
updated=method.replace(anchor,anchor+'''    If mSignedArcRecovered > 0 And mSolver.UsesSpatialHistory And Not mUsesEnvelopeControl Then
        accuracyScale = GeomMin(accuracyScale, 0.0001) ' Последующие геометрические корни сохраняют точность исходной подтверждённой дуговой пары; иначе её малый шаг может оказаться ниже погрешности начального поля.
        mPostArcRefinedProbes = mPostArcRefinedProbes + 1
    End If
''',1)
code=code.replace(method,updated,1)
anchor='    If Len(mProbeDetails) > 0 Then Details = Details & " | Последняя неудачная проба: " & mProbeDetails\n';assert code.count(anchor)==1
code=code.replace(anchor,'    If mPostArcRefinedProbes > 0 Then Details = Details & "; геометрических проб с внутренней точностью дуги: " & CStr(mPostArcRefinedProbes)\n'+anchor,1)
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
values={name:code.replace('\n','\r\n').encode('utf-8-sig'),
    'Compile-TestExternalPoint.ps1':compiler.replace("-ModuleNames 'modTestCoupledElasticState'","-ModuleNames 'CCoupledDeformationContinuation'",1).replace('\n','\r\n').encode('utf-8-sig'),
    'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()}
for n,data in values.items():(root/n).write_bytes(data);prep['inputs'][n]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='ConsistentPostArcGeometricAccuracy',actual_import_this_pass=['CCoupledDeformationContinuation'],actual_component_count=256,
    Newton_arc_material_and_tests_exact797=True,after_verified_arc_only=True,inner_geometric_scale_matches_arc0001=True,
    envelope_H_original_accuracy_contract_unchanged=True,primary30_and_global96_unchanged=True,
    physical_acceptance_and_both_step_error_limits_unchanged=True,no_working_promotion=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('800 POST-ARC GEOMETRIC PRECISION SAME AS VERIFIED ARC; H/MATERIAL/30/96 UNCHANGED; NO PHYSICAL CAPACITY CLAIM')
