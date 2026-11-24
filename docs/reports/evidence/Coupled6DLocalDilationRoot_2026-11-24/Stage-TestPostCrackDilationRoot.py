"""ДЛЯ ТЕСТОВ. Решает исходную локальную согласованность потока скаляром вместо неустойчивой итерации дилатансии."""
from pathlib import Path
from hashlib import sha256
from runpy import run_path
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold705PostCrackDilationRoot_20261124';parent=base/'Cold704PostCrackLocalGuide_20261124'
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'DilationRootCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
(root/'Matrix/Native').mkdir(parents=True,exist_ok=True)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
path='src/Materials/CCoupledConcreteResponse.cls';code=(root/path).read_text('utf-8-sig')
old=(here/'TestPostCrackLocalGuide.cls.txt').read_text('utf-8-sig');assert code.count(old)==1
code=code.replace(old,(here/'TestPostCrackDilationRoot.cls.txt').read_text('utf-8-sig'),1)
before='        Optional ByVal guideCompression As Double = 0#, Optional ByVal guideDilation As Double = 0#) As Boolean'
assert code.count(before)==1
code=code.replace(before,'''        Optional ByVal guideCompression As Double = 0#, Optional ByVal guideDilation As Double = 0#, _
        Optional ByVal holdDilationTest As Boolean = False, Optional ByRef flowResidualTest As Double = 0#, _
        Optional ByRef flowReadyTest As Boolean = False) As Boolean''',1)
before='    tension = committed.Tension: compression = committed.Compression\n';assert code.count(before)==1
code=code.replace(before,before+'    flowReadyTest = False\n',1)
before='        UpdatePoint = True\n        Exit Function\n';assert code.count(before)==1
code=code.replace(before,'        If holdDilationTest Then flowResidualTest = 0#: flowReadyTest = True\n'+before,1)
before='        apexTransition = (j > MATERIAL_ITERATIONS)\n';assert code.count(before)==1
code=code.replace(before,before+'        If apexTransition And holdDilationTest Then Exit Function\n',1)
before='                If j <> 1 Or dilationRate <> 0# Or mDilationMax = 0# Then Exit Function';assert code.count(before)==1
code=code.replace(before,before.replace(' Then',' Or holdDilationTest Then'),1)
before='''        If apexTransition Then
            If r <> 1# Or nextCompression <> committed.Compression Then Exit Function'''
assert code.count(before)==1
code=code.replace(before,'''        If holdDilationTest Then
            If Abs(deltaTension) <= STATE_TOLERANCE * (mFt / mElastic(3, 3) + nextTension) And _
                    Abs(deltaCompression) <= STATE_TOLERANCE * (mFc / mElastic(3, 3) + nextCompression) And _
                    Abs(residual) <= YIELD_TOLERANCE_RATIO * cc Then
                flowResidualTest = deltaDilation
                flowReadyTest = (Abs(deltaDilation) <= STATE_TOLERANCE And _
                    DilationConsistent(Abs(deltaDilation), plasticMultiplier, multiplierDivisor, YIELD_TOLERANCE_RATIO * cc / mPressureModulus))
                Exit For
            End If
        Else
'''+before,1)
before='''        tension = tension + MAXIMUM_DAMPING * deltaTension''';assert code.count(before)==1
code=code.replace(before,'        End If\n'+before,1)
before='        dilationRate = dilationRate + MAXIMUM_DAMPING * deltaDilation';assert code.count(before)==1
code=code.replace(before,'        If Not holdDilationTest Then '+before.strip(),1)
values[path]=code.replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig').replace("-ModuleNames 'modTestCoupledElasticState'","-ModuleNames 'CCoupledConcreteResponse'")
values['Compile-TestExternalPoint.ps1']=compiler.replace('\n','\r\n').encode('utf-8-sig')
values['ParentCodeUpdate.json']=(parent/'docs/regression/CodeUpdate.json').read_bytes()
values['ParentCompileClosure.json']=(parent/'CompilePointClosure.json').read_bytes()
for name in (Path(__file__).name,'TestPostCrackDilationRoot.cls.txt'):values[name]=(here/name).read_bytes()
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope'].update(case='OriginalFullStepDilationClosure',actual_import_this_pass=['CCoupledConcreteResponse'],
    scalar_dilation_root_with_original_acceptance=True,unsealed_inner_probe_never_published=True,
    global_curve_and_original_prefix_unchanged=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('LOCAL DILATION ROOT STAGED; ORIGINAL FULL STEP; UNSEALED INNER PROBES NOT ACCEPTED')
