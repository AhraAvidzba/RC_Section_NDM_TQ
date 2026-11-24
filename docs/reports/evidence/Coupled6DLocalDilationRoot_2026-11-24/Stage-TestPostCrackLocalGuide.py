"""ДЛЯ ТЕСТОВ. Проверяет локальный прогноз из дробления с возвратом к исходному полному материальному шагу."""
from pathlib import Path
from hashlib import sha256
from runpy import run_path
import json
from argparse import ArgumentParser

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parser=ArgumentParser(description=__doc__);parser.add_argument('--root',type=Path,default=base/'Cold703PostCrackLocalGuide_20261124')
root=parser.parse_args().root.resolve();parent=base/'Cold702PostCrackIntegration_20261124'
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'LocalGuideCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
(root/'Matrix/Native').mkdir(parents=True,exist_ok=True)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
path='src/Materials/CCoupledConcreteResponse.cls';code=(root/path).read_text('utf-8-sig')
before='''Private Function UpdatePoint(ByRef strain() As Double, ByRef committed As TConcreteHistory, _
        ByRef updated As TConcreteHistory, ByRef apparent() As Double, ByRef virginElastic As Boolean) As Boolean'''
after='''Private Function UpdatePoint(ByRef strain() As Double, ByRef committed As TConcreteHistory, _
        ByRef updated As TConcreteHistory, ByRef apparent() As Double, ByRef virginElastic As Boolean, _
        Optional ByVal useGuide As Boolean = False, Optional ByVal guideTension As Double = 0#, _
        Optional ByVal guideCompression As Double = 0#, Optional ByVal guideDilation As Double = 0#) As Boolean'''
assert code.count(before)==1;code=code.replace(before,after,1)
before='    If Not firstReturnReady Then dilationRate = CurrentDilation(kt, kc, r, trialTrace, trialQ)\n'
assert code.count(before)==1
code=code.replace(before,before+'''    If useGuide Then
        tension = guideTension: compression = guideCompression: dilationRate = guideDilation
        mTension.Evaluate tension, unusedStress, ct, dt, kt
        mCompression.Evaluate compression, unusedStress, cc, dc, kc
        firstReturnReady = False
    End If
''',1)
before='        ReadPostCrackUpdateTest = UpdateWithSubsteps(strain, mHistory(1, mCommitted), updated, stress, virgin, substeps)'
assert code.count(before)==1
code=code.replace(before,'        ReadPostCrackUpdateTest = TestTryPostCrackLocalGuide(strain, stress, substeps)',1)
code+='\n'+(here/'TestPostCrackLocalGuide.cls.txt').read_text('utf-8-sig')
values[path]=code.replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig').replace("-ModuleNames 'modTestCoupledElasticState'","-ModuleNames 'CCoupledConcreteResponse'")
values['Compile-TestExternalPoint.ps1']=compiler.replace('\n','\r\n').encode('utf-8-sig')
values['ParentCodeUpdate.json']=(parent/'docs/regression/CodeUpdate.json').read_bytes()
values['ParentCompileClosure.json']=(parent/'CompilePointClosure.json').read_bytes()
for name in (Path(__file__).name,'TestPostCrackLocalGuide.cls.txt'):values[name]=(here/name).read_bytes()
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope'].update(case='LocalMaterialGuideAndOriginalFullStep',actual_import_this_pass=['CCoupledConcreteResponse'],
    guide_applies_only_to_test_readout_divisions0=True,original_full_curve_and_prefix_numerics_unchanged=True,
    original_material_equations_and_acceptance_tests_unchanged=True,intermediate_material_history_not_accepted=True,
    physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('LOCAL MATERIAL GUIDE STAGED; ORIGINAL FULL STEP RECHECKED; ONLY MATERIAL CLASS IMPORT')
