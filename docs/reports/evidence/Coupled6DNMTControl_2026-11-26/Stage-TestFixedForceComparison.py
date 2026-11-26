"""ДЛЯ ТЕСТОВ. Готовит одинаковые силовые уравнения для штатного Newton и SciPy на сохранённом максимуме и выше него."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold986AdaptiveLoadingHistory_20261126';root=base/'Cold988FixedForceExternal_20261126'
assert json.loads((parent/'PointTraceSourceComposition.json').read_bytes())['components']==257
data=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](parent/'Input/Reverse1.bin')
peak=int(np.argmax(data['lambdas']));factor=float(data['lambdas'][peak]);assert peak==15
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
ancestor=root/'FixedForceCompilationAncestor';ancestor.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(ancestor/name)
entry=(here/'TestFixedForceComparison.bas.txt').read_text('utf-8-sig').rstrip()
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
anchor='Private mReverseSidesTest As Collection, mReverseFieldsTest As Collection, mReverseTraceTest As CCoupledLoadTrace';assert code.count(anchor)==1
code=code.replace(anchor,anchor+'\nPrivate mFixedForceLambdaTest As Double, mFixedForcePrefixTest As Long, mFixedForceRevisionTest As Long\nPrivate mFixedForceSeedTest() As Double, mFixedForceProjectionTest() As Double',1)
anchor='Public Function TestRunAcceptedLoadingHistory() As String';assert code.count(anchor)==1
code=code.replace(anchor,entry+'\n\n'+anchor,1)
vector=run_path(str(here/'Stage-TestColdRawReplay.py'))['TestVector']
updates={name:code.replace('\n','\r\n').encode('utf-8-sig'),
    'Compile-TestExternalPoint.ps1':(here/'Compile-TestExternalPoint.ps1').read_bytes(),
    'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes(),
    'TestFixedForceComparison.bas.txt':(entry+'\n').replace('\n','\r\n').encode('utf-8-sig'),
    'Input/FixedForceCase.bin':vector([factor,peak+1,peak+1]),
    'Input/FixedForceCase.json':(json.dumps(dict(target_lambda=factor,accepted_prefix=peak+1,seed_index=peak+1,
        known_peak_lambda=factor,target_ratio_to_known_peak=1.,same60_equations_native_and_external=True,
        material_history_is_known_peak_prefix_not_virgin_monotone_path=True,physical_capacity_certified=False),sort_keys=True,indent=2)+'\n').encode()}
prep=json.loads((root/'Preparation.json').read_bytes())
for name,data in updates.items():
    if name.endswith('.bas'):assert max(map(len,data.decode('utf-8-sig').splitlines()))<1024
    (root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='NativeVersusSciPyFixedForce',actual_import_this_pass=['modTestCoupledElasticState'],actual_component_count=257,
    all256_numerical_components_exact986=True,only_new_test_COM_callbacks_and_native_probe=True,
    exactly_fixed_lambda60_field_unknowns=True,no_material_history_acceptance=True,no_future_equilibrium_input=True,
    physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'FIXED FORCE NATIVE/SCIPY;KNOWN MAX',factor,'PREFIX',peak+1,'ALL256 NUMERICAL COMPONENTS UNCHANGED')
