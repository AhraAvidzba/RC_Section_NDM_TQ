"""ДЛЯ ТЕСТОВ. Сравнивает автоматическое и фиксированное численное интегрирование одной неизменной материальной пробы."""
from pathlib import Path
from hashlib import sha256
from runpy import run_path
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold702PostCrackIntegration_20261124';parent=base/'Cold701PostCrackLateralRoot_20261124'
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'IntegrationCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
(root/'Matrix/Native').mkdir(parents=True,exist_ok=True)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
path='tests/modTestCoupledElasticState.bas';code=(root/path).read_text('utf-8-sig')
before=(here/'TestPostCrackLateralProbe.bas.txt').read_text('utf-8-sig')
assert code.count(before)==1
code=code.replace(before,(here/'TestPostCrackFixedIntegration.bas.txt').read_text('utf-8-sig'),1)
code=code.replace('Option Explicit\n','Option Explicit\nPrivate mLateralDivisionsTest As Long, mLateralSubstepsTest As Long\n',1)
values[path]=code.replace('\n','\r\n').encode('utf-8-sig')
path='src/Materials/CCoupledConcreteResponse.cls';code=(root/path).read_text('utf-8-sig')+'\n'+(here/'TestPostCrackFixedIntegration.cls.txt').read_text('utf-8-sig')
values[path]=code.replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig').replace("-ModuleNames 'modTestCoupledElasticState'","-ModuleNames @('CCoupledConcreteResponse','modTestCoupledElasticState')")
values['Compile-TestExternalPoint.ps1']=compiler.replace('\n','\r\n').encode('utf-8-sig')
values['ParentCodeUpdate.json']=(parent/'docs/regression/CodeUpdate.json').read_bytes()
values['ParentCompileClosure.json']=(parent/'CompilePointClosure.json').read_bytes()
for name in (Path(__file__).name,'TestPostCrackFixedIntegration.bas.txt','TestPostCrackFixedIntegration.cls.txt'):values[name]=(here/name).read_bytes()
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope'].update(case='ActualCDPIntegrationDiscontinuityTest',actual_import_this_pass=['CCoupledConcreteResponse','modTestCoupledElasticState'],
    physical_material_methods_unchanged=True,only_test_reader_added_to_concrete=True,fixed_divisions=[0,2,4,8,16,32,64],
    full_cold_prefix_replayed=True,working_material_unchanged=True)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('INTEGRATION DISCONTINUITY TEST STAGED; ACTUAL MATERIAL METHODS UNCHANGED; READER ONLY')
