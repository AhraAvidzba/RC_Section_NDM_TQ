"""ДЛЯ ТЕСТОВ. Добавляет плотную сетку вокруг прошлого секущего прогноза, не удаляя исходную сетку и её отрицательные результаты."""
from pathlib import Path
from hashlib import sha256
from runpy import run_path
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold706PostCrackDilationDense_20261124';parent=base/'Cold705PostCrackDilationRoot_20261124'
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'DenseCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
(root/'Matrix/Native').mkdir(parents=True,exist_ok=True)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
path='tests/modTestCoupledElasticState.bas';code=(root/path).read_text('utf-8-sig')
anchor='    Dim divisions As Variant, index As Long, previousValid As Boolean, valid As Boolean, component As Long\n';assert code.count(anchor)==1
code=code.replace(anchor,'''    Dim seed() As Double: seed = prediction
    operators.PointStrain 1, seed, strain
    Dim predictedLateral As Double: predictedLateral = 0.5 * (strain(1) + strain(2))
'''+anchor,1)
before='''        For index = -10 To 20
            value = index * 0.0000025''';assert code.count(before)==1
code=code.replace(before,'''        For index = -10 To 51
            If mLateralDivisionsTest <> 0 And index > 20 Then Exit For
            If index = 21 Then previousValid = False
            If index <= 20 Then
                value = index * 0.0000025
            Else
                value = predictedLateral + (index - 31) * 0.00000005
            End If''',1)
values[path]=code.replace('\n','\r\n').encode('utf-8-sig')
values['Compile-TestExternalPoint.ps1']=(here/'Compile-TestExternalPoint.ps1').read_bytes()
values['ParentCodeUpdate.json']=(parent/'docs/regression/CodeUpdate.json').read_bytes()
values['ParentCompileClosure.json']=(parent/'CompilePointClosure.json').read_bytes()
values[Path(__file__).name]=Path(__file__).read_bytes()
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope'].update(case='DilationClosureDensePastPredictor',actual_import_this_pass=['modTestCoupledElasticState'],
    original_coarse_grid_preserved=True,additional_dense_points=31,dense_center_uses_only_past_secant=True,
    material_code_and_tolerances_unchanged=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('DENSE GRID ADDED; PAST SECANT ONLY; ALL ORIGINAL COARSE SAMPLES RETAINED')
