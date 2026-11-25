"""ДЛЯ ТЕСТОВ. Проверяет общую автоматическую последовательность с новым выбором исходящего шага от исходной истории70."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold851LoadingSecantReparameterization_20261125';root=base/'Cold852AutomaticLoadingSecantWalk_20261125'
proof=json.loads((parent/'LoadingSecantReparameterizationProof.json').read_bytes())
assert not proof['cases'][0]['arc_solved'] and all(p['arc_solved'] for p in proof['cases'][1:])
initial=base/'Cold846AutomaticHGuideWalk_20261125/Input/Reverse1.bin'
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0,binary=initial)
archive=root/'AutomaticLoadingSecantCompilationAncestor';archive.mkdir()
for n in ('CompilePoint.log','CompilePointClosure.json'):(root/n).rename(archive/n)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
old='    TestDiagnoseNewLoadingCoordinate\n';assert code.count(old)==1
code=code.replace(old,'    TestRunContinuedOutgoingGuard\n    TestRunAutomaticLoadingWalk\n    TestRecentLoadingResetGuard\n',1)
assert code.count('If newPoints >= 32 Or totalCalls + 3 > 96')==1
start=code.index('Private Sub TestWriteAutomaticHGuideWalk(');end=code.index('\nEnd Sub',start)
method=code[start:end];old='magic = "NDMWALKH00000001"';assert method.count(old)==1
method=method.replace(old,'magic = "NDMWALKH00000002"',1)
anchor='    Close #handle';assert method.count(anchor)==1
method=method.replace(anchor,'''    Dim loading As Variant: loading = arc.LoadingGuideVerification
    TestWriteReverseBinaryVector handle, Array(Abs(CLng(Not IsEmpty(loading))))
    If Not IsEmpty(loading) Then
        Dim loadingVector As Variant
        For Each loadingVector In loading: TestWriteReverseBinaryVector handle, loadingVector: Next loadingVector
        Dim loadingField() As Double: loadingField = loading(1)
        TestWriteElasticPrincipalRoot 1000 + attempt, loadingField
    End If
'''+anchor,1)
code=code[:start]+method+code[end:];values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
values.update({'Compile-TestExternalPoint.ps1':(here/'Compile-TestExternalPoint.ps1').read_bytes(),
    'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for n,data in values.items():(root/n).write_bytes(data);prep['inputs'][n]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='AutomaticHGuideAndLoadingSecantWalkFrom70',actual_import_this_pass=['modTestCoupledElasticState'],actual_component_count=256,
    all_working_math_and_materials_exact851=True,input_exact70_no_future_field=True,maximum_calls96=True,maximum_attempts32=True,
    original_full_half_lambda_locality_and_three_station_steel_criteria=True,H_and_loading_guides_recorded_before_acceptance=True,
    no_cold_capacity_claim=True,physical_capacity_certified=False,no_working_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print(root,'EXACT70;AUTOMATIC H AND LOADING GUIDES;ORIGINAL96 CALLS')
