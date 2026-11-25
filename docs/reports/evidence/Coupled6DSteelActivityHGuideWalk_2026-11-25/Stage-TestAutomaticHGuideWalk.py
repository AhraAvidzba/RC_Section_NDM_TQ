"""ДЛЯ ТЕСТОВ. Продолжает ветвь от истории70 без заранее заданного H-корня или границы нагружения."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold845SteelActivityArcPolicy_20261125';root=base/'Cold846AutomaticHGuideWalk_20261125'
proof=json.loads((parent/'SteelActivityArcPolicyProof.json').read_bytes());assert proof['automatic_Hguide_certificate_bit_identical_parent']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'AutomaticHGuideWalkCompilationAncestor';archive.mkdir()
for n in ('CompilePoint.log','CompilePointClosure.json'):(root/n).rename(archive/n)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
old='    TestReadSteelPlasticActivity\n    TestTrainedSteelPlasticActivity\n    TestNoSteelPlasticActivity\n    TestRunAutomaticHGuideArc\n'
assert code.count(old)==1
code=code.replace(old,'    TestRunContinuedOutgoingGuard\n    TestRunAutomaticLoadingWalk\n    TestRecentLoadingResetGuard\n',1)
old='If newPoints >= 8 Or totalCalls + 3 > 96';assert code.count(old)==1
code=code.replace(old,'If newPoints >= 32 Or totalCalls + 3 > 96',1)
code=code.replace("' Проверяет до восьми первых равновесий после проверенной границы нагружения;", "' Проверяет до32 новых равновесий общей рабочей операцией;",1)
anchor='        info = certificate(0)\n        If solved Then'
assert code.count(anchor)==1
code=code.replace(anchor,'''        info = certificate(0)
        TestWriteAutomaticHGuideWalk attempt, arc
        If solved Then''',1)
code+='''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет только фактически вычисленное H-направление до принятия корня; данные не возвращаются в алгоритм.
Private Sub TestWriteAutomaticHGuideWalk(ByVal attempt As Long, ByVal arc As CCoupledArcContinuation)
    Dim guide As Variant: guide = arc.EnvelopeGuideVerification
    Dim handle As Integer: handle = FreeFile
    Open mPolishRootTest & "\\Native\\LoadingWalk.Guide" & Format$(attempt, "000") & ".bin" For Binary Access Write As #handle
    Dim magic As String * 16: magic = "NDMWALKH00000001"
    Put #handle, , magic
    TestWriteReverseBinaryVector handle, Array(attempt, arc.NextStepRatioLimit, Abs(CLng(Not IsEmpty(guide))))
    If Not IsEmpty(guide) Then
        Dim vector As Variant
        For Each vector In guide: TestWriteReverseBinaryVector handle, vector: Next vector
    End If
    Close #handle
End Sub
'''
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
values.update({'Compile-TestExternalPoint.ps1':(here/'Compile-TestExternalPoint.ps1').read_bytes(),
    'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for name,raw in values.items():
    (root/name).write_bytes(raw);prep['inputs'][name]=sha256(raw).hexdigest().upper()
prep['directed_scope']=dict(case='AutomaticHGuideAndOutgoingWalkFrom70',actual_import_this_pass=['modTestCoupledElasticState'],actual_component_count=256,
    all_working_math_and_materials_exact845=True,input_exact70_no_future_field=True,maximum_calls96=True,maximum_attempts32=True,
    original_full_half_lambda_and_three_station_steel_criteria=True,automatic_Hguide_recorded_before_acceptance=True,
    no_cold_capacity_claim=True,physical_capacity_certified=False,no_working_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print(root)
