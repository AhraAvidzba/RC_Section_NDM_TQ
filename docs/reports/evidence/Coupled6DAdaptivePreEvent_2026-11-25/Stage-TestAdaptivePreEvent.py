"""ДЛЯ ТЕСТОВ. Встраивает ограниченное уточнение подхода перед событием в рабочую дуговую операцию; исходная холодная история10 неизменна."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,re

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold791ResolvedLoadingEvent_20261125';root=base/'Cold793AdaptivePreEvent_20261125'
quarter=json.loads((base/'Cold792PreEventQuarter_20261125/PreEventQuarterProof.json').read_bytes())
assert not quarter['half_verified'] and quarter['field_half_error']<.005 and quarter['lambda_half_error']>quarter['lambda_half_bound']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'AdaptivePreEventCompilationAncestor';archive.mkdir()
for n in ('CompilePoint.log','CompilePointClosure.json'):(root/n).rename(archive/n)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
name='src/Solver/CCoupledNewtonSolver.cls';code=(root/name).read_text('utf-8-sig')
code+='''
' Читает пробную поверхность CDP от принятой пластической истории; состояние материала не обновляется.
Friend Function TryReadLoadingSurface(ByVal point As Long, ByVal side As Long, ByRef field() As Double, ByRef value As Double) As Boolean
    If point < 1 Or point > mPointCount Or side < -1 Or side > 1 Then Err.Raise 5, "CCoupledNewtonSolver", "Неверная точка или станция материальной поверхности."
    If LBound(field) <> 1 Or UBound(field) <> mDegreeCount Then Err.Raise 5, "CCoupledNewtonSolver", "Размер поля не совпадает с подготовленным сечением."
    Dim station As TYieldStation, strain(1 To COMPONENT_COUNT) As Double, unused(1 To COMPONENT_COUNT) As Double
    station.Point = point: station.Side = side
    YieldStrain point, side, field, strain
    TryReadLoadingSurface = ReadTransitionValue(station, strain, value, unused, False)
End Function
'''
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
name='src/Solver/CCoupledArcContinuation.cls';code=(root/name).read_text('utf-8-sig')
code=code.replace('Private mCertificate As Variant, mDetails As String, mError As Double\n','Private mCertificate As Variant, mDetails As String, mError As Double\nPrivate mApproaches As Collection\n',1)
anchor='    mCalls = 0: mPrimary = 0: mError = -1#: mDetails = vbNullString: mCertificate = Empty\n';assert code.count(anchor)==1
code=code.replace(anchor,anchor+'    Set mApproaches = New Collection\n',1)
anchor='    accepted = fields(count): previous = fields(count - 1): movement = accepted\n';assert code.count(anchor)==1
code=code.replace(anchor,anchor+'''    Dim actualAccepted() As Double, degree As Long: mSolver.CopyAccepted actualAccepted
    If UBound(actualAccepted) <> UBound(accepted) Then Err.Raise 5, "CCoupledArcContinuation", "Размер принятого поля не совпадает с историей продолжения."
    For degree = 1 To UBound(accepted)
        If actualAccepted(degree) <> accepted(degree) Then Err.Raise 5, "CCoupledArcContinuation", "Принятая материальная история не соответствует последнему полю продолжения."
    Next degree
''',1)
code=code.replace('    Dim faceActual() As Double, faceHalfActual() As Double, eventHalfLambda As Double, eventError As Double, eventSolved As Boolean\n',
    '    Dim faceActual() As Double, faceHalfActual() As Double, eventHalfLambda As Double, eventError As Double, eventSolved As Boolean\n    Dim eventLambda As Double, selectedKind As Long\n',1)
code=code.replace('        candidate = full: candidateLambda = lf: phaseSteps = pairSteps\n','        candidate = full: candidateLambda = lf: phaseSteps = pairSteps: selectedKind = 1\n',1)
old='            eventInfo, face, faceHalf, faceActual, faceHalfActual, eventHalfLambda, eventError, eventSolved)'
new='            eventInfo, face, faceHalf, faceActual, faceHalfActual, eventHalfLambda, eventError, eventSolved, callAllowance, eventLambda, selectedKind)'
assert code.count(old)==1;code=code.replace(old,new,1)
code=code.replace('mError, lambdaError, lambdaBound, candidateLambda, eventHalfLambda, eventError, Abs(CLng(eventSolved))), _',
    'mError, lambdaError, lambdaBound, candidateLambda, eventHalfLambda, eventError, Abs(CLng(eventSolved)), eventLambda, selectedKind, mApproaches.Count), _',1)
old='ByRef actualHalf() As Double, ByRef halfLambda As Double, ByRef error As Double, ByRef eventSolved As Boolean) As Boolean'
new='ByRef actualHalf() As Double, ByRef halfLambda As Double, ByRef error As Double, ByRef eventSolved As Boolean, _\n        ByVal callAllowance As Long, ByRef eventLambda As Double, ByRef selectedKind As Long) As Boolean'
assert code.count(old)==1;code=code.replace(old,new,1)
anchor='    If Not eventSolved Then Exit Function\n';assert code.count(anchor)==1
code=code.replace(anchor,anchor+'    eventLambda = factor\n',1)
old='''    If error > FIELD_ERROR_LIMIT Or Abs(2# * halfLambda - reference - factor) > lambdaBound Or progress <= 0# Then Exit Function
    If factor <= 0# Or halfLambda <= 0# Or 0.5 * spacing * mModulus <= 8# * GeomMax(Bound(factor), Bound(halfLambda)) Then Exit Function
    candidate = face: mError = error: TryEvent = True
'''
new='''    Dim halfYield As Double, beforeSurface As Boolean
    beforeSurface = mSolver.TryReadLoadingSurface(CLng(eventInfo(1)), CLng(eventInfo(2)), half, halfYield)
    If error <= FIELD_ERROR_LIMIT And Abs(2# * halfLambda - reference - factor) <= lambdaBound And progress > 0# And _
            factor > 0# And halfLambda > 0# And 0.5 * spacing * mModulus > 8# * GeomMax(Bound(factor), Bound(halfLambda)) And beforeSurface And halfYield < 0# Then
        candidate = face: mError = error: selectedKind = 2: TryEvent = True
        Exit Function
    End If
    If TryPreEventApproach(accepted, half, actualHalf, halfLambda, reference, projection, coordinate, spacing, _
            arcProjection, q, callAllowance, eventInfo, phaseSteps, candidate, factor) Then
        selectedKind = 3: TryEvent = True
    End If
'''
assert code.count(old)==1;code=code.replace(old,new,1)
code+='\n'+(here/'CandidatePreEventApproach.cls.txt').read_text('utf-8-sig')+'\n'
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
name='src/Solver/CCoupledDeformationContinuation.cls';code=(root/name).read_text('utf-8-sig')
old='                mEnvelopeDepartures = 0 And mDepartureRecoveredPoints = 0 Then '
new='                mEnvelopeDepartures = 0 And mDepartureRecoveredPoints = 0 And mSignedArcRecovered = 0 Then '
assert code.count(old)==1;code=code.replace(old,new,1)
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
method,=re.findall(r'(?ms)^Private Sub TestRunOperationalArc\(\).*?^End Sub',code)
changed=method.replace('ratioData(1), 6, factor, projection)','ratioData(1), 12, factor, projection)',1).replace('Or arc.Calls > 6','Or arc.Calls > 12',1)
changed=changed.replace('magic = "NDMOPERARC000002"','magic = "NDMOPERARC000003"',1)
changed=changed.replace('"OperationalArc.Event", traceField, status(23)','"OperationalArc.Event", traceField, status(27)',1)
anchor='    mReverseSolverTest.DiscardStep\n    If solved Then TestExportSignedArcTrial';assert changed.count(anchor)==1
changed=changed.replace(anchor,'''    Dim approaches As Collection: Set approaches = arc.ApproachVerification
    Dim approachHandle As Integer: approachHandle = FreeFile
    Open mPolishRootTest & "\\Native\\OperationalArc.Approaches.bin" For Binary Access Write As #approachHandle
    Dim approachMagic As String * 16: approachMagic = "NDMARCAPPR000001"
    Put #approachHandle, , approachMagic
    TestWriteReverseBinaryVector approachHandle, Array(approaches.Count)
    Dim record As Variant, vector As Variant
    For Each record In approaches
        For Each vector In record: TestWriteReverseBinaryVector approachHandle, vector: Next vector
    Next record
    Close #approachHandle
    mReverseSolverTest.DiscardStep
    If solved Then TestExportSignedArcTrial''',1)
anchor='    End If\nEnd Sub';at=changed.rfind(anchor);assert at>0
changed=changed[:at]+changed[at:].replace(anchor,'''    End If
    Dim attempt As Long, context As Variant
    For attempt = 1 To approaches.Count
        record = approaches(attempt): context = record(0)
        traceField = record(1): TestExportSignedArcTrial "OperationalArc.ApproachCandidate" & CStr(attempt), traceField, context(2)
        traceField = record(2)
        If UBound(traceField) > 1 Then TestExportSignedArcTrial "OperationalArc.ApproachProbe" & CStr(attempt), traceField, context(3)
    Next attempt
End Sub''',1)
assert changed!=method;code=code.replace(method,changed,1)
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
imports=['CCoupledArcContinuation','CCoupledDeformationContinuation','CCoupledNewtonSolver','modTestCoupledElasticState']
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
values.update({'Compile-TestExternalPoint.ps1':compiler.replace("-ModuleNames 'modTestCoupledElasticState'",'-ModuleNames @('+','.join(repr(n) for n in imports)+')',1).replace('\n','\r\n').encode('utf-8-sig'),
    'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),
    'CandidatePreEventApproach.cls.txt':(here/'CandidatePreEventApproach.cls.txt').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for n,data in values.items():
    if n.endswith(('.cls','.bas')):assert max(map(len,data.decode('utf-8-sig').splitlines()))<1024
    (root/n).write_bytes(data);prep['inputs'][n]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='OperationalBoundedPreEventApproach',actual_import_this_pass=imports,actual_component_count=256,
    original_cold10_and_ratio1_unchanged=True,accepted_field_matches_history_exactly=True,
    at_most4_pre_event_halvings_same_accepted_history=True,event30_primary_and_global96_calls_unchanged=True,
    calls_reserve_exact_candidate_restore=True,field_lambda_force_gauge_yield_and_locality_checked=True,
    event_and_selected_candidate_lambdas_recorded_separately=True,controller_initial32_limit_ends_after_confirmed_arc_recovery=True,
    no_working_promotion=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('793 OPERATIONAL PRE-EVENT HALVING4; SHARED30 EVENT/96 GLOBAL; EXACT SOURCE10;4 IMPORTS/256 COMPONENTS')
