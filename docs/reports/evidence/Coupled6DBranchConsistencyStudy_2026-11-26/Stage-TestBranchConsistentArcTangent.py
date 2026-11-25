"""ДЛЯ ТЕСТОВ. Согласует активные материальные ветви исходящего касательного прогноза с его собственным направлением в общем ядре."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold867FixedMetricKernel_20261126';source=base/'Cold869FixedMetricNextScan_20261126';root=base/'Cold873BranchConsistentArcTangent20_20261126'
audit=json.loads((base/'Cold872PointDerivativeAudit20_20261126/PointDerivativeAuditProof.json').read_bytes());assert audit['native']['failed']==0 and audit['all_working_material_methods_exact867']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0,source/'Input/Reverse1.bin')
archive=root/'BranchConsistentTangentCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
name='src/Solver/CCoupledNewtonSolver.cls';code=(root/name).read_text('utf-8-sig');start=code.index('Friend Function TrySpatialLoadTangent(');end=code.index('\nEnd Function',start)+len('\nEnd Function')
body='''Friend Function TrySpatialLoadTangent(ByRef baseLoad() As Double, ByRef direction() As Double, _
        ByRef previous() As Double, ByRef accepted() As Double, ByVal factor As Double, ByVal previousFactor As Double, _
        ByVal alpha As Double, ByRef tangent() As Double, ByRef probe() As Double, _
        ByVal maximumBuilds As Long, ByRef builds As Long) As Boolean
    builds = 0
    If Not mUseSpatialHistory Or mCandidateReady Or mHybridActive Or maximumBuilds < 1 Or alpha <= 0# Then Exit Function
    Dim revision As Long: revision = mAcceptedRevision
    Dim i As Long, j As Long, iteration As Long, projection() As Double, zero() As Double, residual() As Double, matrix() As Double
    Dim incoming() As Double, coordinate As Double, spacing As Double
    incoming = accepted
    For i = 1 To mDegreeCount: incoming(i) = accepted(i) - previous(i): Next i
    If Not SelectSpatialDirectionControl(previous, accepted, incoming, projection, coordinate, spacing) Then Exit Function
    probe = accepted: ReDim zero(1 To mDegreeCount + 1)
    For i = 1 To mDegreeCount: probe(i) = accepted(i) + incoming(i) / 1024#: Next i
    On Error GoTo Failed
    For iteration = 1 To GeomMin(4, maximumBuilds)
        Dim emptyProjection() As Double: ReDim emptyProjection(1 To mDegreeCount)
        If Not BeginControlLinearization(baseLoad, direction, emptyProjection, 0#, probe, factor) Then Exit Function
        builds = builds + 1
        If Not TryReadBranchLinearization(zero, residual, matrix) Then GoTo Finished
        Dim system As New CCoupledLinearSystem, physical() As Double, rhs() As Double
        ReDim physical(1 To mDegreeCount, 1 To mDegreeCount): ReDim rhs(1 To mDegreeCount)
        For i = 1 To mDegreeCount
            rhs(i) = -matrix(mDegreeCount + 1, i) / mHybridScales(mDegreeCount + 1)
            For j = 1 To mDegreeCount: physical(j, i) = matrix(j, i): Next j
        Next i
        If Not system.Solve(physical, rhs, mDegreeCount) Then GoTo Finished
        ReDim tangent(1 To mDegreeCount)
        For i = 1 To mDegreeCount: tangent(i) = rhs(i) * mHybridScales(i): Next i
        EndControlLinearization
        Dim origin() As Double, length As Double, orientation As Double, scale As Double, nextProbe() As Double
        ReDim origin(1 To mDegreeCount)
        length = SpatialFieldDistance(origin, tangent): orientation = alpha * alpha * (factor - previousFactor)
        For i = 1 To mDegreeCount: orientation = orientation + projection(i) * tangent(i) * spacing: Next i
        If length <= 0# Or orientation = 0# Then GoTo Finished
        scale = Sgn(orientation) * spacing / length: nextProbe = accepted
        For i = 1 To mDegreeCount: nextProbe(i) = accepted(i) + scale * tangent(i) / 1024#: Next i
        Dim beforeSignature() As Long, afterSignature() As Long, same As Boolean
        If Not ReadSpatialBranchSignature(probe, beforeSignature) Then GoTo Finished
        If Not ReadSpatialBranchSignature(nextProbe, afterSignature) Then GoTo Finished
        same = True
        For i = 1 To UBound(beforeSignature)
            If beforeSignature(i) <> afterSignature(i) Then same = False: Exit For
        Next i
        If same Then TrySpatialLoadTangent = True: GoTo Finished
        probe = nextProbe ' Производная пересчитывается только при смене ветви, а не ради уменьшения невязки или изменения истории.
    Next iteration
Finished:
    EndControlLinearization
    If mAcceptedRevision <> revision Then Err.Raise 5, "CCoupledNewtonSolver", "Чтение касательной изменило принятую историю."
    Exit Function
Failed:
    Dim number As Long, description As String: number = Err.Number: description = Err.Description
    EndControlLinearization
    Err.Raise number, "CCoupledNewtonSolver.TrySpatialLoadTangent", description
End Function

' Читает признаки пластического приращения, знака главного напряжения и подшагов на каждой из трёх станций.
' Используется только при согласовании направления; материал и принятая история остаются неизменными.
Private Function ReadSpatialBranchSignature(ByRef field() As Double, ByRef signature() As Long) As Boolean
    ReDim signature(1 To mPointCount * 9)
    Dim side As Long, point As Long, at As Long, response As ICoupledPointResponse
    Dim reinforced As CCoupledReinforcedResponse, concrete As CCoupledConcreteResponse
    Dim strain(1 To 6) As Double, plastic As Boolean, principal As Long, substeps As Long, ready As Boolean
    For side = -1 To 1
        Select Case side
            Case -1: Set response = mMinusResponse
            Case 0: Set response = mResponse
            Case 1: Set response = mPlusResponse
        End Select
        For point = 1 To mPointCount
            YieldStrain point, side, field, strain
            If TypeOf response Is CCoupledReinforcedResponse Then
                Set reinforced = response: ready = reinforced.TryReadConcreteBranch(point, strain, plastic, principal, substeps)
            ElseIf TypeOf response Is CCoupledConcreteResponse Then
                Set concrete = response: ready = concrete.TryReadConcreteBranch(point, strain, plastic, principal, substeps)
            Else
                Exit Function
            End If
            If Not ready Then Exit Function
            signature(at + 1) = Abs(CLng(plastic)): signature(at + 2) = principal: signature(at + 3) = substeps: at = at + 3
        Next point
    Next side
    ReadSpatialBranchSignature = True
End Function'''
code=code[:start]+body+code[end:]
# Описание относится к актуальному контракту, без ссылки на экспериментальную историю.
old="' Пробное поле выбирает активную ветвь материала; оно не является равновесием и не принимается."
assert code.count(old)==1;code=code.replace(old,"' Признаки ветвей проверяются по найденному направлению и при необходимости согласуются до четырёх сборок; равновесие не принимается.",1)
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
name='src/Solver/CCoupledArcContinuation.cls';code=(root/name).read_text('utf-8-sig')
old='''        mCalls = mCalls + 1
        If mSolver.TrySpatialLoadTangent(mBase, mDirection, previous, accepted, factor(4), tangent, probe) Then'''
new='''        Dim tangentBuilds As Long, tangentReady As Boolean
        tangentReady = mSolver.TrySpatialLoadTangent(mBase, mDirection, previous, accepted, factor(4), factor(3), _
            alpha, tangent, probe, callAllowance - 3, tangentBuilds)
        mCalls = mCalls + tangentBuilds ' Каждая сборка касательной включена в общий бюджет; три вызова остаются для пары и восстановления.
        If tangentReady Then'''
assert code.count(old)==1;code=code.replace(old,new,1)
old='mPredictionVerification = Array(Array(1# / 1024#, dl, tangentLength, orientation, alpha), probe, tangent, movement)'
assert code.count(old)==1;code=code.replace(old,'mPredictionVerification = Array(Array(1# / 1024#, dl, tangentLength, orientation, alpha, tangentBuilds, 1#), probe, tangent, movement)',1)
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig');old="-ModuleNames 'modTestCoupledElasticState'";assert compiler.count(old)==1
values['Compile-TestExternalPoint.ps1']=compiler.replace(old,"-ModuleNames @('CCoupledNewtonSolver','CCoupledArcContinuation')",1).replace('\n','\r\n').encode('utf-8-sig')
values.update({'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),
 'SourcePointDerivativeAuditProof.json':(base/'Cold872PointDerivativeAudit20_20261126/PointDerivativeAuditProof.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='SelfConsistentOutgoingTangentFromVerified20',actual_import_this_pass=['CCoupledNewtonSolver','CCoupledArcContinuation'],actual_component_count=256,
 source20_exact=True,unchanged_primary_corrector_from867=True,maximum_tangent_builds4=True,all_builds_counted_in_original96_budget=True,
 original_material_force_and_steel_limits=True,same_existing_public_test_driver=True,no_future_field_used=True,physical_capacity_certified=False,no_release_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'GENERAL OUTGOING BRANCH SIGNATURE FIX;TWO IMPORTS;UNCHANGED PRIMARY CORRECTOR;ALL BUILDS COUNTED')
