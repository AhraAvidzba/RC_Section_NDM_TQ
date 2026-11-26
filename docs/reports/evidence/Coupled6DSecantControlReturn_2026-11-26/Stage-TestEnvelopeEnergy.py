"""ДЛЯ ТЕСТОВ. Меняет только численное управление на среднюю потреблённую энергию существующих огибающих CDP."""

from argparse import ArgumentParser
from hashlib import sha256
import json
from pathlib import Path
import re
from runpy import run_path


def TestStage(repo,root,source,checkpoint,binary,predictor=False,dogleg=False):
    """Использует исходный возврат CDP, три принятые истории и рабочие силовые уравнения; термодинамическая эквивалентность не заявляется."""
    here=Path(__file__).resolve().parent
    run_path(str(here/'Stage-TestCheckpointFieldProgress.py'))['TestStage'](repo,root,source,checkpoint,binary)
    prep=json.loads((root/'Preparation.json').read_bytes())
    name='src/Materials/CCoupledConcreteResponse.cls'
    code=(root/name).read_text('utf-8-sig')+'''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Повторяет существующую пробу от принятой истории и возвращает доли энергии огибающих; историю не записывает.
Friend Function TryReadTrialEnvelopeEnergyTest(ByVal point As Long, ByRef strain() As Double, _
        ByRef tension As Double, ByRef compression As Double) As Boolean
    If Not mASDFlow Is Nothing Then Err.Raise 5, "CCoupledConcreteResponse", "Энергии огибающих тестируются только для CDP."
    Dim updated As TConcreteHistory, stress(1 To 6) As Double, elastic As Boolean, substeps As Long
    If Not UpdateWithSubsteps(strain,mHistory(point,mCommitted),updated,stress,elastic,substeps) Then Exit Function
    tension = updated.TensionEnergy: compression = updated.CompressionEnergy
    TryReadTrialEnvelopeEnergyTest = True
End Function
'''
    (root/name).write_bytes(code.replace('\n','\r\n').encode('utf-8-sig'));prep['inputs'][name]=sha256((root/name).read_bytes()).hexdigest().upper()
    name='src/Solver/CCoupledNewtonSolver.cls';code=(root/name).read_text('utf-8-sig')
    anchor='    mHybridOriginRevision = mAcceptedRevision: ActivateHybridContext\n    BeginNewtonLineTest = True'
    assert code.count(anchor)==1
    (root/name).write_bytes(code.replace('\n','\r\n').encode('utf-8-sig'));prep['inputs'][name]=sha256((root/name).read_bytes()).hexdigest().upper()
    name='tests/CEnvelopeEnergyEquationsTest.cls';(root/name).write_bytes((here/Path(name).name).read_bytes())
    prep['inputs'][name]=sha256((root/name).read_bytes()).hexdigest().upper();prep['modules'].append(Path(name).stem)
    name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
    template=(here/'TestNativeCheckpointSpatialSequence.bas.txt').read_text('utf-8-sig')
    body=template.split('    Dim handle As Integer:')[0].replace('SpatialSequence','EnvelopeEnergy')
    body+='''    Dim handle As Integer: handle = FreeFile
    Open mHardPreparedPathTest & ".csv" For Output As #handle
    Print #handle, "MODE;EnvelopeConsumedEnergyCoordinate;OriginalForceTolerance;CDPOnly;NotTotalThermodynamicDissipation;NotCapacity"
    Dim attempt As Long, factor As Double, currentFactor As Double, fraction As Double, degree As Long
    Dim wrapper As CEnvelopeEnergyEquationsTest, equations As ICoupledEquationSystem, hybrid As CCoupledHybridSolver
    Dim zero() As Double, solution() As Double, residual() As Double, matrix() As Double, field() As Double
    Dim checked() As Double, actual() As Double, seed As Variant, load(1 To 6) As Double
    Dim baseline As Double, increment As Double, target As Double, energy As Double, tolerance As Double
    Dim maximum As Double, started As Double, rowScale As Double, signedAdvance As Double, oldLength As Double, steel As Double
    Dim solved As Boolean, ready As Boolean, verified As Boolean, unchanged As Boolean, forward As Boolean
    Dim revision As Long: revision = solver.AcceptedRevision
    accepted = fields(fields.Count): previous = fields(fields.Count - 1)
    trace.ReadPoint trace.Count,fraction,currentFactor
    rowScale = calibration(0)/calibration(3)
    tolerance = 0.5*calibration(0)*TestSpatialControlResolution(operators,currentFactor,currentFactor,options,calibration(0))
    For attempt = 1 To 4
        If Not solver.BeginNewtonLineTest(mReverseBaseTest,mReverseDirectionTest,projection,0#,accepted,currentFactor) Then _
            Err.Raise 5, , "Не открыт неизменный контекст энергетической пробы."
        Set wrapper = New CEnvelopeEnergyEquationsTest
        wrapper.InitializeTest solver,operators,concrete,sides(1),sides(3), _
            calibration(6)/calibration(8),calibration(7)/calibration(8),rowScale
        baseline = wrapper.AcceptedEnergyTest
        increment = GeomMax(baseline*0.001,calibration(6)/calibration(8)*0.000001)*0.25^(attempt-1)
        target = baseline + increment: wrapper.SetTargetTest target
        Print #handle, "TARGET;" & CStr(attempt) & ";" & TestOnsetNumber(baseline) & ";" & TestOnsetNumber(increment) & ";" & _
            TestOnsetNumber(target) & ";" & TestOnsetNumber(tolerance) & ";" & TestOnsetNumber(rowScale)
        If increment <= 2#*tolerance/rowScale Then
            solver.EndNewtonLineTest: Print #handle, "UNRESOLVED_ENERGY;NotPhysicalFailure": Exit For
        End If
        Set equations = wrapper: Set hybrid = New CCoupledHybridSolver
        ReDim zero(1 To UBound(accepted)+1): started = Timer
        solved = hybrid.TrySolve(equations,zero,CLng((2*(UBound(accepted)+1)+1)*(options(2)+1)+options(2)), _
            CLng(options(2)),tolerance,0.00000000001,0.2,solution,relativeDifferenceStep:=0.000001, _
            automaticScaling:=True,freshJacobian:=True,completeNewtonLineSearch:=True)
        maximum = 0#: ready = False
        If hybrid.HasBestPoint Then
            ready = equations.TryEvaluate(solution,False,residual,matrix)
            If ready Then
                For degree = 1 To UBound(residual): maximum = GeomMax(maximum,Abs(residual(degree))): Next degree
                solver.ReadNewtonLineFieldTest solution,field,factor
                ready = wrapper.TrialEnergyTest(field,energy)
                Print #handle, "FINAL;" & CStr(attempt) & ";" & TestOnsetNumber(factor) & TestHardPreparedVector(field)
            End If
        End If
        Print #handle, "RESULT;" & CStr(attempt) & ";" & CStr(solved) & ";" & CStr(ready) & ";" & TestOnsetNumber(Timer-started) & ";" & _
            CStr(hybrid.Steps) & ";" & CStr(hybrid.FunctionCalls) & ";" & TestOnsetNumber(maximum) & ";" & TestOnsetNumber(energy)
        solver.EndNewtonLineTest
        If solver.AcceptedRevision <> revision Then Err.Raise 5, , "Энергетическая проба изменила принятую историю."
        If solved And ready And maximum <= tolerance Then
            For degree = 1 To 6: load(degree) = mReverseBaseTest(degree)+factor*mReverseDirectionTest(degree): Next degree
            seed = field
            verified = solver.TryStep(load(1),load(2),load(3),load(4),load(5),load(6),options(3),options(4),1,initialField:=seed)
            unchanged = False: forward = False: steel = 0#
            If verified Then
                unchanged = True: solver.CopyTrial checked,actual
                For degree = 1 To UBound(field): If checked(degree) <> field(degree) Then unchanged = False
                Next degree
                forward = solver.SpatialIncrementAdvances(previous,accepted,field,signedAdvance,oldLength)
                steel = reinforced.TrialSteelUtilization
            End If
            Print #handle, "NATIVE;" & CStr(attempt) & ";" & CStr(verified) & ";" & CStr(unchanged) & ";" & CStr(forward) & ";" & TestOnsetNumber(steel)
            If verified And unchanged And forward And factor > 0# And steel <= 1# And energy-baseline > 2#*tolerance/rowScale Then
                solver.AcceptStep: fields.Add field: trace.Add 1#,factor
                TestWriteSpatialSequenceSnapshot fields,trace,projection,target,actual,attempt
                Set mReverseFieldsTest = fields: Set mReverseTraceTest = trace.CopyPrefix(trace.Count)
                TestExportReverseCandidate 999
                Print #handle, "ACCEPTED;" & CStr(attempt) & ";" & CStr(fields.Count) & ";" & TestOnsetNumber(factor)
                Exit For
            End If
            solver.DiscardStep
        End If
    Next attempt
    Print #handle, "END;" & CStr(fields.Count) & ";NotPhysicalFailureOrCapacity"
    Close #handle: mPolishExecutedTest = True
End Sub'''
    if predictor:
        anchor='    Dim revision As Long: revision = solver.AcceptedRevision'
        assert body.count(anchor)==1
        body=body.replace(anchor,'''    Dim revision As Long, previousEnergy As Double, previousFactor As Double, energyStep As Double, ratio As Double
    Dim predicted() As Double, predictedFactor As Double''',1)
        anchor='    For attempt = 1 To 4\n        If Not solver.BeginNewtonLineTest(mReverseBaseTest,mReverseDirectionTest,projection,0#,accepted,currentFactor) Then _'
        assert body.count(anchor)==1
        body=body.replace(anchor,'''    baseline = TestAcceptedEnvelopeEnergy(operators,concrete,sides,calibration)
    trace.ReadPoint trace.Count-1,fraction,previousFactor
    If Not solver.ReplayAcceptedFields(fields,fields.Count-1) Then Err.Raise 5, , "Не восстановлена предыдущая история для прогноза энергии."
    previousEnergy = TestAcceptedEnvelopeEnergy(operators,concrete,sides,calibration)
    If Not solver.ReplayAcceptedFields(fields,fields.Count) Then Err.Raise 5, , "Не восстановлена текущая история после прогноза энергии."
    revision = solver.AcceptedRevision: energyStep = baseline-previousEnergy
    For attempt = 1 To 4
        increment = GeomMax(baseline*0.001,calibration(6)/calibration(8)*0.000001)*0.25^(attempt-1)
        ratio = 1#: If energyStep > 2#*tolerance/rowScale Then ratio = GeomMin(1024#,increment/energyStep)
        predicted = accepted
        For degree = 1 To solver.PhysicalDegreeCount: predicted(degree) = accepted(degree)+ratio*(accepted(degree)-previous(degree)): Next degree
        predictedFactor = currentFactor+ratio*(currentFactor-previousFactor)
        Print #handle, "PREDICTOR;" & CStr(attempt) & ";" & TestOnsetNumber(previousEnergy) & ";" & TestOnsetNumber(energyStep) & ";" & _
            TestOnsetNumber(ratio) & ";" & TestOnsetNumber(predictedFactor)
        If Not solver.BeginNewtonLineTest(mReverseBaseTest,mReverseDirectionTest,projection,0#,predicted,predictedFactor) Then _''',1)
        body+='''

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Суммирует те же принятые энергии огибающих при восстановлении предыдущей и текущей историй.
Private Function TestAcceptedEnvelopeEnergy(ByVal operators As CCoupledElasticState, ByVal central As CCoupledConcreteResponse, _
        ByVal sides As Collection, ByVal calibration As Variant) As Double
    Dim point As Long, side As Long, x As Double, y As Double, area As Double, total As Double
    Dim tension As Double, compression As Double, response As CCoupledConcreteResponse
    For point = 1 To operators.ConcretePointCount
        operators.Point point,x,y,area: total = total+area
        For side = -1 To 1
            If side < 0 Then
                Set response = sides(1)
            ElseIf side > 0 Then
                Set response = sides(3)
            Else
                Set response = central
            End If
            response.AcceptedPointDamage point,tension,compression
            TestAcceptedEnvelopeEnergy = TestAcceptedEnvelopeEnergy + area*(calibration(6)*tension+calibration(7)*compression)/calibration(8)
        Next side
    Next point
    TestAcceptedEnvelopeEnergy = TestAcceptedEnvelopeEnergy/(3#*total)
End Function
'''
    if dogleg:
        anchor='automaticScaling:=True,freshJacobian:=True,completeNewtonLineSearch:=True)'
        assert body.count(anchor)==1
        body=body.replace(anchor,'automaticScaling:=True,freshJacobian:=True,completeNewtonLineSearch:=False)',1)
    code,count=re.subn(r'(?ms)^Public Sub TestRunReversePolishModel\(.*?^End Sub',lambda _:body,code)
    assert count==1 and max(map(len,code.splitlines()))<1024
    (root/name).write_bytes(code.replace('\n','\r\n').encode('utf-8-sig'));prep['inputs'][name]=sha256((root/name).read_bytes()).hexdigest().upper()
    for name in (Path(__file__).name,'CEnvelopeEnergyEquationsTest.cls'):
        (root/name).write_bytes((here/name).read_bytes());prep['inputs'][name]=sha256((root/name).read_bytes()).hexdigest().upper()
    prep['directed_scope'].update(case='EnvelopeEnergyControl',maximum_attempts=4,maximum_new_accepted_steps=1,
        original_material_law_unchanged=True,original_force_tolerance=True,working_sources_changed=False,
        nonlinear_constraint='area-mean of three station consumed uniaxial envelope energies Gt/l*kt + Gc/l*kc',
        full_thermodynamic_dissipation_equivalence_not_claimed=True,CDP_only=True,
        energy_secant_predictor=predictor, dogleg_numerical_corrector=dogleg,
        binary_context_coordinate_is_envelope_energy_not_linear_strain=True,actual_lambda_uncertainty_not_available=True,
        physical_capacity_certified=False)
    (root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
    print('ENVELOPE ENERGY NUMERICAL COORDINATE READY; ORIGINAL PHYSICAL MATERIAL AND FORCE ACCURACY')


if __name__=='__main__':
    p=ArgumentParser(description=__doc__);p.add_argument('--repo',type=Path,default=Path.cwd())
    for name in ('root','source','checkpoint','binary'):p.add_argument('--'+name,type=Path,required=True)
    p.add_argument('--predictor',action='store_true');p.add_argument('--dogleg',action='store_true')
    a=p.parse_args();TestStage(a.repo.resolve(),a.root.resolve(),a.source.resolve(),a.checkpoint.resolve(),a.binary.resolve(),a.predictor,a.dogleg)
