"""ДЛЯ ТЕСТОВ. После собственной ASD-кривой решает равновесие при точно заданной предельной деформации критического продольного стержня."""
from pathlib import Path

here=Path(__file__).resolve().parent;source=here/'Stage-TestPureBendingPostCurveForce.py';code=source.read_text('utf-8-sig')
code=code.replace('Cold1013PureBendingPostCurveForce_20261126','Cold1015PureBendingSteelLimit_20261126')
begin=code.index("extra='''");end=code.index("'''\nsource=",begin)+3
extra='''    ' -------------------- ДЛЯ ТЕСТОВ --------------------
    ' Контролирует фактическую осевую деформацию продольного стержня на той же станции, где принят наибольший запас.
    Dim point As Long, side As Long, criticalPoint As Long, criticalSide As Long
    Dim strain() As Double, gradient() As Double, epsilon As Double, maximumUsage As Double, usage As Double, targetStrain As Double
    For point = operators.ConcretePointCount + 1 To operators.PointCount
        If operators.TransverseGroupAt(point) = 0 Then
            operators.PointStrain point, field, strain: operators.AxialGradientAtPoint point, field, gradient
            For side = -1 To 1
                epsilon = strain(3) + CDbl(side) * 0.001 * gradient(3)
                If epsilon >= 0# Then usage = epsilon / 0.025 Else usage = -epsilon / 0.018
                If usage > maximumUsage Then
                    maximumUsage = usage: criticalPoint = point: criticalSide = side
                    If epsilon >= 0# Then targetStrain = 0.025 Else targetStrain = -0.018
                End If
            Next side
        End If
    Next point
    If criticalPoint = 0 Or maximumUsage >= 1# Then Err.Raise 5, , "Не найден допустимый исходный стержень для контроля его предельной деформации."
    Dim unitField() As Double: ReDim unitField(1 To operators.DegreeCount)
    ReDim projection(1 To operators.DegreeCount)
    For degree = 1 To operators.DegreeCount
        unitField(degree) = 1#
        operators.PointStrain criticalPoint, unitField, strain: operators.AxialGradientAtPoint criticalPoint, unitField, gradient
        projection(degree) = strain(3) + CDbl(criticalSide) * 0.001 * gradient(3)
        unitField(degree) = 0#
    Next degree
    Dim currentStrain As Double, priorStrain As Double
    For degree = 1 To operators.DegreeCount
        currentStrain = currentStrain + projection(degree) * field(degree)
        priorStrain = priorStrain + projection(degree) * previous(degree)
    Next degree
    predicted = field: guess = factor
    If currentStrain <> priorStrain Then
        multiplier = (targetStrain - currentStrain) / (currentStrain - priorStrain)
        For degree = 1 To operators.DegreeCount: predicted(degree) = field(degree) + multiplier * (field(degree) - previous(degree)): Next degree
        guess = factor + multiplier * (factor - previousFactor)
    End If
    solved = solver.TryDeformationStep(baseLoad, direction, projection, targetStrain, guess, trialFactor, 0.0000005, 0.0000005, 30, initialField:=predicted)
    effort = solver.ReadPolishEffortTest
    handle = FreeFile
    Open folder & "\\Native\\PureBendingSteelLimit.bin" For Binary Access Write As #handle
    magic = "NDMBENDSTEEL0001": Put #handle, , magic
    TestWriteReverseBinaryVector handle, Array(options(2), criticalPoint, criticalSide, targetStrain, currentStrain, factor)
    TestWriteReverseBinaryVector handle, projection
    TestWriteReverseBinaryVector handle, Array(Abs(CLng(solved)), trialFactor, effort(0), effort(1), effort(7))
    If solved Then
        solver.CopyTrial nextField, actual
        TestWriteReverseBinaryVector handle, nextField: TestWriteReverseBinaryVector handle, actual
        TestWriteASDColdStressTest handle, nextField
        solver.AcceptStep
    Else
        solver.DiscardStep
    End If
    Close #handle: handle = 0
'''
code=code[:begin]+"extra="+repr(extra)+code[end:]
code=code.replace("case='PureBendingPostCurveForce'","case='PureBendingSteelLimit'")
code=code.replace("post_curve_force_ratios=[1,1.02,1.05,1.1,1.2,1.3]","steel_strain_control_after_own_curvature_path=True")
code=code.replace('PREPARED1013; EXISTING FORCE SOLVER AFTER THIS RUN CURVATURE PATH; NUMERICAL CLASSES UNCHANGED','PREPARED1015; EXISTING DEFORMATION SOLVER AT LONGITUDINAL STRAIN LIMIT; NUMERICAL CLASSES UNCHANGED')
exec(compile(code,str(source),'exec'),{'__file__':str(Path(__file__).resolve()),'__name__':'pure_bending_steel_limit_stage_test'})
