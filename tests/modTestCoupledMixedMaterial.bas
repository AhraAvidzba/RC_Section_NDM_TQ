Attribute VB_Name = "modTestCoupledMixedMaterial"
Option Explicit

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Выгружает отклик пространственного бетона для независимого сравнения
' с кубом OpenSees ASDConcrete3D при одинаковой истории деформаций.

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Задаёт осевую деформацию и два инженерных сдвига. После нагрузки до
' единицы проходит разгрузку до 0,6 и повторное нагружение до единицы.
Public Function RunTestCoupledMixedMaterial(ByVal stepsText As String) As String
    On Error GoTo Failed
    Dim steps As Long: steps = CLng(stepsText)
    If steps <> 40 And steps <> 160 Then Err.Raise 5, , "Укажите 40 или 160 шагов нагружения."

    Const POISSON As Double = 0.2
    Const PEAK_XZ As Double = 0.00048
    Const PEAK_YZ As Double = 0.00024
    Const ENERGY_ROUNDING As Double = 0.000000000002
    Dim concrete As New CCoupledConcreteResponse, response As ICoupledPointResponse
    concrete.Initialize 1, 30000#, POISSON, 18#, 30#, 3#, 0.002, 0.1, 10#, 100#
    Set response = concrete
    Dim cases As Variant: cases = Array("Compression", "Shear", "Tension")
    Dim axial As Variant: axial = Array(-0.00018, 0#, 0.00009)
    Dim valley As Long: valley = 3 * steps \ 5
    Dim caseIndex As Long, probe As Long, lastProbe As Long, passed As Long, failures As Long, component As Long
    Dim strain(1 To 6) As Double, stress(1 To 6) As Double, unused(1 To 6, 1 To 6) As Double
    Dim factor As Double, et As Double, ec As Double, previousEt As Double, previousEc As Double
    Dim report As String, line As String, status As ESolverFailureCode
    lastProbe = 3 * steps - 2 * valley
    report = "MIXED_MATERIAL;E_MPa=30000;nu=0.2;fc0_MPa=18;fc_MPa=30;ft_MPa=3;Gt_N_per_mm=0.1;Gc_N_per_mm=10;l0_mm=100" & vbCrLf
    report = report & "MIXED_COLUMNS;case;step;factor;xx;yy;zz;xy;xz;yz;tension_energy_ratio;compression_energy_ratio" & vbCrLf
    For caseIndex = LBound(cases) To UBound(cases)
        response.ClearHistory
        previousEt = 0#: previousEc = 0#
        For probe = 1 To lastProbe
            If probe <= steps Then
                factor = CDbl(probe) / steps
            ElseIf probe <= 2 * steps - valley Then
                factor = CDbl(2 * steps - probe) / steps
            Else
                factor = CDbl(probe - 2 * (steps - valley)) / steps
            End If
            strain(3) = axial(caseIndex) * factor
            strain(1) = -POISSON * strain(3): strain(2) = strain(1) ' Свободное поперечное расширение только для упругого эталона.
            strain(5) = PEAK_XZ * factor: strain(6) = PEAK_YZ * factor
            status = response.EvaluatePoint(1, strain, stress, unused, ctmStressOnly)
            If status <> sfcNone Then
                report = report & "FAIL: NUMERICAL_FAILURE " & cases(caseIndex) & " step " & CStr(probe) & " at " & Trim$(Str$(factor)) & vbCrLf
                failures = failures + 1
                Exit For
            End If
            response.AcceptTrial
            concrete.AcceptedPointDamage 1, et, ec
            If et + ENERGY_ROUNDING < previousEt Or ec + ENERGY_ROUNDING < previousEc Then
                report = report & "FAIL: material history decreases at " & cases(caseIndex) & " step " & CStr(probe) & vbCrLf
                failures = failures + 1
            Else
                passed = passed + 1
            End If
            previousEt = et: previousEc = ec
            line = "MIXED_POINT;" & cases(caseIndex) & ";" & CStr(probe) & ";" & Trim$(Str$(factor))
            For component = 1 To 6
                line = line & ";" & Trim$(Str$(stress(component)))
            Next component
            report = report & line & ";" & Trim$(Str$(et)) & ";" & Trim$(Str$(ec)) & vbCrLf
        Next probe
    Next caseIndex
    RunTestCoupledMixedMaterial = report & "TOTAL_COUPLED_MIXED_MATERIAL: passed=" & passed & "; failed=" & failures
    Exit Function
Failed:
    RunTestCoupledMixedMaterial = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function
