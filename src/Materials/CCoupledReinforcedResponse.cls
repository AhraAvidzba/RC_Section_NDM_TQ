VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCoupledReinforcedResponse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Материальный отклик бетона с направленной ненапрягаемой арматурой
' ==========================================================================
' Бетонный закон вычисляется также в точках замещённого объёма.
' В этих точках суммируется отклик стали и вычитается совпадающий
' отклик бетона. История закладывается для каждой точки интегрирования.

Implements ICoupledPointResponse

Private Const COMPONENT_COUNT As Long = 6
Private mConcrete As ICoupledPointResponse
Private mSteel() As CCoupledSteelModel
Private mDirection() As Double
Private mPlastic() As Double, mBackstress() As Double
Private mAxialStress() As Double ' Фактическое продольное напряжение стали до вычитания замещённого бетона.
Private mFirstSteel As Long, mLastSteel As Long, mFirstTransverse As Long
Private mAcceptedBuffer As Long, mTrialBuffer As Long
Private mTrialSteelUtilization As Double
Private mAcceptedSteelUtilization As Double
Private mTrialSteelPoint As Long, mAcceptedSteelPoint As Long
Private mTrialTransversePoint As Long, mAcceptedTransversePoint As Long
Private mTrialTransverseUtilization As Double, mAcceptedTransverseUtilization As Double

' Один раз сопоставляет точкам их диаграммы и проекции.
' Models индексируется классом стали, указанным в сформированной геометрии.
Friend Sub Initialize(ByVal operators As CCoupledElasticState, ByVal concrete As ICoupledPointResponse, _
        ByVal models As Object)
    If operators Is Nothing Or concrete Is Nothing Then _
        Err.Raise 5, "CCoupledReinforcedResponse", "Для связанного отклика требуется подготовленное сечение и бетон."
    Set mConcrete = Nothing
    Erase mSteel: Erase mDirection
    mFirstSteel = operators.ConcretePointCount + 1
    mLastSteel = operators.PointCount
    mFirstTransverse = mLastSteel + 1
    If mLastSteel >= mFirstSteel Then
        If models Is Nothing Then Err.Raise 5, "CCoupledReinforcedResponse", "Не подготовлены диаграммы стали."
        ReDim mSteel(mFirstSteel To mLastSteel)
        ReDim mDirection(1 To COMPONENT_COUNT, mFirstSteel To mLastSteel)
        Dim point As Long, j As Long, projection(1 To COMPONENT_COUNT) As Double, material As String
        For point = mFirstSteel To mLastSteel
            If mFirstTransverse > mLastSteel And operators.TransverseGroupAt(point) > 0 Then mFirstTransverse = point
            material = operators.SteelClassAt(point)
            If Not models.Exists(material) Then Err.Raise vbObjectError + 1781, "CCoupledReinforcedResponse", _
                "Для связанного расчёта нет параметров стали класса " & material & "."
            Set mSteel(point) = models.Item(material)
            operators.SteelProjection point, projection
            For j = 1 To COMPONENT_COUNT
                mDirection(j, point) = projection(j)
            Next j
        Next point
    End If
    Set mConcrete = concrete
    ICoupledPointResponse_ClearHistory
End Sub

' Показывает наличие продольной или поперечной арматуры в балансе сечения.
Friend Property Get HasSteel() As Boolean
    HasSteel = (mFirstSteel > 0 And mLastSteel >= mFirstSteel)
End Property

' Показывает максимальное отношение деформации к пределу на последнем
' полностью вычисленном пробном поле. Первый вызов точки сбрасывает максимум.
Friend Property Get TrialSteelUtilization() As Double
    TrialSteelUtilization = mTrialSteelUtilization
End Property

' Показывает максимум принятого состояния, не изменяемый отклонёнными пробами.
Friend Property Get AcceptedSteelUtilization() As Double
    AcceptedSteelUtilization = mAcceptedSteelUtilization
End Property

' Возвращает точку максимального использования на последней полной пробе.
Friend Property Get TrialSteelCriticalPoint() As Long
    TrialSteelCriticalPoint = mTrialSteelPoint
End Property

' Возвращает контрольную точку принятого состояния по всей арматуре.
Friend Property Get AcceptedSteelCriticalPoint() As Long
    AcceptedSteelCriticalPoint = mAcceptedSteelPoint
End Property

' Возвращает точку максимального использования поперечной арматуры.
' При её отсутствии сохраняется нулевой индекс.
Friend Property Get AcceptedTransverseCriticalPoint() As Long
    AcceptedTransverseCriticalPoint = mAcceptedTransversePoint
End Property

' Возвращает использование по деформации в принятой точке хомута или ветви.
Friend Property Get AcceptedTransverseSteelUtilization() As Double
    AcceptedTransverseSteelUtilization = mAcceptedTransverseUtilization
End Property

' Считывает принятое напряжение стержня материальной точки в МПа.
Friend Property Get AcceptedSteelStress(ByVal point As Long) As Double
    AcceptedSteelStress = mAxialStress(point, mAcceptedBuffer)
End Property

' Обновляет независимую материальную точку. В поправке замещения площадь
' берётся у геометрических операторов и здесь не учитывается повторно.
Private Function ICoupledPointResponse_EvaluatePoint(ByVal point As Long, ByRef strain() As Double, _
        ByRef stress() As Double, ByRef tangent() As Double, _
        Optional ByVal tangentMode As ECoupledTangentMode = ctmFull) As ESolverFailureCode
    If point = 1 Then
        mTrialSteelUtilization = 0#: mTrialSteelPoint = 0
        mTrialTransverseUtilization = 0#: mTrialTransversePoint = 0
    End If
    ICoupledPointResponse_EvaluatePoint = mConcrete.EvaluatePoint(point, strain, stress, tangent, tangentMode)
    If ICoupledPointResponse_EvaluatePoint <> sfcNone Or point < mFirstSteel Then Exit Function
    Dim j As Long, k As Long, axialStrain As Double, axialStress As Double, axialTangent As Double
    Dim utilization As Double
    For j = 1 To COMPONENT_COUNT
        axialStrain = axialStrain + mDirection(j, point) * strain(j)
    Next j
    mSteel(point).EvaluateTrial axialStrain, mPlastic(point, mAcceptedBuffer), mBackstress(point, mAcceptedBuffer), _
        axialStress, axialTangent, mPlastic(point, mTrialBuffer), mBackstress(point, mTrialBuffer)
    mAxialStress(point, mTrialBuffer) = axialStress
    utilization = mSteel(point).StrainUtilization(axialStrain)
    If utilization > mTrialSteelUtilization Then
        mTrialSteelUtilization = utilization: mTrialSteelPoint = point
    End If
    If point >= mFirstTransverse Then
        If utilization > mTrialTransverseUtilization Or mTrialTransversePoint = 0 Then
            mTrialTransverseUtilization = utilization: mTrialTransversePoint = point
        End If
    End If
    For j = 1 To COMPONENT_COUNT
        stress(j) = axialStress * mDirection(j, point) - stress(j)
        If tangentMode <> ctmStressOnly Then
            For k = 1 To COMPONENT_COUNT
                If tangentMode = ctmFull Or k = 3 Or k >= 5 Then _
                    tangent(j, k) = axialTangent * mDirection(j, point) * mDirection(k, point) - tangent(j, k)
            Next k
        End If
    Next j
End Function

' Публикует новый буфер после принятия бетонных проб, без копирования массивов.
Private Sub ICoupledPointResponse_AcceptTrial()
    mConcrete.AcceptTrial
    mAcceptedBuffer = mTrialBuffer
    mTrialBuffer = 3 - mAcceptedBuffer
    mAcceptedSteelUtilization = mTrialSteelUtilization
    mAcceptedSteelPoint = mTrialSteelPoint
    mAcceptedTransversePoint = mTrialTransversePoint
    mAcceptedTransverseUtilization = mTrialTransverseUtilization
End Sub

' Удаляет историю независимого сочетания в обоих материалах.
Private Sub ICoupledPointResponse_ClearHistory()
    mConcrete.ClearHistory
    mAcceptedBuffer = 1: mTrialBuffer = 2
    mTrialSteelUtilization = 0#: mAcceptedSteelUtilization = 0#
    mTrialSteelPoint = 0: mAcceptedSteelPoint = 0
    mTrialTransversePoint = 0: mAcceptedTransversePoint = 0
    mTrialTransverseUtilization = 0#: mAcceptedTransverseUtilization = 0#
    Erase mPlastic: Erase mBackstress: Erase mAxialStress
    If mLastSteel >= mFirstSteel Then
        ReDim mPlastic(mFirstSteel To mLastSteel, 1 To 2)
        ReDim mBackstress(mFirstSteel To mLastSteel, 1 To 2)
        ReDim mAxialStress(mFirstSteel To mLastSteel, 1 To 2)
    End If
End Sub
