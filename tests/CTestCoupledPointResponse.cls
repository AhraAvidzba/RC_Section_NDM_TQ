VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTestCoupledPointResponse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: независимый нелинейный отклик точки для проверки равновесия
' ==========================================================================
' Гладкий закон содержит кубический член и необратимый максимум |eps_zz|.
' Проверяет алгоритм сходимости и публикации; параметры не описывают бетон.

Implements ICoupledPointResponse

Private mCount As Long, mAcceptedPeak() As Double, mTrialPeak() As Double
Private mFail As Boolean, mAcceptedSteps As Long
Private mMaximumIncrement As Double
Private mMinimumAxialStrain As Double
Private mCubic As Double

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Задаёт число расчётных точек. Ненагруженное состояние задаёт ClearHistory.
Friend Sub TestInitialize(ByVal pointCount As Long)
    mCount = pointCount
    mCubic = 1E+10#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Задаёт аналитическую ветвь с пиком при осевой деформации 0,001.
Friend Sub TestUseDescendingAxialLaw()
    mCubic = -1E+10#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Включает контролируемый материальный отказ после записи первого кандидата.
Friend Property Let TestFail(ByVal value As Boolean)
    mFail = value
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Ограничивает допустимую разницу относительно принятого сжатия или растяжения.
Friend Property Let TestMaximumIncrement(ByVal value As Double)
    mMaximumIncrement = value
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' При положительном пороге отклоняет пробы с меньшей осевой деформацией.
Friend Property Let TestMinimumAxialStrain(ByVal value As Double)
    mMinimumAxialStrain = value
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает принятый максимум деформации первой контрольной точки.
Friend Property Get TestAcceptedPeak() As Double
    TestAcceptedPeak = mAcceptedPeak(1)
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает число действительно принятых шагов.
Friend Property Get TestAcceptedSteps() As Long
    TestAcceptedSteps = mAcceptedSteps
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Считает диагональный отклик с нелинейным упрочнением в направлении z.
' Отрицательные параметры для провокации отказа в закон не вводятся.
Private Function ICoupledPointResponse_EvaluatePoint(ByVal point As Long, ByRef strain() As Double, _
        ByRef stress() As Double, ByRef tangent() As Double, _
        Optional ByVal tangentMode As ECoupledTangentMode = ctmFull, _
        Optional ByVal recordTrial As Boolean = True) As ESolverFailureCode
    Const MODULUS As Double = 30000#
    Dim j As Long, k As Long, e As Double
    For j = 1 To 6
        For k = 1 To 6: tangent(j, k) = 0#: Next k
        tangent(j, j) = MODULUS
        If j > 3 Then tangent(j, j) = MODULUS * 0.5
        stress(j) = tangent(j, j) * strain(j)
    Next j
    e = strain(3)
    stress(3) = stress(3) + mCubic * e ^ 3
    tangent(3, 3) = tangent(3, 3) + 3# * mCubic * e ^ 2
    Dim peak As Double: peak = mAcceptedPeak(point)
    If Abs(e) > peak Then peak = Abs(e)
    If recordTrial Then mTrialPeak(point) = peak
    If mFail And point = 2 Then ICoupledPointResponse_EvaluatePoint = sfcNumericalFailure
    If mMaximumIncrement > 0# And peak - mAcceptedPeak(point) > mMaximumIncrement Then _
        ICoupledPointResponse_EvaluatePoint = sfcNumericalFailure
    If mMinimumAxialStrain > 0# And e < mMinimumAxialStrain Then _
        ICoupledPointResponse_EvaluatePoint = sfcNumericalFailure
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Принимает только целый пробный массив и считает события фиксации.
Private Sub ICoupledPointResponse_AcceptTrial()
    mAcceptedPeak = mTrialPeak
    mAcceptedSteps = mAcceptedSteps + 1
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сбрасывает состояние контрольных точек между независимыми сочетаниями.
Private Sub ICoupledPointResponse_ClearHistory()
    ReDim mAcceptedPeak(1 To mCount)
    ReDim mTrialPeak(1 To mCount)
    mAcceptedSteps = 0
End Sub
