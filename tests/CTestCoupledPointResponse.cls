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
Private mCubic As Double, mTwinPeaks As Boolean, mBlocksRefinedDesign As Boolean
Private mCoarseShoulderBump As Boolean

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Задаёт число расчётных точек. Ненагруженное состояние задаёт ClearHistory.
Friend Sub TestInitialize(ByVal pointCount As Long)
    mCount = pointCount
    mCubic = 1E+10#: mTwinPeaks = False: mBlocksRefinedDesign = False
    mCoarseShoulderBump = False
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Задаёт аналитическую ветвь с пиком при осевой деформации 0,001.
Friend Sub TestUseDescendingAxialLaw()
    mCubic = -1E+10#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Задаёт два максимума между крупными шагами деформационного продолжения.
Friend Sub TestUseTwinPeakAxialLaw()
    mTwinPeaks = True
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Добавляет отклик крупного шага, который исчезает при принятии промежуточного поля.
Friend Sub TestUseDiscardedShoulderLaw()
    TestUseTwinPeakAxialLaw
    mCoarseShoulderBump = True
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Вводит гладкий локальный вклад 0,4 в множитель до принятия промежуточного поля.
Private Sub TestAddCoarseShoulder(ByVal strain As Double, ByVal acceptedPeak As Double, _
        ByRef stress As Double, ByRef tangent As Double)
    Const PEAK_STRAIN As Double = 0.0006, HALF_WIDTH As Double = 0.00002
    Const AMPLITUDE As Double = 0.4 * 410000# / 21500#
    If acceptedPeak >= 0.00045 Then Exit Sub
    Dim point As Double: point = (strain - PEAK_STRAIN) / HALF_WIDTH
    If Abs(point) >= 1# Then Exit Sub
    Dim blend As Double: blend = 1# - point * point
    stress = stress + AMPLITUDE * blend * blend
    tangent = tangent - 4# * AMPLITUDE * point * blend / HALF_WIDTH
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает напряжение за вычетом силы стержня на плато 4 кН.
' Суммарный отклик сечения описывает полином по безразмерной деформации.
Private Sub TestTwinPeakStress(ByVal strain As Double, ByRef stress As Double, ByRef tangent As Double)
    Const SCALE_STRAIN As Double = 0.0001, START_STRAIN As Double = 0.0004
    Const NET_AREA As Double = 21500#, DESIGN_FORCE As Double = 410000#
    Const STEEL_FORCE As Double = 4000#
    If strain < START_STRAIN Then
        tangent = (0.75 * DESIGN_FORCE - STEEL_FORCE) / (NET_AREA * START_STRAIN)
        stress = tangent * strain
    Else
        Dim t As Double: t = (strain - START_STRAIN) / SCALE_STRAIN
        Dim factor As Double: factor = (((-218# * t + 1309#) * t - 2428#) * t + 1379#) * t + 450#
        Dim slope As Double: slope = ((-872# * t + 3927#) * t - 4856#) * t + 1379#
        stress = (DESIGN_FORCE * factor / 600# - STEEL_FORCE) / NET_AREA
        tangent = DESIGN_FORCE * slope / (600# * NET_AREA * SCALE_STRAIN)
    End If
End Sub


' -------------------- ДЛЯ ТЕСТОВ --------------------
' Исключает точную силовую точку только после принятия уточнённого шага.
Friend Sub TestBlockRefinedDesign()
    mBlocksRefinedDesign = True
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
    If mTwinPeaks Then
        TestTwinPeakStress e, stress(3), tangent(3, 3)
        If mCoarseShoulderBump Then TestAddCoarseShoulder e, mAcceptedPeak(point), stress(3), tangent(3, 3)
    Else
        stress(3) = stress(3) + mCubic * e ^ 3
        tangent(3, 3) = tangent(3, 3) + 3# * mCubic * e ^ 2
    End If
    Dim peak As Double: peak = mAcceptedPeak(point)
    If Abs(e) > peak Then peak = Abs(e)
    If recordTrial Then mTrialPeak(point) = peak
    If mBlocksRefinedDesign And mAcceptedPeak(point) >= 0.0008 And mAcceptedPeak(point) < 0.00081135 Then
        If Abs(e - 0.0008114) <= 0.00000005 Then ICoupledPointResponse_EvaluatePoint = sfcNumericalFailure
    End If
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
