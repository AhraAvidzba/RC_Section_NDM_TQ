VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTorsionPlanarResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =====================================================================
' Предварительная проверка замкнутого хомута и нормальных напряжений
' при заданном T и штатном N, Mx, My. В равновесии складываются деформации.
' Предел сечения по кручению и диагональное разрушение бетона сюда не входят.
' =====================================================================

Private mTorque As Double, mFlow As Double, mHoopStress As Double, mHoopReserve As Double
Private mLongitudinalForce As Double, mNormalUtilization As Double
Private mHoopMeta As CResultMeta, mNormalMeta As CResultMeta
Private mSteelTable As Variant ' ID, начальное напряжение, начальная деформация и итоговое напряжение.

' Регистрирует сочетание с ненулевым T и начальные состояния проверок.
Friend Sub Initialize(ByVal torque As Double)
    mTorque = torque
    Set mHoopMeta = New CResultMeta
    Set mNormalMeta = New CResultMeta
    mHoopMeta.SetNotRequested rkGeneric
    mNormalMeta.SetNotRequested rkGeneric
End Sub

' Сохраняет ошибку исходных параметров или построения оси хомута.
Friend Sub SetInputError(ByVal details As String)
    mHoopMeta.SetResult rsInvalidInput, rcInvalidInput, rkGeneric, details
    mNormalMeta.SetResult rsBlockedByDependency, rcBlockedByDependency, rkGeneric, "Нет проверенного контура хомута."
End Sub

' Проверяет отдельное напряжение хомута по заданному сопротивлению.
Friend Sub CheckHoop(ByVal prepared As CTorsionPlanarPreparation, ByRef longitudinalForce As Double)
    prepared.Evaluate mTorque, mFlow, mHoopStress, mHoopReserve, mLongitudinalForce
    longitudinalForce = mLongitudinalForce
    If mHoopReserve < 1# Then
        mHoopMeta.SetResult rsCheckFailed, rcCheckFailed, rkGeneric, "Хомута недостаточно на заданное кручение."
    Else
        mHoopMeta.SetResult rsSuccess, rcCheckPassed, rkGeneric, "Напряжение хомута от T не превышает его расчётное сопротивление."
    End If
End Sub

' Сохраняет машинную причину отсутствия исходных деформаций.
Friend Sub SetFieldUnavailable(ByVal details As String, ByVal physicalLimit As Boolean)
    If physicalLimit Then
        mNormalMeta.SetResult rsCheckFailed, rcPhysicalLimitExceeded, rkGeneric, details
    Else
        mNormalMeta.SetResult rsInvalidInput, rcInvalidInput, rkGeneric, details
    End If
End Sub

' Записывает продольные начальные деформации и действительные итоговые
' напряжения из решения на N, Mx, My с исходным полем от T.
Friend Sub CheckNormalState(ByVal section As CSectionModel, ByVal normalState As CSectionStateResult)
    If normalState Is Nothing Then
        MissingNormalState
        Exit Sub
    End If
    If Not normalState.Converged Then
        Set mNormalMeta = normalState.ResultMeta
        Exit Sub
    End If
    Dim field As CSectionInitialStrains: Set field = normalState.InitialStrainField
    If field Is Nothing Or Not normalState.HasMaterialSnapshot Then
        MissingNormalState
        Exit Sub
    End If

    Dim strain0 As Double, kx As Double, ky As Double
    strain0 = normalState.Epsilon0: kx = normalState.KappaX: ky = normalState.KappaY
    Dim steel As CMaterialDiagram, concrete As CMaterialDiagram
    Set steel = normalState.SteelDiagramSnapshot: Set concrete = normalState.ConcreteDiagramSnapshot
    Dim i As Long, initialStrain As Double, totalStrain As Double
    Dim out() As Variant: ReDim out(1 To section.RebarCount, 1 To 4)
    For i = 1 To section.RebarCount
        initialStrain = field.SteelStrain(i)
        totalStrain = strain0 + kx * section.RebarY(i) + ky * section.RebarX(i) + initialStrain
        out(i, 1) = section.RebarID(i): out(i, 2) = steel.GetStress(initialStrain)
        out(i, 3) = initialStrain: out(i, 4) = steel.GetStress(totalStrain)
    Next i
    mSteelTable = out
    mNormalUtilization = GeomMax(StrainUtilization(concrete, normalState.MinConcreteStrain, normalState.MaxConcreteStrain), _
        StrainUtilization(steel, normalState.MinSteelStrain, normalState.MaxSteelStrain))
    If Not normalState.WithinPhysicalRange Or normalState.ExtensionUsed Then
        Set mNormalMeta = normalState.ResultMeta
    Else
        mNormalMeta.SetResult rsSuccessWithWarning, rcCheckPassed, rkGeneric, _
            "Нормальные деформации проверены при начальном поле от T. Диагональное разрушение бетона и предел всего сечения по T здесь не оцениваются."
    End If
End Sub

' Определяет максимум долей сжатого и активного растянутого пределов.
Private Function StrainUtilization(ByVal diagram As CMaterialDiagram, ByVal minimum As Double, ByVal maximum As Double) As Double
    If diagram.PhysicalCompressionStrain < 0# And minimum < 0# Then StrainUtilization = minimum / diagram.PhysicalCompressionStrain
    If diagram.HasPhysicalTensionLimit And maximum > 0# Then _
        StrainUtilization = GeomMax(StrainUtilization, maximum / diagram.PhysicalTensionStrain)
End Function

' Сохраняет причину отсутствия проверки суммы напряжений.
Private Sub MissingNormalState()
    mNormalMeta.SetResult rsBlockedByDependency, rcBlockedByDependency, rkGeneric, "Нет подтверждённого плоского НДС по N, Mx, My для сложения напряжений."
End Sub

' Возвращает знаковый исходный момент в Н*мм.
Public Property Get Torque() As Double
    Torque = mTorque
End Property

' Возвращает модуль потока касательных усилий в Н/мм.
Public Property Get ShearFlow() As Double
    ShearFlow = mFlow
End Property

' Возвращает оценённое напряжение хомута в МПа.
Public Property Get HoopStress() As Double
    HoopStress = mHoopStress
End Property

' Возвращает отношение допускаемого по хомуту T к заданному |T|.
Public Property Get HoopReserve() As Double
    HoopReserve = mHoopReserve
End Property

' Возвращает максимальное использование по сумме нормальных напряжений.
Public Property Get NormalUtilization() As Double
    NormalUtilization = mNormalUtilization
End Property

' Показывает, получено ли достоверное поле для числовой проверки деформаций.
Friend Property Get HasNormalState() As Boolean
    HasNormalState = IsArray(mSteelTable)
End Property

' Возвращает подробную таблицу продольных стержней во внутренних единицах.
Public Property Get SteelTable() As Variant
    SteelTable = mSteelTable
End Property

' Возвращает независимые статус и пояснение проверки хомута.
Public Property Get HoopMeta() As CResultMeta
    Set HoopMeta = mHoopMeta.Clone
End Property

' Возвращает статус экспериментальной проверки сложенных напряжений.
Public Property Get NormalMeta() As CResultMeta
    Set NormalMeta = mNormalMeta.Clone
End Property
