VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLoadPathDescriptor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Разбор заданного способа изменения нагрузки
' ==========================================================================
' Преобразует явный выбор LoadPath в зависимость Target(lambda) = Offset + lambda*Base.
' Выбор Auto оставляет решение о траектории вызывающей модели расчёта.
' Отделяет постоянную часть усилий от масштабируемой и сохраняет признаки
' для выбора численного метода, включая осевую силовую траекторию.

Private Const KEY_NONE As String = "None" ' Начальное состояние описателя до назначения нагрузки.
Private Const KEY_AUTO As String = "Auto" ' Общий поиск выбирает и последовательно пробует допустимые способы изменения нагрузки.
Private Const KEY_MX As String = "LambdaMx" ' Масштабируется только пользовательский Mx.
Private Const KEY_MY As String = "LambdaMy" ' Масштабируется только пользовательский My.
Private Const KEY_MXY As String = "LambdaMxy" ' Масштабируется пользовательский вектор Mx/My.
Private Const KEY_N As String = "LambdaN" ' Масштабируется N вместе с моментом от эксцентриситета N.
Private Const KEY_NMXY As String = "LambdaNMxy" ' Масштабируется весь расчётный вектор N/Mx/My.
Private Const KEY_T As String = "LambdaT" ' Масштабируется кручение в пользовательской точке.
Private Const KEY_QX As String = "LambdaQx" ' Масштабируется Qx и её вклад в кручение относительно центра расчёта.
Private Const KEY_QY As String = "LambdaQy" ' Масштабируется Qy и её вклад в кручение относительно центра расчёта.
Private Const KEY_ALL6 As String = "LambdaAll6" ' Масштабируются все шесть усилий.

Private mRawText As String ' Исходная строка из таблицы сочетаний.
Private mKey As String ' Нормализованный внутренний ключ lambda-траектории.
Private mValid As Boolean ' False только при явном неизвестном значении пользователя.
Private mErrorMessage As String ' Человекочитаемая причина InputErr.
Private mHasPath As Boolean ' True, если путь удалось выбрать явно или по умолчанию.
Private mHasScaledLoad As Boolean ' True, если выбранный путь содержит ненулевую масштабируемую компоненту.
Private mForceOnly As Boolean ' True для lambda*N и lambda*NMxy без пользовательских Mx/My.
Private mScalesN As Boolean ' Масштабируется ли продольная сила N.
Private mScalesMx As Boolean ' Масштабируется ли пользовательский Mx.
Private mScalesMy As Boolean ' Масштабируется ли пользовательский My.

Private mPath As CLoadPathVector ' Постоянная и масштабируемая части усилий для зависимости Offset + lambda*Base.

' Задает безопасное начальное состояние: путь отсутствует, объект вектора уже создан.
Private Sub Class_Initialize()
    mValid = True
    mKey = KEY_NONE
    Set mPath = New CLoadPathVector
End Sub

' Разбирает выбранный способ изменения нагрузки через общий объект CSectionLoadState.
' Тот учитывает эксцентриситет продольной силы и фильтр малых моментов,
' а здесь нагрузка только разделяется на постоянную и масштабируемую части.
Public Sub InitializeFromLoadState(ByVal rawPath As String, ByVal loadState As CSectionLoadState)
    If loadState Is Nothing Then _
        Err.Raise vbObjectError + 3821, "CLoadPathDescriptor", "Не переданы продольная сила, изгибающие моменты и точка приложения нагрузки (CSectionLoadState)."

    Clear
    mRawText = Trim$(rawPath)

    Dim normalizedUserMx As Double
    Dim normalizedUserMy As Double
    normalizedUserMx = loadState.LoadPointMx
    normalizedUserMy = loadState.LoadPointMy

    mKey = NormalizeKey(mRawText)

    If Len(mKey) = 0 Then
        If Len(mRawText) > 0 Then
            mValid = False
            mErrorMessage = LoadPathChoices
        Else
            mKey = KEY_NONE
        End If
        Exit Sub
    End If

    mHasPath = (mKey <> KEY_NONE)
    If mKey = KEY_AUTO Then mHasPath = (loadState.N <> 0# Or normalizedUserMx <> 0# Or normalizedUserMy <> 0#)
    If Not mHasPath Then Exit Sub

    BuildFlags loadState, normalizedUserMx, normalizedUserMy
    BuildOffsetBase loadState, normalizedUserMx, normalizedUserMy
End Sub

' Нормализует общий выбор пути без выбора физической траектории Auto.
' Пустой ввод разрешен и означает Auto; неизвестное значение возвращает
' пустой ключ для адресной ошибки вызывающего Excel-слоя.
Public Function NormalizeKey(ByVal rawPath As String) As String
    Dim value As String
    value = LCase$(Trim$(rawPath))
    value = Replace$(value, " ", vbNullString)
    value = Replace$(value, ChrW$(&H3BB), "lambda")

    Select Case value
        Case "lambda*mx", "lambdamx", "mx"
            NormalizeKey = KEY_MX
        Case "lambda*my", "lambdamy", "my"
            NormalizeKey = KEY_MY
        Case "lambda*mxy", "lambdamxy", "mxy"
            NormalizeKey = KEY_MXY
        Case "lambda*n", "lambdan", "nload", "n"
            NormalizeKey = KEY_N
        Case "lambda*nmxy", "lambdanmxy", "nmxy", "all"
            NormalizeKey = KEY_NMXY
        Case "lambda*t", "lambdat"
            NormalizeKey = KEY_T
        Case "lambda*qx", "lambdaqx"
            NormalizeKey = KEY_QX
        Case "lambda*qy", "lambdaqy"
            NormalizeKey = KEY_QY
        Case "lambda*all6", "lambdaall6", "all6"
            NormalizeKey = KEY_ALL6
        Case "auto", vbNullString
            NormalizeKey = KEY_AUTO
    End Select
End Function

' Возвращает понятный список выбора для диагностики ввода из книги и API.
Friend Property Get LoadPathChoices() As String
    Dim multiplier As String: multiplier = ChrW$(&H3BB) & "*"
    LoadPathChoices = "Выберите Auto, " & multiplier & "Mx, " & multiplier & "My, " & multiplier & "Mxy, " & _
        multiplier & "N, " & multiplier & "NMxy, " & multiplier & "T, " & multiplier & "Qx, " & multiplier & "Qy или " & _
        multiplier & "All6. Пустое значение означает Auto."
End Property

' Отличает пути с одной масштабируемой пространственной составляющей.
Friend Function IsSingleSpatialKey(ByVal normalizedKey As String) As Boolean
    IsSingleSpatialKey = (normalizedKey = KEY_T Or normalizedKey = KEY_QX Or normalizedKey = KEY_QY)
End Function

' Оставляет явные плоские пути. При выборе T или Q нормальные трещины
' используют свой Auto по N/Mx/My, а All6 сохраняет пропорцию N/Mx/My.
Friend Function NormalCrackKey(ByVal rawPath As String) As String
    NormalCrackKey = NormalizeKey(rawPath)
    If IsSingleSpatialKey(NormalCrackKey) Then
        NormalCrackKey = KEY_AUTO
    ElseIf NormalCrackKey = KEY_ALL6 Then
        NormalCrackKey = KEY_NMXY
    End If
End Function

' Описывает неприменимость поиска при нулевой выбранной составляющей.
Friend Property Get MissingSpatialLoadReason() As String
    If IsSingleSpatialKey(mKey) And Not mHasScaledLoad Then _
        MissingSpatialLoadReason = "Поиск предела по " & DisplayName & " не выполняется: выбранное усилие в точке приложения равно нулю."
End Property

' Определяет, какие компоненты нагрузки будут масштабироваться. Ненулевая
' компонента сохраняется независимо от величины нужного lambda; дополнительный
' порог здесь не применяется. Нулевые компоненты дают нулевую часть Base
' в той же общей формуле нагрузки.
Private Sub BuildFlags(ByVal loadState As CSectionLoadState, ByVal userMxValue As Double, ByVal userMyValue As Double)
    Dim nValue As Double: nValue = loadState.N
    Select Case mKey
        Case KEY_MX
            mScalesMx = True
            mHasScaledLoad = (userMxValue <> 0#)
        Case KEY_MY
            mScalesMy = True
            mHasScaledLoad = (userMyValue <> 0#)
        Case KEY_MXY
            mScalesMx = True
            mScalesMy = True
            mHasScaledLoad = HasMomentVector(userMxValue, userMyValue)
        Case KEY_N
            mScalesN = True
            mHasScaledLoad = (nValue <> 0#)
        Case KEY_NMXY
            mScalesN = True
            mScalesMx = True
            mScalesMy = True
            mHasScaledLoad = (nValue <> 0# Or HasMomentVector(userMxValue, userMyValue))
        Case KEY_T
            mHasScaledLoad = (loadState.LoadPointTorque <> 0#)
        Case KEY_QX
            mHasScaledLoad = (loadState.Qx <> 0#)
        Case KEY_QY
            mHasScaledLoad = (loadState.Qy <> 0#)
        Case KEY_ALL6
            mScalesN = True: mScalesMx = True: mScalesMy = True
            mHasScaledLoad = (nValue <> 0# Or HasMomentVector(userMxValue, userMyValue) Or loadState.HasSpatialComponents)
        Case KEY_AUTO
            mHasScaledLoad = (nValue <> 0# Or userMxValue <> 0# Or userMyValue <> 0#)
    End Select

    mForceOnly = (mKey = KEY_N Or mKey = KEY_NMXY)
    If mKey = KEY_ALL6 Then mForceOnly = Not loadState.HasSpatialComponents
    mForceOnly = mForceOnly And Not HasMomentVector(userMxValue, userMyValue)
End Sub

' Строит постоянную и масштабируемую части нагрузки.
' Момент от смещения точки приложения N остается связанным с N: если N
' масштабируется, этот момент масштабируется вместе с ней; если масштабируются
' только Mx/My, момент от N остается в Offset.
Private Sub BuildOffsetBase(ByVal loadState As CSectionLoadState, ByVal userMxValue As Double, _
        ByVal userMyValue As Double)
    Dim nValue As Double
    Dim mxFromN As Double
    Dim myFromN As Double
    nValue = loadState.N
    mxFromN = loadState.AxialMxAboutPoint(0#) ' Момент от эксцентриситета N относительно исходного начала модели, Н*мм.
    myFromN = loadState.AxialMyAboutPoint(0#) ' Вторая компонента того же переноса N; остаётся связанной с масштабированием силы.

    Select Case mKey
        Case KEY_MX
            mPath.Initialize nValue, 0#, mxFromN, userMxValue, loadState.InternalMy, 0#
        Case KEY_MY
            mPath.Initialize nValue, 0#, loadState.InternalMx, 0#, myFromN, userMyValue
        Case KEY_MXY
            mPath.Initialize nValue, 0#, mxFromN, userMxValue, myFromN, userMyValue
        Case KEY_N
            mPath.Initialize 0#, nValue, userMxValue, mxFromN, userMyValue, myFromN
        Case KEY_NMXY, KEY_AUTO
            mPath.Initialize 0#, nValue, 0#, loadState.InternalMx, 0#, loadState.InternalMy
        Case KEY_T, KEY_QX, KEY_QY, KEY_ALL6
            Set mPath = SpatialVectorAboutPoint(loadState, 0#, 0#)
    End Select
End Sub

' Разделяет силы в пользовательской точке и переносит Offset и Base
' по отдельности. Поэтому выбор Q меняет кручение от эксцентриситета Q.
Friend Function SpatialVectorAboutPoint(ByVal loadState As CSectionLoadState, _
        ByVal pointX As Double, ByVal pointY As Double) As CLoadPathVector
    Dim result As New CLoadPathVector, target() As Double
    loadState.LoadAboutPoint pointX, pointY, target
    If mKey = KEY_ALL6 Then
        result.Initialize 0#, target(1), 0#, target(2), 0#, target(3), 0#, target(4), 0#, target(5), 0#, target(6)
    Else
        Dim torqueOffset As Double, qxOffset As Double, qyOffset As Double
        Dim torqueBase As Double, qxBase As Double, qyBase As Double
        torqueOffset = loadState.LoadPointTorque: qxOffset = target(5): qyOffset = target(6)
        Select Case mKey
            Case KEY_T
                torqueBase = torqueOffset: torqueOffset = 0#
            Case KEY_QX
                qxBase = qxOffset: qxOffset = 0#
            Case KEY_QY
                qyBase = qyOffset: qyOffset = 0#
            Case Else
                Err.Raise 5, "CLoadPathDescriptor", "Передайте явную пространственную траекторию усилий."
        End Select
        torqueOffset = loadState.TranslateLoadPointTorque(torqueOffset, qxOffset, qyOffset, pointX, pointY)
        torqueBase = loadState.TranslateLoadPointTorque(torqueBase, qxBase, qyBase, pointX, pointY)
        result.Initialize target(1), 0#, target(2), 0#, target(3), 0#, torqueOffset, torqueBase, qxOffset, qxBase, qyOffset, qyBase
    End If
    Set SpatialVectorAboutPoint = result
End Function

' Проверяет наличие ненулевого пользовательского вектора момента после общего фильтра малых моментов.
Private Function HasMomentVector(ByVal mxValue As Double, ByVal myValue As Double) As Boolean
    HasMomentVector = (mxValue <> 0# Or myValue <> 0#)
End Function

' Сбрасывает разобранный путь перед чтением новой строки сочетания нагрузок.
Private Sub Clear()
    mRawText = vbNullString
    mKey = KEY_NONE
    mValid = True
    mErrorMessage = vbNullString
    mHasPath = False
    mHasScaledLoad = False
    mForceOnly = False
    mScalesN = False
    mScalesMx = False
    mScalesMy = False
    Set mPath = New CLoadPathVector
End Sub

' Возвращает название способа изменения нагрузки в виде, принятом на Results.
' Используется для вывода и не меняет расчётную зависимость усилий.
Public Function DisplayNameForKey(ByVal loadPathKey As String) As String
    Select Case loadPathKey
        Case KEY_MX
            DisplayNameForKey = ChrW$(&H3BB) & "*Mx"
        Case KEY_MY
            DisplayNameForKey = ChrW$(&H3BB) & "*My"
        Case KEY_MXY
            DisplayNameForKey = ChrW$(&H3BB) & "*Mxy"
        Case KEY_N
            DisplayNameForKey = ChrW$(&H3BB) & "*N"
        Case KEY_NMXY
            DisplayNameForKey = ChrW$(&H3BB) & "*NMxy"
        Case KEY_T
            DisplayNameForKey = ChrW$(&H3BB) & "*T"
        Case KEY_QX
            DisplayNameForKey = ChrW$(&H3BB) & "*Qx"
        Case KEY_QY
            DisplayNameForKey = ChrW$(&H3BB) & "*Qy"
        Case KEY_ALL6
            DisplayNameForKey = ChrW$(&H3BB) & "*All6"
        Case Else
            DisplayNameForKey = loadPathKey
    End Select
End Function

' Нормализованный внутренний ключ lambda-траектории.
Public Property Get Key() As String
    Key = mKey
End Property

' Возвращает название выбранной траектории нагрузки для вывода пользователю.
Public Property Get DisplayName() As String
    DisplayName = DisplayNameForKey(mKey)
End Property

' False только при явном неизвестном значении пользователя.
Public Property Get Valid() As Boolean
    Valid = mValid
End Property

' Человекочитаемая причина InputErr.
Public Property Get ErrorMessage() As String
    ErrorMessage = mErrorMessage
End Property

' True, если путь удалось выбрать явно или по умолчанию.
Public Property Get HasPath() As Boolean
    HasPath = mHasPath
End Property

' True, если выбранный путь содержит ненулевую масштабируемую компоненту.
Public Property Get HasScaledLoad() As Boolean
    HasScaledLoad = mHasScaledLoad
End Property

' True для осевых путей N, NMxy и All6 при отсутствии переменных моментов и поперечных усилий.
Public Property Get ForceOnly() As Boolean
    ForceOnly = mForceOnly
End Property

' Масштабируется ли продольная сила N.
Public Property Get ScalesN() As Boolean
    ScalesN = mScalesN
End Property

' Масштабируется ли пользовательский Mx.
Public Property Get ScalesMx() As Boolean
    ScalesMx = mScalesMx
End Property

' Масштабируется ли пользовательский My.
Public Property Get ScalesMy() As Boolean
    ScalesMy = mScalesMy
End Property

' Возвращает постоянную часть продольной силы в N = NOffset + lambda*NBase, Н.
Public Property Get NOffset() As Double
    NOffset = mPath.NOffset
End Property

' Возвращает масштабируемую часть продольной силы в N = NOffset + lambda*NBase, Н.
Public Property Get NBase() As Double
    NBase = mPath.NBase
End Property

' Возвращает постоянную часть момента в Mx = MxOffset + lambda*MxBase, Н*мм.
Public Property Get MxOffset() As Double
    MxOffset = mPath.MxOffset
End Property

' Возвращает масштабируемую часть момента в Mx = MxOffset + lambda*MxBase, Н*мм.
Public Property Get MxBase() As Double
    MxBase = mPath.MxBase
End Property

' Возвращает постоянную часть момента в My = MyOffset + lambda*MyBase, Н*мм.
Public Property Get MyOffset() As Double
    MyOffset = mPath.MyOffset
End Property

' Возвращает масштабируемую часть момента в My = MyOffset + lambda*MyBase, Н*мм.
Public Property Get MyBase() As Double
    MyBase = mPath.MyBase
End Property

' Вычисляет N, Mx, My при заданном множителе lambda по формуле Offset + lambda*Base.
' Все сочетания используют этот метод, чтобы одинаково формировать пробную нагрузку.
Public Sub Target(ByVal lambdaValue As Double, ByRef nValue As Double, _
        ByRef mxValue As Double, ByRef myValue As Double)
    mPath.Target lambdaValue, nValue, mxValue, myValue
End Sub

' Возвращает отдельный объект с постоянной и масштабируемой частями нагрузки.
' Изменение пробных точек при поиске не затрагивает исходное описание сочетания.
Public Function Vector() As CLoadPathVector
    Dim result As New CLoadPathVector, offset() As Double, direction() As Double
    mPath.Components offset, direction
    result.Initialize offset(1), direction(1), offset(2), direction(2), offset(3), direction(3), _
        offset(4), direction(4), offset(5), direction(5), offset(6), direction(6)
    Set Vector = result
End Function
