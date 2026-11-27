VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCapacityResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Результат расчёта несущей способности
' ==========================================================================
' Принимает итог общего поиска предела и сохраняет инженерные величины.
' Хранит усилия проверенной точки в пользовательской точке приложения.
' Арматурный предел при неопределённой общей несущей способности
' доступен отдельно от предельных усилий и CapacityState.
' Способ поиска и чтение настроек остаются у вызывающего расчёта.

Private mFrozen As Boolean ' После передачи готового результата его заполнение и очистка запрещены; сохранённые числа доступны для чтения.
Private Const CAPACITY_PATH_NONE As String = "None" ' Признак того, что несущая способность для сочетания не рассчитывалась.

Private mSearchResult As CLimitSearchResult ' Типизированный результат общего LimitSearch.
Private mStateResult As CSectionStateResult ' НДС найденного предела или проверенной арматурной точки.
Private mLimitState As String ' Физическая причина предела или текст ошибки.
Private mSolutionMethod As String ' Фактически принятая ветка: UltimateStrain или LoadMultiplier.
Private mPathResolved As String ' Проверенный ключ выбранного изменения нагрузки при поиске предела.
Private mLambdaCapacity As Double ' Найденный множитель lambda.
Private mUtilCapacity As Double ' Использование несущей способности.
Private mNUltimate As Double ' Предельная N в точке приложения нагрузки.
Private mMxUltimate As Double ' Предельный Mx в точке приложения нагрузки.
Private mMyUltimate As Double ' Предельный My в точке приложения нагрузки.
Private mMomentUltimate As Double ' Модуль полного предельного момента в точке приложения нагрузки.
Private mPureAxialForMethod As Boolean ' Признак выбора осевого численного расчёта после переноса нагрузки к центру приведённого сечения.
Private mResultMeta As CResultMeta ' Статус, код причины и пояснение результата несущей способности.
Private mCoupledSearch As Boolean, mIsTransverseLimit As Boolean
Private mCoupledLimitLambda As Double, mConfirmedLowerLambda As Double, mHasConfirmedLowerBound As Boolean
Private mPeakLambda As Double, mPeakAccuracy As Double
Private mTUltimate As Double, mQxUltimate As Double, mQyUltimate As Double

' Задаёт начальный статус неприменимой проверки, чтобы ранний выход
' не оставил незаполненную или ложноположительную причину.
Private Sub Class_Initialize()
    Set mResultMeta = New CResultMeta
    mResultMeta.SetNotApplicable rkCapacity
End Sub

' Добавляет восстановленное предельное НДС либо сохраняет причину, по
' которой поля вывода получить не удалось при подтверждённом пределе.
Friend Sub AttachCoupledState(ByVal state As CSectionStateResult, ByVal warning As String)
    AssertWritable
    Set mStateResult = state
    If Len(warning) = 0 Then Exit Sub
    Dim comment As String
    comment = Trim$(mResultMeta.ResultComment & " " & warning)
    If mResultMeta.InternalStatus = rsSuccess Then
        mResultMeta.SetResult rsSuccessWithWarning, mResultMeta.ResultCode, rkCapacity, comment, mResultMeta.DiagnosticDetails
    Else
        mResultMeta.ResultComment = comment
    End If
End Sub

' Сохраняет подтверждённый арматурный предел. При lambda < 1 критерий
' достаточен для FAIL. При lambda >= 1 полный запас требует проверки
' бетонного механизма и число остаётся множителем арматурного предела.
Friend Sub InitializeCoupledLimit(ByVal lambda As Double, ByRef ultimate() As Double, _
        ByVal pathKey As String, ByVal mechanism As String, ByVal state As CSectionStateResult, _
        Optional ByVal details As String = vbNullString, Optional ByVal transverseLimit As Boolean = False)
    AssertWritable
    If lambda <= 0# Or LBound(ultimate) <> 1 Or UBound(ultimate) <> 6 Then _
        Err.Raise 5, "CCapacityResult", "Передайте шесть усилий и положительный множитель арматурного предела."

    ResetNumericSnapshot
    mCoupledSearch = True: mIsTransverseLimit = transverseLimit
    mCoupledLimitLambda = lambda
    mPathResolved = pathKey
    mSolutionMethod = "CoupledContinuation"
    mLimitState = mechanism
    StoreCoupledVector ultimate
    Set mStateResult = state
    Set mResultMeta = New CResultMeta
    If lambda < 1# Then
        mLambdaCapacity = lambda: mUtilCapacity = 1# / lambda
        mResultMeta.SetResult rsCheckFailed, rcPhysicalLimitExceeded, rkCapacity, _
            mechanism & ". Предел достигнут при множителе " & Format$(lambda, "0.000000") & ".", details
    Else
        mResultMeta.SetResult rsNumericalFailure, rcCoupledConcreteLimitUnresolved, rkCapacity, _
            mechanism & ". Определён арматурный предел. Для полного запаса требуется проверка предела бетонной ветви.", details
    End If
End Sub

' Сохраняет подтверждённый максимум. Проверка у единицы требует
' отделить пик от расчётной нагрузки с учётом оценённой погрешности.
Friend Sub InitializeCoupledPeak(ByVal lambda As Double, ByVal accuracy As Double, _
        ByVal designVerified As Boolean, ByRef ultimate() As Double, ByVal pathKey As String, ByVal details As String)
    AssertWritable
    If lambda <= 0# Or accuracy < 0# Or LBound(ultimate) <> 1 Or UBound(ultimate) <> 6 Then _
        Err.Raise 5, "CCapacityResult", "Передайте шесть усилий и оценку подтверждённого максимума."
    ResetNumericSnapshot
    mCoupledSearch = True: mPeakLambda = lambda: mPeakAccuracy = accuracy
    mPathResolved = pathKey
    mSolutionMethod = "CoupledDeformationControl"
    mLimitState = "Максимум нагрузки на прослеженной ветви"
    StoreCoupledVector ultimate
    Dim reason As String, status As EResultInternalStatus, code As EResultCode
    reason = mLimitState & ". Множитель " & Format$(lambda, "0.000000") & _
        ", оценка погрешности " & Format$(accuracy, "0.000000") & "."
    If lambda + accuracy < 1# And Not designVerified Then
        status = rsCheckFailed: code = rcPhysicalLimitExceeded
    ElseIf lambda - accuracy > 1# And designVerified Then
        status = rsSuccess: code = rcCheckPassed
    Else
        status = rsNumericalFailure: code = rcNumericalFailure
        If Abs(lambda - 1#) <= accuracy Then
            reason = reason & " Погрешность охватывает заданную нагрузку (lambda = 1)."
        Else
            reason = reason & " Положение заданной нагрузки не подтверждено на восходящей ветви."
        End If
    End If
    If status <> rsNumericalFailure Then mLambdaCapacity = lambda: mUtilCapacity = 1# / lambda
    Set mResultMeta = New CResultMeta
    mResultMeta.SetResult status, code, rkCapacity, reason, details
End Sub

' Сохраняет усилия общей предельной или арматурной точки в точке приложения.
Private Sub StoreCoupledVector(ByRef ultimate() As Double)
    mNUltimate = ultimate(1): mMxUltimate = ultimate(2): mMyUltimate = ultimate(3)
    mTUltimate = ultimate(4): mQxUltimate = ultimate(5): mQyUltimate = ultimate(6)
    mMomentUltimate = Sqr(mMxUltimate * mMxUltimate + mMyUltimate * mMyUltimate)
End Sub

' Сохраняет подтверждённое превышение при нагружении постоянной части пути.
Friend Sub InitializeCoupledBaseFail(ByVal mechanism As String, ByVal pathKey As String, _
        Optional ByVal details As String = vbNullString, Optional ByVal transverseLimit As Boolean = False)
    InitializeBaseFail mechanism & ". Предел достигнут при нагружении постоянной части.", pathKey
    mCoupledSearch = True: mIsTransverseLimit = transverseLimit
    mSolutionMethod = "CoupledContinuation"
    mResultMeta.DiagnosticDetails = details
End Sub

' Сохраняет причину без предельной точки. Последний принятый множитель
' является нижней границей прослеженного участка по критерию арматуры.
Friend Sub InitializeCoupledUnresolved(ByVal meta As CResultMeta, ByVal pathKey As String, _
        Optional ByVal lowerLambda As Double = -1#)
    AssertWritable
    ResetNumericSnapshot
    mCoupledSearch = meta.Calculated
    mLimitState = meta.ResultComment
    mPathResolved = pathKey
    mSolutionMethod = "CoupledContinuation"
    Set mResultMeta = New CResultMeta
    mResultMeta.SetResult meta.InternalStatus, meta.ResultCode, rkCapacity, meta.ResultComment, _
        meta.DiagnosticDetails, meta.Applies, meta.Calculated
    mResultMeta.SourceKeys = meta.SourceKeys
    If meta.InternalStatus = rsNumericalFailure And lowerLambda >= 0# Then
        mHasConfirmedLowerBound = True: mConfirmedLowerLambda = lowerLambda
    End If
End Sub

' Сохраняет полный результат успешного или неуспешного поиска.
' При численной ошибке или ошибке ввода результат также содержит статус
' и пояснение; пакетный расчёт не разбирает внутренние поля решателя.
Public Sub InitializeFromSearch(ByVal searchResult As CLimitSearchResult, _
        ByVal loadPath As CLoadPathDescriptor, _
        ByVal pureAxialForMethod As Boolean, _
        ByVal loadReferenceX As Double, ByVal loadReferenceY As Double, _
        ByVal momentFilter As CMomentZeroFilter, ByVal concreteDepthForMx As Double, _
        ByVal concreteDepthForMy As Double)
    AssertWritable
    ResetNumericSnapshot
    If Not searchResult Is Nothing Then Set mSearchResult = searchResult.Clone
    mPureAxialForMethod = pureAxialForMethod
    If loadPath Is Nothing Then
        InitializeInvalidInput "Не подготовлено правило изменения N, Mx и My при поиске предела (LoadPath).", CAPACITY_PATH_NONE
        Exit Sub
    End If

    mPathResolved = loadPath.Key
    If searchResult Is Nothing Then
        InitializeCalcError "Поиск несущей способности не вернул результат.", loadPath.Key
        Exit Sub
    End If

    If searchResult.ResultKind <> rkCapacity Then
        InitializeCalcError "Поиск вернул результат, не относящийся к несущей способности.", loadPath.Key
        Exit Sub
    End If

    Set mResultMeta = searchResult.Meta.Clone
    ' Если поиск подтвердил предел ниже заданной нагрузки, сам поиск успешен,
    ' но текущее сочетание не проходит проверку прочности. История переходов
    ' Auto сохраняется в пользовательском пояснении при итоговом FAIL.
    If searchResult.Succeeded And searchResult.Converged Then
        If searchResult.LambdaUltimate > 0# And searchResult.LambdaUltimate < 1# Then
            mResultMeta.SetResult rsCheckFailed, rcPhysicalLimitExceeded, rkCapacity, _
                Trim$(mResultMeta.ResultComment & " Несущая способность не проходит: предел по выбранному пути достигается при lambda = " & _
                Replace$(Format$(searchResult.LambdaUltimate, "0.000000"), ",", ".") & _
                ", раньше заданной нагрузки (lambda = 1)."), mResultMeta.DiagnosticDetails, _
                mResultMeta.Applies, mResultMeta.Calculated
        End If
    End If
    mLimitState = searchResult.LimitState
    If Len(Trim$(mLimitState)) = 0 Then mLimitState = mResultMeta.ResultComment
    mSolutionMethod = searchResult.ActualMethod
    mLambdaCapacity = searchResult.LambdaUltimate

    StoreUltimateVector loadPath, loadReferenceX, loadReferenceY, momentFilter, _
        concreteDepthForMx, concreteDepthForMy
    StoreUtilization
    StoreCapacityStateIfAvailable
End Sub

' Сохраняет ошибку входных данных, выявленную до запуска поиска.
Public Sub InitializeInvalidInput(ByVal reason As String, _
        Optional ByVal pathKey As String = CAPACITY_PATH_NONE)
    AssertWritable
    If StrComp(pathKey, CAPACITY_PATH_NONE, vbTextCompare) = 0 Then pathKey = vbNullString ' При ошибке разбора пути он не выбран; служебный None не является фактически выполненной траекторией и не должен подменять пустое поле.
    InitializeEarlyResult rsInvalidInput, rcInvalidInput, reason, pathKey
End Sub

' Сохраняет внутреннюю ошибку программы, выявленную до запуска поиска.
Public Sub InitializeCalcError(ByVal reason As String, _
        Optional ByVal pathKey As String = CAPACITY_PATH_NONE)
    AssertWritable
    InitializeEarlyResult rsInternalError, rcInternalError, reason, pathKey
End Sub

' Отмечает запрошенный профиль, которому недоступен выбранный вид предела.
' Траектория поиска и числовой запас остаются пустыми.
Friend Sub InitializeUnsupportedConfiguration(ByVal reason As String)
    AssertWritable
    InitializeEarlyResult rsInvalidConfiguration, rcInvalidConfiguration, reason, vbNullString
End Sub

' Отмечает проверку, выключенную профилем, без найденной точки и без solve.
' Это отличается от физической неприменимости нулевой lambda-траектории.
Public Sub InitializeNotRequested(ByVal reason As String, _
        Optional ByVal pathKey As String = CAPACITY_PATH_NONE)
    AssertWritable
    InitializeEarlyResult rsNotRequested, rcNotRequested, reason, pathKey
End Sub

' Сохраняет незапущенный расчёт без ошибки, если нагрузка не задаёт
' ненулевого направления для поиска предельного множителя lambda.
Public Sub InitializeSkipped(ByVal reason As String, _
        Optional ByVal pathKey As String = CAPACITY_PATH_NONE)
    AssertWritable
    InitializeEarlyResult rsNotApplicable, rcNotApplicable, reason, pathKey
End Sub

' Сохраняет BaseFail, если постоянная часть нагрузки при lambda = 0
' уже превышает критерий несущей способности.
Public Sub InitializeBaseFail(ByVal reason As String, _
        Optional ByVal pathKey As String = CAPACITY_PATH_NONE)
    AssertWritable
    InitializeEarlyResult rsCheckFailed, rcInitialStateBeyondLimit, NormalizeBaseFailReason(reason), pathKey, True
End Sub

' Нормализует объяснение BaseFail в понятную формулировку об исходной части пути.
' Машинный код причины остается в ResultCode = INITIAL_STATE_BEYOND_LIMIT.
Private Function NormalizeBaseFailReason(ByVal reason As String) As String
    NormalizeBaseFailReason = Trim$(reason)
    NormalizeBaseFailReason = Replace$(NormalizeBaseFailReason, _
        "Начальное состояние при lambda = 0 не проходит физический критерий:", _
        "Несущая способность не проходит уже для исходной части выбранного пути:", _
        1, 1, vbTextCompare)
    If Len(NormalizeBaseFailReason) = 0 Then _
        NormalizeBaseFailReason = "Несущая способность не проходит уже для исходной части выбранного пути."
End Function

' Сохраняет ранний результат по внутреннему коду причины,
' не восстанавливая его из текстового обозначения статуса.
Private Sub InitializeEarlyResult(ByVal internalStatus As EResultInternalStatus, _
        ByVal resultCode As EResultCode, ByVal reason As String, _
        ByVal pathKey As String, Optional ByVal calculated As Boolean = False)
    ResetNumericSnapshot
    mLimitState = reason
    mSolutionMethod = vbNullString
    mPathResolved = pathKey
    Set mResultMeta = New CResultMeta
    mResultMeta.SetResult internalStatus, resultCode, rkCapacity, reason, vbNullString, True, calculated

End Sub

' Исключает выдачу точки из другого расчёта при повторном заполнении объекта.
' Отсутствующие числа не заменяются прошлым успешным состоянием.
Private Sub ResetNumericSnapshot()
    Set mSearchResult = Nothing
    Set mStateResult = Nothing
    mLambdaCapacity = 0#
    mUtilCapacity = 0#
    mNUltimate = 0#
    mMxUltimate = 0#
    mMyUltimate = 0#
    mMomentUltimate = 0#
    mPureAxialForMethod = False
    mCoupledSearch = False: mIsTransverseLimit = False
    mCoupledLimitLambda = 0#: mConfirmedLowerLambda = 0#: mHasConfirmedLowerBound = False
    mPeakLambda = 0#: mPeakAccuracy = 0#
    mTUltimate = 0#: mQxUltimate = 0#: mQyUltimate = 0#
End Sub

' Переносит найденный полный предельный вектор в пользовательскую точку
' приложения нагрузки через общий CSectionLoadState.
Private Sub StoreUltimateVector(ByVal loadPath As CLoadPathDescriptor, _
        ByVal loadReferenceX As Double, ByVal loadReferenceY As Double, _
        ByVal momentFilter As CMomentZeroFilter, ByVal concreteDepthForMx As Double, _
        ByVal concreteDepthForMy As Double)
    If mSearchResult Is Nothing Then Exit Sub
    If Not mSearchResult.HasLimitPoint Then Exit Sub

    Dim nValue As Double
    Dim mxAboutOrigin As Double
    Dim myAboutOrigin As Double

    nValue = mSearchResult.NUltimate
    mxAboutOrigin = mSearchResult.MxUltimate
    myAboutOrigin = mSearchResult.MyUltimate

    Dim loadState As CSectionLoadState
    Set loadState = New CSectionLoadState
    loadState.InitializeFromPointMoments nValue, mxAboutOrigin, myAboutOrigin, _
        0#, 0#, loadReferenceX, loadReferenceY, momentFilter, _
        concreteDepthForMx, concreteDepthForMy

    mNUltimate = loadState.N
    mMxUltimate = loadState.LoadPointMx
    mMyUltimate = loadState.LoadPointMy
    mMomentUltimate = Sqr(mMxUltimate * mMxUltimate + mMyUltimate * mMyUltimate)
End Sub

' Определяет использование несущей способности только по подтверждённому пределу.
' Последняя проба и техническая граница поиска пределом не считаются.
Private Sub StoreUtilization()
    mUtilCapacity = 0#
    If mSearchResult Is Nothing Then Exit Sub
    If Not mSearchResult.HasLimitPoint Then Exit Sub
    If mSearchResult.LambdaUltimate > 0# Then mUtilCapacity = 1# / mSearchResult.LambdaUltimate
End Sub

' Готовит CapacityState только для подтверждённой физической предельной точки.
' Неудачные пробы без равновесия не становятся сохранённым НДС для повторного использования.
Private Sub StoreCapacityStateIfAvailable()
    Set mStateResult = Nothing
    If mSearchResult Is Nothing Then Exit Sub
    If Not mSearchResult.HasLimitPoint Then Exit Sub
    Set mStateResult = mSearchResult.PointState
End Sub

' Возвращает отдельную копию результата поиска, чтобы вызывающий код
' не изменил готовый итог. Ссылки на неизменяемые конечные НДС сохраняются.
Public Property Get SearchResult() As CLimitSearchResult
    If Not mSearchResult Is Nothing Then Set SearchResult = mSearchResult.Clone
End Property

' Возвращает независимую копию статуса и пояснения несущей способности.
' Видимое обозначение определяется общими правилами без отдельной строковой копии.
Public Property Get ResultMeta() As CResultMeta
    If mResultMeta Is Nothing Then InitializeCalcError "CapacityResult ResultMeta не сформирован.", mPathResolved
    Set ResultMeta = mResultMeta.Clone
End Property

' Возвращает подробный журнал общего поиска предела.
' Для вывода не требуется читать внутренние рабочие поля CCapacitySolver.
Public Property Get DiagnosticLog() As String
    If Not mSearchResult Is Nothing Then DiagnosticLog = mSearchResult.DiagnosticLog
    If mCoupledSearch Then DiagnosticLog = mResultMeta.DiagnosticDetails
End Property

' Показывает, действительно ли запускался численный поиск LimitSearch.
' Ранние N/A или InputErr являются заполненным результатом, но не означают
' выполненного поиска равновесия или предела.
Public Property Get SearchExecuted() As Boolean
    SearchExecuted = mCoupledSearch
    If Not mSearchResult Is Nothing Then SearchExecuted = mSearchResult.SearchExecuted
End Property

' Возвращает диагностический журнал последнего внутреннего CSectionSolver,
' если предельная lambda-точка действительно была решена равновесием.
Public Property Get LastSolverDiagnosticLog() As String
    If mSearchResult Is Nothing Then Exit Property
    LastSolverDiagnosticLog = mSearchResult.LastSolverDiagnosticLog
End Property

' Объясняет раннюю остановку до LimitSearch для отчёта выполнения.
' Пакетный расчёт выводит готовую причину без повторного её определения.
Public Function EarlyStopReport(ByVal loadPath As CLoadPathDescriptor) As String
    If loadPath Is Nothing Then
        EarlyStopReport = "Поиск несущей способности остановлен: путь нагрузки не подготовлен."
    ElseIf Not loadPath.Valid Then
        EarlyStopReport = "Поиск несущей способности остановлен: неизвестное значение LoadPath."
    ElseIf Not loadPath.HasPath Then
        EarlyStopReport = "Поиск несущей способности не требуется: в сочетании нет ни момента, ни продольной силы."
    ElseIf Not loadPath.HasScaledLoad Then
        EarlyStopReport = "Поиск несущей способности остановлен: выбранный путь не содержит ненулевой масштабируемой нагрузки."
    Else
        EarlyStopReport = "Поиск несущей способности остановлен до численных проб: " & mResultMeta.ResultComment
    End If
End Function

' Объясняет, какой способ поиска несущей способности фактически принят.
' Сам выбор уже выполнен; здесь формируется только текст отчёта.
Public Function MethodChoiceReport(ByVal solutionStrategy As String, _
        ByVal loadPath As CLoadPathDescriptor) As String
    If Len(Trim$(mSolutionMethod)) = 0 Then Exit Function

    If StrComp(solutionStrategy, "Auto", vbTextCompare) = 0 Then
        If StrComp(mSolutionMethod, "UltimateStrain", vbTextCompare) = 0 Then
            MethodChoiceReport = "Итог поиска несущей: Auto принял результат UltimateStrain, резервный переход на LoadMultiplier не потребовался."
        ElseIf StrComp(mSolutionMethod, "LoadMultiplier", vbTextCompare) = 0 Then
            If mPureAxialForMethod Then
                MethodChoiceReport = "Итог поиска несущей: Auto использовал LoadMultiplier сразу, потому что траектория чисто осевая относительно центра приведенного сечения."
            Else
                MethodChoiceReport = "Итог поиска несущей: Auto переключился на LoadMultiplier, потому что UltimateStrain не дал устойчивую предельную точку; подробная причина ниже в диагностике solver-а."
            End If
        End If
    ElseIf StrComp(solutionStrategy, "UltimateStrain", vbTextCompare) = 0 And _
            StrComp(mSolutionMethod, "LoadMultiplier", vbTextCompare) = 0 Then
        If mPureAxialForMethod Then
            MethodChoiceReport = "Итог поиска несущей: фактически использован LoadMultiplier, так как траектория чисто осевая относительно центра приведенного сечения и не решается устойчиво прямым UltimateStrain."
        Else
            MethodChoiceReport = "Итог поиска несущей: фактически использован LoadMultiplier."
            If Not loadPath Is Nothing Then
                If loadPath.ForceOnly Then _
                    MethodChoiceReport = "Итог поиска несущей: UltimateStrain не дал устойчивую предельную точку для силовой траектории с эксцентриситетом; фактически использован LoadMultiplier."
            End If
        End If
    End If
End Function

' Возвращает НДС общего предела для вывода как CapacityState.
Public Property Get StateResult() As CSectionStateResult
    If Not IsSteelLimitOnly Then Set StateResult = mStateResult
End Property

' Возвращает НДС арматурного предела для его самостоятельной диагностики.
Friend Property Get SteelLimitStateResult() As CSectionStateResult
    If HasSteelLimitPoint Then Set SteelLimitStateResult = mStateResult
End Property

' Возвращает внешний статус по сохранённому ResultMeta и общим правилам.
' Физическую причину BaseFail определяет расчёт, а не текст комментария.
Public Property Get Status() As String
    Dim policy As CResultStatusPolicy
    Set policy = New CResultStatusPolicy
    Status = policy.ExternalStatus(ResultMeta)
End Property

' Физическая причина предела или текст ошибки.
Public Property Get LimitState() As String
    LimitState = mLimitState
End Property

' Фактически принятая ветка: UltimateStrain или LoadMultiplier.
Public Property Get SolutionMethod() As String
    SolutionMethod = mSolutionMethod
End Property

' Возвращает проверенный ключ изменения нагрузки при поиске предела.
Public Property Get PathResolved() As String
    PathResolved = mPathResolved
End Property

' Найденный множитель lambda.
Public Property Get LambdaCapacity() As Double
    LambdaCapacity = mLambdaCapacity
End Property

' Сообщает, что сохранены подтверждённые предельные усилия.
Friend Property Get HasUltimatePoint() As Boolean
    If IsSteelLimitOnly Then Exit Property
    HasUltimatePoint = (mCoupledLimitLambda > 0# Or mPeakLambda > 0#)
    If Not mSearchResult Is Nothing Then HasUltimatePoint = mSearchResult.HasLimitPoint
End Property

' Отличает найденное арматурное событие от подтверждённой общей предельной точки.
Private Function IsSteelLimitOnly() As Boolean
    IsSteelLimitOnly = (mResultMeta.ResultCode = rcCoupledConcreteLimitUnresolved)
End Function

' Подтверждает арматурную точку с положительным множителем переменной нагрузки.
Friend Property Get HasSteelLimitPoint() As Boolean
    HasSteelLimitPoint = (mCoupledLimitLambda > 0#)
End Property

' Передаёт шесть усилий арматурной точки в пользовательской точке
' приложения в Н и Н·мм. При отсутствии точки выходной массив пуст.
Friend Function CopySteelLimitVector(ByRef values() As Double) As Boolean
    Erase values
    If Not HasSteelLimitPoint Then Exit Function
    ReDim values(1 To 6)
    values(1) = mNUltimate: values(2) = mMxUltimate: values(3) = mMyUltimate
    values(4) = mTUltimate: values(5) = mQxUltimate: values(6) = mQyUltimate
    CopySteelLimitVector = True
End Function

' Возвращает множитель подтверждённого максимума, если он найден.
Friend Property Get PeakLambda() As Double
    PeakLambda = mPeakLambda
End Property

' Возвращает оценку погрешности максимума по двум разбиениям и невязке.
Friend Property Get PeakAccuracy() As Double
    PeakAccuracy = mPeakAccuracy
End Property

' Отличает поиск пространственной модели от одномерного LimitSearch.
Friend Property Get IsCoupledSearch() As Boolean
    IsCoupledSearch = mCoupledSearch
End Property

' Возвращает множитель именно арматурного предела связанной модели.
Friend Property Get CoupledLimitLambda() As Double
    CoupledLimitLambda = mCoupledLimitLambda
End Property

' Отмечает, что арматурный предел определила поперечная арматура.
Friend Property Get IsTransverseLimit() As Boolean
    IsTransverseLimit = mIsTransverseLimit
End Property

' Статус только арматурного предела. Бетонный механизм оценивается в общем результате.
Friend Property Get TransverseLimitStatus() As String
    If Not mIsTransverseLimit Then Exit Property
    Dim meta As CResultMeta: Set meta = ResultMeta
    If mCoupledLimitLambda >= 1# Then meta.SetResult rsSuccess, rcCheckPassed, rkCapacity
    Dim policy As New CResultStatusPolicy
    TransverseLimitStatus = policy.ExternalStatus(meta)
End Property

' Сообщает о подтверждённой нижней границе пройденного участка.
Friend Property Get HasConfirmedLowerBound() As Boolean
    HasConfirmedLowerBound = mHasConfirmedLowerBound
End Property

' Возвращает нижнюю границу участка без превышения арматурного критерия.
Friend Property Get ConfirmedLowerLambda() As Double
    ConfirmedLowerLambda = mConfirmedLowerLambda
End Property

' Возвращает предельное кручение в пользовательской точке, Н·мм.
Friend Property Get TUltimate() As Double
    If Not IsSteelLimitOnly Then TUltimate = mTUltimate
End Property

' Возвращает предельную поперечную силу X, Н.
Friend Property Get QxUltimate() As Double
    If Not IsSteelLimitOnly Then QxUltimate = mQxUltimate
End Property

' Возвращает предельную поперечную силу Y, Н.
Friend Property Get QyUltimate() As Double
    If Not IsSteelLimitOnly Then QyUltimate = mQyUltimate
End Property

' Использование несущей способности.
Public Property Get UtilCapacity() As Double
    UtilCapacity = mUtilCapacity
End Property

' Предельная N в точке приложения нагрузки.
Public Property Get NUltimate() As Double
    If Not IsSteelLimitOnly Then NUltimate = mNUltimate
End Property

' Предельный Mx в точке приложения нагрузки.
Public Property Get MxUltimate() As Double
    If Not IsSteelLimitOnly Then MxUltimate = mMxUltimate
End Property

' Предельный My в точке приложения нагрузки.
Public Property Get MyUltimate() As Double
    If Not IsSteelLimitOnly Then MyUltimate = mMyUltimate
End Property

' Возвращает внутреннюю продольную силу сохранённого предельного НДС.
' Без состояния возвращает 0; наличие и допустимость проверяются отдельно.
Public Property Get CapacityStateN() As Double
    If CapacityStateAvailable Then CapacityStateN = mStateResult.Nint
End Property

' Возвращает внутренний момент Mx предельного состояния относительно начала модели, Н*мм.
' При отсутствии состояния возвращает 0; его наличие проверяется отдельно.
Public Property Get CapacityStateMx() As Double
    If CapacityStateAvailable Then CapacityStateMx = mStateResult.Mxint
End Property

' Возвращает внутренний момент My предельного состояния относительно начала модели, Н*мм.
' При отсутствии состояния возвращает 0; его наличие проверяется отдельно.
Public Property Get CapacityStateMy() As Double
    If CapacityStateAvailable Then CapacityStateMy = mStateResult.Myint
End Property

' Модуль полного предельного момента в точке приложения нагрузки.
Public Property Get MomentUltimate() As Double
    If Not IsSteelLimitOnly Then MomentUltimate = mMomentUltimate
End Property

' Показывает, сохранено ли предельное напряжённо-деформированное состояние CapacityState.
Public Property Get CapacityStateAvailable() As Boolean
    CapacityStateAvailable = (Not mStateResult Is Nothing) And Not IsSteelLimitOnly
End Property

' Возвращает деформацию в начале координат сохранённого предельного НДС.
' Без состояния возвращает 0; это не подтверждение найденного равновесия.
Public Property Get CapacityEpsilon0() As Double
    If CapacityStateAvailable Then CapacityEpsilon0 = mStateResult.Epsilon0
End Property

' Возвращает сохранённую кривизну kappaX предельного состояния, 1/мм.
' Ноль при отсутствии состояния не подтверждает найденное равновесие.
Public Property Get CapacityKappaX() As Double
    If CapacityStateAvailable Then CapacityKappaX = mStateResult.KappaX
End Property

' Возвращает сохранённую кривизну kappaY предельного состояния, 1/мм.
' Ноль при отсутствии состояния не подтверждает найденное равновесие.
Public Property Get CapacityKappaY() As Double
    If CapacityStateAvailable Then CapacityKappaY = mStateResult.KappaY
End Property

' Показывает, выбран ли осевой численный расчёт после переноса нагрузки
' к центру приведённого сечения.
Public Property Get PureAxialForMethod() As Boolean
    PureAxialForMethod = mPureAxialForMethod
End Property

' Выдает найденный физический множитель как запас. Диагностическая lambda
' незавершенного поиска, BaseFail и отсутствующий предел запасом не являются.
Public Property Get ReserveFactor() As Double
    If mLambdaCapacity <= 0# Then Exit Property
    Dim policy As CResultStatusPolicy
    Set policy = New CResultStatusPolicy
    If Not policy.IsMetaFinished(ResultMeta) Then Exit Property
    ReserveFactor = mLambdaCapacity
End Property

' Завершает заполнение результата и предельного CapacityState, запрещая их изменение.
' Если точка не найдена, НДС не создаётся; повторный вызов безопасен.
Friend Sub Freeze()
    If mFrozen Then Exit Sub
    If Not mStateResult Is Nothing Then mStateResult.Freeze
    mFrozen = True
End Sub

' Запрещает изменение результата после передачи потребителям.
' Для повторного расчёта создаётся отдельный объект; сохранённый результат не изменяется.
Private Sub AssertWritable()
    If mFrozen Then Err.Raise vbObjectError + 4220, "CCapacityResult", _
        "Опубликованный результат нельзя изменить; для нового расчета нужен новый объект."
End Sub

' Сохраняет журнал независимых численных попыток, не меняя физический результат, усилия и поля выбранной попытки.
Friend Sub AppendCoupledSearchDiagnostic(ByVal details As String)
    AssertWritable
    If Not mCoupledSearch Then Err.Raise 5, "CCapacityResult", "Дополнительная диагностика требует результата связанного поиска."
    If Len(details) = 0 Then Exit Sub
    mResultMeta.DiagnosticDetails = mResultMeta.DiagnosticDetails & " | " & details
End Sub
