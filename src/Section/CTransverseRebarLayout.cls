VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTransverseRebarLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Фактические группы поперечной арматуры
' ==========================================================================
' Каждая группа задаёт замкнутый хомут либо открытую анкеровочную ветвь
' по общей геометрии сечения. Неизвестная схема и явно заданное отсутствие
' различаются. Диаметр, шаг вдоль элемента и принадлежность бетону проверяются
' до подготовки квадратуры. Анкеровку задаёт пользователь.

Private mPaths As Collection
Private mCount As Long
Private mCapacity As Long
Private mDiameter() As Double, mSpacing() As Double
Private mSteelClass() As String, mSourceReference() As String
Private mExplicitNone As Boolean

' Готовит контейнер неизменяемых геометрических путей и их параметров.
Private Sub Class_Initialize()
    Set mPaths = New Collection
End Sub

' Подтверждает явное отсутствие поперечной арматуры в данном сечении.
Public Sub DeclareNone(Optional ByVal sourceReference As String = vbNullString)
    If mCount <> 0 Then InputError sourceReference, "Уже заданы группы поперечной арматуры."
    mExplicitNone = True
End Sub

' Добавляет фактически заданную группу без автоматического замыкания и дорисовки.
' Диаметр и шаг задаются в мм, путь должен быть завершён и иметь уникальный ID.
Public Sub AddGroup(ByVal path As CGeometryPath, ByVal diameter As Double, _
        ByVal spacing As Double, ByVal steelClass As String, _
        Optional ByVal sourceReference As String = vbNullString)
    If mExplicitNone Then InputError sourceReference, "Для сечения уже задано отсутствие поперечной арматуры."
    If path Is Nothing Then InputError sourceReference, "Не задан путь поперечной арматуры."
    If Not path.IsSealed Then InputError sourceReference, "Сначала завершите путь поперечной арматуры."
    If diameter <= 0# Then InputError sourceReference, "Диаметр поперечной арматуры должен быть положительным."
    If spacing <= 0# Then InputError sourceReference, "Шаг поперечной арматуры должен быть положительным."
    If diameter > spacing Then InputError sourceReference, "Диаметр поперечной арматуры превышает шаг вдоль элемента."
    steelClass = Trim$(steelClass)
    If Len(steelClass) = 0 Then InputError sourceReference, "Не задан класс поперечной арматуры."
    If HasGroup(path.ID) Then InputError sourceReference, "Повторяется ID группы поперечной арматуры " & path.ID & "."
    If mCount = mCapacity Then ReserveGroup
    mPaths.Add path, path.ID
    mCount = mCount + 1
    mDiameter(mCount) = diameter
    mSpacing(mCount) = spacing
    mSteelClass(mCount) = steelClass
    mSourceReference(mCount) = sourceReference
End Sub

' Готовит ось хомута по выбранной достоверной границе. Пустой ID допускается
' для единственного наружного контура; слой задан до поверхности стержня.
Public Sub AddContourGroup(ByVal section As CSectionModel, ByVal groupID As String, ByVal loopID As String, _
        ByVal diameter As Double, ByVal spacing As Double, ByVal cover As Double, ByVal steelClass As String, _
        Optional ByVal sourceReference As String = vbNullString)
    If section Is Nothing Then InputError sourceReference, "Сначала сформируйте бетонное сечение."
    If diameter <= 0# Then InputError sourceReference, "Диаметр поперечной арматуры должен быть положительным."
    If cover < 0# Then InputError sourceReference, "Защитный слой поперечной арматуры должен быть неотрицательным."
    Dim query As New CSectionGeometryQuery, selected As CGeometryPath, axis As CGeometryPath
    On Error GoTo InvalidContour
    Set selected = query.MaterialContourPath(section.Contours, loopID)
    Set axis = query.OffsetPath(groupID, selected, cover + diameter / 2#)
    Dim outline As CConcreteRegion
    Set outline = query.RegionFromPath(axis)
    query.ValidatePreparedBoundary outline
    On Error GoTo 0
    AddGroup axis, diameter, spacing, steelClass, sourceReference
    Exit Sub
InvalidContour:
    InputError sourceReference, "Не удалось построить ось хомута " & groupID & ". " & Err.Description
End Sub

' Готовит запас места для параметров групп без перестройки на каждой строке.
Private Sub ReserveGroup()
    If mCapacity = 0 Then
        mCapacity = 8
    Else
        mCapacity = 2 * mCapacity
    End If
    ReDim Preserve mDiameter(1 To mCapacity)
    ReDim Preserve mSpacing(1 To mCapacity)
    ReDim Preserve mSteelClass(1 To mCapacity)
    ReDim Preserve mSourceReference(1 To mCapacity)
End Sub

' Проверяет уникальность ID по ключам общего контейнера путей.
Private Function HasGroup(ByVal id As String) As Boolean
    Dim existing As CGeometryPath
    On Error Resume Next
    Set existing = mPaths.Item(id)
    On Error GoTo 0
    HasGroup = Not existing Is Nothing
End Function

' Истина после явного ввода отсутствия или хотя бы одной фактической группы.
Public Property Get IsSpecified() As Boolean
    IsSpecified = mExplicitNone Or mCount > 0
End Property

' Число фактических групп, ноль также возможен при ещё не заполненном вводе.
Public Property Get Count() As Long
    Count = mCount
End Property

' Возвращает неизменяемый путь указанной группы в координатах сечения.
Public Property Get Path(ByVal index As Long) As CGeometryPath
    Set Path = mPaths.Item(index)
End Property

' Диаметр указанной группы, мм.
Public Property Get Diameter(ByVal index As Long) As Double
    Diameter = mDiameter(index)
End Property

' Фактический шаг группы по продольной оси элемента, мм.
Public Property Get Spacing(ByVal index As Long) As Double
    Spacing = mSpacing(index)
End Property

' Материал группы из пользовательской таблицы или каталога материалов.
Public Property Get SteelClass(ByVal index As Long) As String
    SteelClass = mSteelClass(index)
End Property

' Ссылка на строку исходной группы для диагностики.
Public Property Get SourceReference(ByVal index As Long) As String
    SourceReference = mSourceReference(index)
End Property

' Передаёт причину ошибки и адрес, если вызывающий читатель его предоставил.
Private Sub InputError(ByVal sourceReference As String, ByVal reason As String)
    If Len(sourceReference) > 0 Then reason = reason & " " & sourceReference
    Err.Raise vbObjectError + 1777, "CTransverseRebarLayout", reason
End Sub
