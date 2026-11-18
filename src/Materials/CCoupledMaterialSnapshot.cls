VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCoupledMaterialSnapshot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Диаграммы материалов принятого пространственного состояния
' ==========================================================================
' Бетон представлен одноосной калибровкой многокомпонентного закона.
' Для каждого применённого класса стали сохраняется своя монотонная ветвь.
' Точки следуют возрастанию параметра нагружения и хранят напряжения в МПа.

Private mSeries As Collection ' Каждая запись: класс стали либо пустота для бетона, деформации, напряжения.
Private mPoints As Long
Private mConcreteLaw As String

' Снимает диаграммы из фактически использованных экземпляров моделей.
' Диаграммы передаются всем сочетаниям одинакового набора как неизменяемые.
Friend Sub Initialize(ByVal concrete As CCoupledConcreteResponse, ByVal modulus As Double, ByVal steel As Object)
    If Not mSeries Is Nothing Then Err.Raise 5, "CCoupledMaterialSnapshot", "Диаграммы уже сохранены."
    Set mSeries = New Collection
    mConcreteLaw = concrete.LawIdentifier
    Dim compressionStrain() As Double, compressionStress() As Double
    Dim tensionStrain() As Double, tensionStress() As Double
    Dim strain() As Double, stress() As Double
    concrete.CopyCalibrationPoints True, modulus, compressionStrain, compressionStress
    concrete.CopyCalibrationPoints False, modulus, tensionStrain, tensionStress
    Dim negativeCount As Long: negativeCount = UBound(compressionStrain)
    Dim point As Long, source As Long, total As Long: total = negativeCount + UBound(tensionStrain) - 1
    ReDim strain(1 To total): ReDim stress(1 To total)
    For point = 1 To negativeCount
        source = negativeCount + 1 - point ' В Rebar и бетоне начало вывода соответствует сжатию.
        strain(point) = compressionStrain(source): stress(point) = compressionStress(source)
    Next point
    For source = 2 To UBound(tensionStrain) ' Обе ветви используют один общий ноль.
        strain(point) = tensionStrain(source): stress(point) = tensionStress(source): point = point + 1
    Next source
    AppendSeries vbNullString, strain, stress

    Dim grade As Variant, model As CCoupledSteelModel
    For Each grade In steel.Keys
        Set model = steel.Item(grade)
        model.CopyCalibrationPoints strain, stress
        AppendSeries CStr(grade), strain, stress
    Next grade
End Sub

' Добавляет копии пар деформация–напряжение в неизменяемый снимок.
Private Sub AppendSeries(ByVal grade As String, ByRef strain() As Double, ByRef stress() As Double)
    mSeries.Add Array(grade, strain, stress)
    mPoints = mPoints + UBound(strain)
End Sub

' Число сохранённых материалов, включая бетон и применённые классы стали.
Friend Property Get SeriesCount() As Long
    SeriesCount = mSeries.Count
End Property

' Обозначает фактический пространственный закон, которому принадлежат кривые.
Friend Property Get ConcreteLawIdentifier() As String
    ConcreteLawIdentifier = mConcreteLaw
End Property

' Возвращает число строк всех кривых для размерности общего массива отчёта.
Friend Property Get PointCount() As Long
    PointCount = mPoints
End Property

' Класс для выбора подписи и ссылки в сохранённом состоянии.
Friend Property Get SteelClass(ByVal index As Long) As String
    Dim record As Variant: record = mSeries.Item(index)
    SteelClass = CStr(record(0))
End Property

' Копирует всю кривую за один вызов для блочной записи.
Friend Sub CopyCurve(ByVal index As Long, ByRef strain() As Double, ByRef stress() As Double)
    Dim record As Variant: record = mSeries.Item(index)
    strain = record(1): stress = record(2)
End Sub
