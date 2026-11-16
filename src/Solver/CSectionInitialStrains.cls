VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSectionInitialStrains"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =====================================================================
' Неизменяемые начальные деформации элементов заданного сечения.
' Добавляются к приращению плоскости при вычислении напряжений.
' Места замены бетона продольными стержнями без начальной деформации бетона.
' =====================================================================

Private mSection As CSectionModel, mRevision As Long
Private mConcrete() As Double, mSteel() As Double

' Закрепляет массивы по порядку элементов одной ревизии модели.
Friend Sub Initialize(ByVal section As CSectionModel, ByRef concrete() As Double, ByRef steel() As Double)
    If Not mSection Is Nothing Or section Is Nothing Then Err.Raise vbObjectError + 4509, TypeName(Me), "Неверная инициализация начальных деформаций."
    If LBound(concrete) <> 1 Or UBound(concrete) <> section.ConcreteCount Or _
            LBound(steel) <> 1 Or UBound(steel) <> section.RebarCount Then _
        Err.Raise vbObjectError + 4510, TypeName(Me), "Число начальных деформаций должно совпадать с элементами сечения."
    mConcrete = concrete: mSteel = steel
    mRevision = section.Revision
    Set mSection = section
End Sub

' Проверяет, что порядок и геометрия элементов сохранились.
Friend Function Matches(ByVal section As CSectionModel) As Boolean
    If section Is Nothing Or mSection Is Nothing Then Exit Function
    Matches = (mSection Is section And mRevision = section.Revision)
End Function

' Копирует данные в подготовленные массивы решателя.
Friend Sub CopyArrays(ByRef concrete() As Double, ByRef steel() As Double)
    concrete = mConcrete: steel = mSteel
End Sub

' Возвращает начальную деформацию бетонного элемента при выводе поля.
Friend Function ConcreteStrain(ByVal index As Long) As Double
    ConcreteStrain = mConcrete(index)
End Function

' Возвращает начальную деформацию продольного стержня при выводе поля.
Friend Function SteelStrain(ByVal index As Long) As Double
    SteelStrain = mSteel(index)
End Function
