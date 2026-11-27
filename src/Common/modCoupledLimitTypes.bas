Attribute VB_Name = "modCoupledLimitTypes"
Option Explicit

' ==========================================================================
' Виды физических материальных ограничений связанного расчёта
' ==========================================================================
' Это причина события, а не отдельная система пользовательских статусов.
Public Enum ECoupledPhysicalLimitKind
    cplNone = 0
    cplSteel = 1
    cplConcreteCompression = 2
    cplConcreteTension = 3
End Enum

' Уточняет стандартный текст арматурного события для фактически определившего бетонного ограничения.
' Для стали сохраняет переданный текст, включая диагностические форматы.
Public Function CoupledPhysicalLimitText(ByVal text As String, ByVal kind As ECoupledPhysicalLimitKind) As String
    Dim material As String, eventName As String
    Select Case kind
        Case cplConcreteCompression: material = "бетона при сжатии": eventName = "Бетонное событие (сжатие)"
        Case cplConcreteTension: material = "бетона при растяжении": eventName = "Бетонное событие (растяжение)"
        Case Else: CoupledPhysicalLimitText = text: Exit Function
    End Select
    text = Replace(text, "Арматурное событие", eventName)
    text = Replace(text, "Точка арматуры", "Точка бетона")
    CoupledPhysicalLimitText = Replace(text, "арматуры", material)
End Function
