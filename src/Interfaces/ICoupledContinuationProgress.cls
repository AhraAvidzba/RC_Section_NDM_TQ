VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICoupledContinuationProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Наблюдение принятой траектории без доступа к изменяемой истории решателя
' ==========================================================================
' Массивы полей, журнал и метаданные передаются отдельными копиями.
' Снимок промежуточный: он не назначает физический предел и не является Results.
' При уточнении плеч очередной снимок может заменять хвост предыдущего пути.

' Получает полный текущий принятый путь перед очередной численной пробой.
' Наблюдатель сам выбирает носитель; расчётное ядро не выполняет ввод-вывод.
Public Sub Report(ByVal fields As Collection, ByVal trace As CCoupledLoadTrace, ByVal meta As CResultMeta)
End Sub
