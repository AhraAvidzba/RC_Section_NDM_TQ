"""ДЛЯ ТЕСТОВ. Начинает восстановление ASD с двух первых подтверждённых переменных ступеней, сохраняя всю постоянную часть пути."""


def TestPatchDirectRecoveryPrefix(code, canonical=False):
    """Повторяет последний узел раннего префикса точной силовой пробой перед существующим деформационным продолжением."""
    before = '''    Dim history As New Collection, seedTrace As CCoupledLoadTrace, field() As Double, index As Long
    For index = 1 To mAcceptedFields.Count
        field = mAcceptedFields.Item(index): history.Add field
    Next index
    Set seedTrace = mAcceptedTrace.CopyPrefix(history.Count) ' Материальная ветвь восстановления не меняет сохранённый силовой префикс.
    mSolver.DiscardStep
    Dim controlled As New CCoupledDeformationContinuation
    controlled.Trace mSolver, mDirectSteel, mOffset, mDirection, history, PreviousAcceptedLambda, mLambda, _
        mResultants, 2#, 0.001, mMaxIterations, mAbsoluteTolerance, mRelativeTolerance, seedTrace
'''
    after = '''    Dim history As New Collection, seedTrace As CCoupledLoadTrace, field() As Double, index As Long
    Dim prefixCount As Long, variablePoints As Long, fraction As Double, factor As Double
    For index = 1 To mAcceptedTrace.Count
        mAcceptedTrace.ReadPoint index, fraction, factor
        If fraction = 1# And factor > 0# Then variablePoints = variablePoints + 1
        If variablePoints = 2 Then prefixCount = index: Exit For
    Next index
    If prefixCount < 2 Then Exit Function
    For index = 1 To prefixCount
        field = mAcceptedFields.Item(index): history.Add field
    Next index
    Set seedTrace = mAcceptedTrace.CopyPrefix(history.Count) ' Постоянная часть сохраняется полностью, начало переменной ветви не содержит измельчённых отказавших шагов.
    Dim seedPrevious As Double, seedLambda As Double, seedResultants() As Double, seedInitial As Variant
    seedTrace.ReadPoint history.Count - 1, fraction, seedPrevious
    seedTrace.ReadPoint history.Count, fraction, seedLambda
    seedInitial = history.Item(history.Count)
    mSolver.DiscardStep
    If Not mSolver.ReplayAcceptedFields(history, history.Count - 1) Then Set mMeta = failure: Exit Function
    If Not TryEquilibrium(seedLambda, False, seedInitial) Then
        mSolver.DiscardStep: Set mMeta = failure: Exit Function
    End If
    mSolver.CopyTrial field, seedResultants: mSolver.AcceptStep
    history.Remove history.Count: history.Add field ' В начальный узел передаются фактические поле и усилия повторно решённой пробы.
    Dim controlled As New CCoupledDeformationContinuation
    controlled.Trace mSolver, mDirectSteel, mOffset, mDirection, history, seedPrevious, seedLambda, _
        seedResultants, 2#, 0.001, mMaxIterations, mAbsoluteTolerance, mRelativeTolerance, seedTrace
'''
    if canonical:
        before, after = before.lower(), after.lower()
    assert code.count(before) == 1
    return code.replace(before, after, 1)
