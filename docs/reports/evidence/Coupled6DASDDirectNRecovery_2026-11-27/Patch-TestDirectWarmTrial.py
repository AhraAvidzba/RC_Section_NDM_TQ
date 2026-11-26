"""ДЛЯ ТЕСТОВ. Сначала решает точную нагрузку от фактического левого состояния, сохраняя межузловой прогноз как резерв."""


def TestPatchDirectWarmTrial(code, canonical=False):
    """Повторяет порядок подтверждённой силовой пробы существующего деформационного контроля; допуски и материал сохраняются."""
    before = '                    If TryEquilibrium(1#, False, initial) Then\n'
    after = '''                    Dim exactSolved As Boolean
                    exactSolved = TryEquilibrium(1#, False)
                    If Not exactSolved Then
                        mSolver.DiscardStep
                        exactSolved = TryEquilibrium(1#, False, initial)
                    End If
                    If exactSolved Then
'''
    if canonical:
        before, after = before.lower(), after.lower()
    assert code.count(before) == 1
    return code.replace(before, after, 1)
