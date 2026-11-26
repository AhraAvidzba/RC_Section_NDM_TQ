"""ДЛЯ ТЕСТОВ. Подключает восстановление точной нагрузки ASD к существующей деформационной ветви без ослабления критериев."""
from pathlib import Path


def TestPatchDirectRecoveryAnalysis(code, canonical=False):
    """Сохраняет силовой путь успешных случаев и разрешает резерв только для калиброванного ASD."""
    before = '''        STRESS_TOLERANCE, RELATIVE_TOLERANCE, captureFields
    If continuation.Advance(1#, response) Then'''
    after = '''        STRESS_TOLERANCE, RELATIVE_TOLERANCE, (captureFields Or mCalibratedASD)
    Dim reached As Boolean: reached = continuation.Advance(1#, response)
    If Not reached And mCalibratedASD Then reached = continuation.TryRecoverDirectByDeformation()
    If reached Then'''
    if canonical:
        before, after = before.lower(), after.lower()
    assert code.count(before) == 1
    return code.replace(before, after, 1)


def TestPatchDirectRecoveryContinuation(code):
    """Добавляет рабочий резерв с повторением материального префикса и проверенной точной силовой пробой."""
    assert 'TryRecoverDirectByDeformation' not in code
    method = (Path(__file__).resolve().parent / 'ASDDirectRecoveryMethod.bas.txt').read_text('utf-8-sig')
    return code.rstrip() + '\n' + method
