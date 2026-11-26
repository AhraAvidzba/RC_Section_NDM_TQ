"""ДЛЯ ТЕСТОВ. Подтверждает точное сохранение всех геометрических операторов при добавлении необязательной станционной проекции стали."""
from pathlib import Path
import re

def TestGeometryZeroOffset(current,legacy):
    """Допускает только заранее заданную добавку side*Bgradient к одному проектору; без side тело и остальной класс точны."""
    before=(legacy/'src/Solver/CCoupledElasticState.cls').read_text('utf-8-sig')
    after=(current/'src/Solver/CCoupledElasticState.cls').read_text('utf-8-sig')
    old='Friend Sub CopySteelStrainProjection(ByVal point As Long, ByRef projection() As Double)'
    new='Friend Sub CopySteelStrainProjection(ByVal point As Long, ByRef projection() As Double, Optional ByVal offset As Double = 0#)'
    assert before.count(old)==1
    expected=before.replace(old,new,1)
    begin=expected.index(new);end=expected.index('End Sub',begin)
    addition='''    If offset <> 0# Then
        projection(mGradientBase + 1) = projection(mGradientBase + 1) + offset * mSteelDirection(3, point)
        projection(mGradientBase + 2) = projection(mGradientBase + 2) + offset * mY(point) * mSteelDirection(3, point)
        projection(mGradientBase + 3) = projection(mGradientBase + 3) + offset * mX(point) * mSteelDirection(3, point)
        projection(mGradientBase + 4) = projection(mGradientBase + 4) + offset * _
            (-mY(point) * mSteelDirection(5, point) + mX(point) * mSteelDirection(6, point))
    End If
'''
    expected=expected[:end]+addition+expected[end:]
    assert expected==after
    return True
