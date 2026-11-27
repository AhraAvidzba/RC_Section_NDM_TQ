"""ДЛЯ ТЕСТОВ. Проверяет масштаб центральной разности существующего QR-резерва ASD, не меняя силовые допуски."""
from pathlib import Path

here=Path(__file__).resolve().parent
source=here/'Stage-TestASDLinearTrustMy.py'
code=source.read_text('utf-8-sig')
before="root=base/'Cold1154ASDLinearTrustMy_20261127'"
assert code.count(before)==1
code=code.replace(before,"root=base/'Cold1155ASDLinearTrustDifferenceMy_20261127'",1)
anchor="source=source.replace('HasASDGradientControl()','HasASDLinearControl()')"
assert code.count(anchor)==1
code=code.replace(anchor,anchor+'''
before='absoluteTolerance, solution, relativeDifferenceStep:=HYBRID_DIFFERENCE_STEP, initialRadius:=0.2)'
assert source.count(before)==1
source=source.replace(before,'absoluteTolerance, solution, relativeDifferenceStep:=0.001, initialRadius:=0.2)',1)
''',1)
code=code.replace("only_change='Existing QR trust-region least-squares and exact linear mapping also used for ASD axial and mixed field controls'",
                  "only_change='Existing QR trust-region for ASD linear measures; central difference 0.001 in coordinates scaled to 1E-5 strain; no equilibrium tolerance change'",1)
exec(compile(code,str(source),'exec'),dict(__file__=str(source),__name__='TestASDLinearTrustDifferenceStage'))
