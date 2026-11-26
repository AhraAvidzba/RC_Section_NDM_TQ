"""ДЛЯ ТЕСТОВ. Проверяет частное исследование допуска и соответствие объявленных параметров реальному коду."""
from pathlib import Path

here=Path(__file__).resolve().parent;source=here/'Check-TestASDEarlySource.py';code=source.read_text('utf-8-sig')
code=code.replace('Cold1026ASDEarlyDeformation_20261126','Cold1034ASDTolerance1E5_20261126')
start=code.index("changed={'CCoupledLoadContinuation'");end=code.index('\nassert compiled',start)
code=code[:start]+"changed={'CCoupledSectionAnalysis':'src/Solver/CCoupledSectionAnalysis.cls','modTestSelectableMaterialCore':'tests/modTestSelectableMaterialCore.bas'}"+code[end:]
code=code.replace('all253_other_components_unchanged_except_identifier_case=True','all254_other_components_unchanged_except_identifier_case=True')
code=code.replace('three_changed_components_exact=True','changed_components_exact=True,two_study_components_exact=True')
code=code.replace('EARLY ASD SOURCE VERIFIED;3 EXACT CHANGED COMPONENTS;253 OTHERS UNCHANGED','ASD TOLERANCE STUDY SOURCE VERIFIED;2 EXACT COMPONENTS;254 OTHERS UNCHANGED')
exec(compile(code,str(source),'exec'),dict(__file__=str(Path(__file__).resolve()),__name__='ASD_tolerance_source_test'))
