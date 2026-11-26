"""ДЛЯ ТЕСТОВ. Сверяет шестиосевой драйвер и наблюдатель LoadPath, сохраняя нативное численное ядро."""
from pathlib import Path

here=Path(__file__).resolve().parent;source=here/'Check-TestASDEarlySource.py';code=source.read_text('utf-8-sig')
code=code.replace('Cold1026ASDEarlyDeformation_20261126','Cold1037ASDNMTCapacity_20261126')
start=code.index("changed={'CCoupledLoadContinuation'");end=code.index('\nassert compiled',start)
code=code[:start]+"changed={'CCoupledSectionAnalysis':'src/Solver/CCoupledSectionAnalysis.cls','modTestSelectableMaterialCore':'tests/modTestSelectableMaterialCore.bas'}"+code[end:]
code=code.replace('all253_other_components_unchanged_except_identifier_case=True','all254_other_components_unchanged_except_identifier_case=True')
code=code.replace('three_changed_components_exact=True','changed_components_exact=True,two_observation_components_exact=True')
code=code.replace('EARLY ASD SOURCE VERIFIED;3 EXACT CHANGED COMPONENTS;253 OTHERS UNCHANGED','ASD NMT SOURCE VERIFIED;2 EXACT OBSERVATION COMPONENTS;254 OTHERS UNCHANGED')
exec(compile(code,str(source),'exec'),dict(__file__=str(Path(__file__).resolve()),__name__='ASD_NMT_source_test'))
