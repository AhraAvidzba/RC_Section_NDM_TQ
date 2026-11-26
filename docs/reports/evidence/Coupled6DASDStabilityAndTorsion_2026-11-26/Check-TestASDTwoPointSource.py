"""ДЛЯ ТЕСТОВ. Проверяет точный изменённый контроллер начального участка ASD и 255 сохранённых компонентов."""
from pathlib import Path

here=Path(__file__).resolve().parent;source=here/'Check-TestASDEarlySource.py';code=source.read_text('utf-8-sig')
code=code.replace('Cold1026ASDEarlyDeformation_20261126','Cold1029ASDTwoPointStart_20261126')
start=code.index("changed={'CCoupledLoadContinuation'");end=code.index('\nassert compiled',start)
code=code[:start]+"changed={'CCoupledLoadContinuation':'src/Solver/CCoupledLoadContinuation.cls'}"+code[end:]
code=code.replace('all253_other_components_unchanged_except_identifier_case=True','all255_other_components_unchanged_except_identifier_case=True')
code=code.replace('three_changed_components_exact=True','changed_components_exact=True,one_updated_controller_exact=True')
code=code.replace('EARLY ASD SOURCE VERIFIED;3 EXACT CHANGED COMPONENTS;253 OTHERS UNCHANGED','TWO-POINT ASD SOURCE VERIFIED;EXACT CONTROLLER;255 OTHER COMPONENTS UNCHANGED')
exec(compile(code,str(source),'exec'),dict(__file__=str(Path(__file__).resolve()),__name__='ASD_two_point_source_test'))
