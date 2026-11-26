"""ДЛЯ ТЕСТОВ. Сверяет параметризованный драйвер и все 255 неизменённых компонентов ядра перед повтором."""
from pathlib import Path

here=Path(__file__).resolve().parent;source=here/'Check-TestASDEarlySource.py';code=source.read_text('utf-8-sig')
code=code.replace('Cold1026ASDEarlyDeformation_20261126','Cold1027ASDCapacityScaling_20261126')
start=code.index("changed={'CCoupledLoadContinuation'");end=code.index('\nassert compiled',start)
code=code[:start]+"changed={'modTestSelectableMaterialCore':'tests/modTestSelectableMaterialCore.bas'}"+code[end:]
code=code.replace('all253_other_components_unchanged_except_identifier_case=True','all255_other_components_unchanged_except_identifier_case=True')
code=code.replace('three_changed_components_exact=True','changed_components_exact=True,one_updated_driver_exact=True')
code=code.replace('EARLY ASD SOURCE VERIFIED;3 EXACT CHANGED COMPONENTS;253 OTHERS UNCHANGED','SCALING ASD SOURCE VERIFIED;EXACT DRIVER;255 KERNEL COMPONENTS UNCHANGED')
exec(compile(code,str(source),'exec'),dict(__file__=str(Path(__file__).resolve()),__name__='ASD_capacity_scaling_source_test'))
