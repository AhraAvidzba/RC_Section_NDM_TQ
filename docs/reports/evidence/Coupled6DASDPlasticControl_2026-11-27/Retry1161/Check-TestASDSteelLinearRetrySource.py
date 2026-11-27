"""ДЛЯ ТЕСТОВ: проверяет, что повтор отличается только сохранением копий тестового наблюдателя."""
from pathlib import Path

source=Path(__file__).with_name('Check-TestPlasticSteelControlMySource.py'); code=source.read_text('utf-8-sig')
changes={
    'Cold1159ASDPlasticSteelMy_20261127':'Cold1161ASDSteelLinearRetry_20261127',
    "changed={'CCoupledElasticState','CCoupledNewtonSolver','CCoupledDeformationContinuation','modTestSelectableMaterialCore'}":"changed={'CCoupledMaterialProgressTest'}",
    "+list((root/'tests').glob(name+'.bas'))":"+list((root/'tests').glob(name+'.bas'))+list((root/'tests').glob(name+'.cls'))",
    'all259_other_parent_components_unchanged':'all262_other_parent_components_unchanged',
    '4 EXACT CHANGES;259 OTHER COMPONENTS UNCHANGED':'1 TEST OBSERVER CHANGE;262 OTHER COMPONENTS UNCHANGED',
}
for before,after in changes.items():
    assert code.count(before)==1,before;code=code.replace(before,after,1)
exec(compile(code,str(source),'exec'),{'__file__':str(source),'__name__':'__main__'})
