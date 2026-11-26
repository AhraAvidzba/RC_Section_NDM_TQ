"""ДЛЯ ТЕСТОВ. Готовит материальный повтор всей новой холодной истории920, если она не совпадает с уже проверенной898."""
from pathlib import Path
here=Path(__file__).resolve().parent;file=here/'Stage-TestStrictGrowthColdStress.py';code=file.read_text('utf-8-sig')
for old,new in [('Cold890StrictGrowthCold96_20261126','Cold920AffineRecoveryCold96_20261126'),
 ('Cold892StrictGrowthColdStress_20261126','Cold921AffineRecoveryColdStress_20261126'),
 ('StrictGrowthCold96Proof.json','AffineRecoveryCold96Proof.json')]:
    assert old in code,old;code=code.replace(old,new)
exec(compile(code,str(file),'exec'),{'__file__':str(file),'__name__':'affine_recovery_cold_stress_staging'})
