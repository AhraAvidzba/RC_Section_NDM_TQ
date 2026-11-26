"""ДЛЯ ТЕСТОВ. Проверяет единичный холодный запуск согласованного восстановления и фактически сохранённую повторно открытую книгу."""
from pathlib import Path
here=Path(__file__).resolve().parent;file=here/'Check-TestStrictGrowthCold96.py';code=file.read_text('utf-8-sig')
for old,new in [('Cold890StrictGrowthCold96_20261126','Cold920AffineRecoveryCold96_20261126'),
 ('Cold888StrictEnvelopeGrowth_20261126','Cold918AffineBranchRecovery_20261126'),
 ('StrictGrowthTrueCold96','AffineRecoveryTrueCold96'),('same_actual_compiled_kernel888','same_actual_compiled_kernel918'),
 ('StrictGrowthCold96Proof.json','AffineRecoveryCold96Proof.json')]:
    assert old in code,old;code=code.replace(old,new)
exec(compile(code,str(file),'exec'),{'__file__':str(file),'__name__':'affine_recovery_true_cold96_test'})
