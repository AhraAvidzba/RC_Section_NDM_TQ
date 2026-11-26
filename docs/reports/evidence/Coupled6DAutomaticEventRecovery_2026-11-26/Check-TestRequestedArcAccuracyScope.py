"""ДЛЯ ТЕСТОВ. Подтверждает единственное изменение масштаба внутренней точности964, не скрывая отличие от963."""
from pathlib import Path
import json

base=Path('D:/CodexScratch/RC_Section_NDM_TQ');root=base/'Cold964RequestedArcAccuracy_20261126';parent=base/'Cold963IncomingEventBudget_20261126'
name='src/Solver/CCoupledArcContinuation.cls';before=(parent/name).read_text('utf-8-sig');after=(root/name).read_text('utf-8-sig')
old='Private Const ACCURACY_SCALE As Double = 0.0001';new='Private Const ACCURACY_SCALE As Double = 1#'
old_comment="' Сохраняет подготовленное ядро и исходную траекторию; внутренние пробы используют более строгую точность без изменения материала."
new_comment="' Сохраняет подготовленное ядро и исходную траекторию; внутренние пробы используют исходную заданную точность без изменения материала."
assert after.count(new)==1 and after.count(new_comment)==1
assert after.replace(new,old,1).replace(new_comment,old_comment,1)==before
for path in (parent/'src').rglob('*'):
    if path.is_file() and path.name!='CCoupledArcContinuation.cls':assert path.read_bytes()==(root/path.relative_to(parent)).read_bytes()
for name in ('tests/modTestCoupledElasticState.bas','Input/Reverse1.bin','Input/OperationalArcRatio.bin','Input/ArcCallAllowance.bin'):
    assert (root/name).read_bytes()==(parent/name).read_bytes()
audit=json.loads((root/'PointTraceSourceComposition.json').read_bytes())
assert audit['components']==257 and audit['imported_this_pass']==['CCoupledArcContinuation']
assert len(audit['other_actual_components_exact_parent_except_identifier_case'])==256
proof=json.loads((root/'CurrentOperationalArcProof.json').read_bytes());assert proof['internal_accuracy_scale']==1
proof.update(actual_components=257,all256_other_actual_components_preserved=True,
    only_internal_accuracy_scale_and_comment_changed_relative963=True,original_requested_force_options_bitexact963=True,
    internal_accuracy_and_lambda_noise_resolution_explicitly_different_from963=True,
    field_limit005_materials_orientation_locality_and_budgets_unchanged=True)
(root/'RequestedArcAccuracyProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('964 SCOPE VERIFIED;INTERNAL SCALE1 vs1e-4;ORIGINAL USER OPTIONS AND OTHER256 COMPONENTS EXACT963')
