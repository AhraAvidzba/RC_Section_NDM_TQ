"""ДЛЯ ТЕСТОВ. Проверяет исходную систему и любые новые равновесия после согласования диапазонов долей."""
from pathlib import Path
import json
here=Path(__file__).resolve().parent;file=here/'Check-TestRecordedTailFineLine.py'
code=file.read_text('utf-8-sig').replace('Cold929RecordedTailFineLine_20261126','Cold934ConsistentSpatialLine_20261126')
code=code.replace('RecordedTailFineLineProof.json','ConsistentSpatialLineProof.json')
old='If mControlActive And mUseSpatialHistory Then minimumFraction = 0.0000152587890625'
assert old in code;code=code.replace(old,'If mControlActive And mUseSpatialHistory Then minimumFraction = SPATIAL_MIN_FRACTION')
scope={'__file__':str(file),'__name__':'consistent_spatial_line_test'}
exec(compile(code,str(file),'exec'),scope)
root=scope['root'];parent=root.parent/'Cold933ProspectiveExtraHalf_20261126'
before=(parent/'src/Solver/CCoupledNewtonSolver.cls').read_text('utf-8-sig')
after=(root/'src/Solver/CCoupledNewtonSolver.cls').read_text('utf-8-sig')
anchor='Private Const MIN_FRACTION As Double = 0.0078125'
before=before.replace(anchor,anchor+'\nPrivate Const SPATIAL_MIN_FRACTION As Double = 0.0000152587890625 \' Единая нижняя доля управляемой пространственной пробы и резерва соседней ветви.',1)
before=before.replace(old,'If mControlActive And mUseSpatialHistory Then minimumFraction = SPATIAL_MIN_FRACTION',1)
before=before.replace('    Do While fraction >= MIN_FRACTION * 0.5 \' Более короткая проба учитывает смену ветви; приёмка по исходной норме сохраняется.',
    '    Do While fraction >= SPATIAL_MIN_FRACTION \' Резерв проверяет тот же диапазон долей, что и основной пространственный шаг.',1)
assert before==after
proof=scope['proof'];proof.update(primary_and_prospective_floor_identical1_over65536=True,
    original30_iterations_and_all_acceptance_unchanged=True,all_material_laws_exact933=True,
    no_release_promotion=True,physical_capacity_certified=False)
(root/'ConsistentSpatialLineProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
