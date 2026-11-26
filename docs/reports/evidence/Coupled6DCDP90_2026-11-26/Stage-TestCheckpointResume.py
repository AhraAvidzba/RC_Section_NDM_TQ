"""ДЛЯ ТЕСТОВ. Восстанавливает десятичный снимок принятой ветви и проверяет следующую фактическую цель без повторения полного поиска."""

from argparse import ArgumentParser
from hashlib import sha256
import json
from pathlib import Path
import re
from runpy import run_path
import struct
import numpy as np


def TestStage(repo, root, source, checkpoint):
    """Не выдаёт текстовый экспорт за исходные Double: каждое восстановленное поле обязано подтвердить исходные допуски."""
    here = Path(__file__).resolve().parent
    run_path(str(here / 'Stage-TestLiveDifference.py'))['TestStage'](repo,root,source)
    prep = json.loads((root / 'Preparation.json').read_bytes())
    old = run_path(str(here / 'Check-TestActualReverseCapture.py'))['TestReadBinary'](root / 'Input/Reverse1.bin')
    read = run_path(str(here / 'Read-TestContinuationCheckpoint.py'))['TestRead']
    snapshot_path = checkpoint / 'Native/Stock.Checkpoint.csv'
    snapshot = read(snapshot_path)
    probes_path = checkpoint / 'Native/Stock.Probes.csv'
    rows = [row.split(';') for row in probes_path.read_text('cp1251').splitlines()]
    probe, = [row for row in rows if row[:2] == ['PROBE','BEGIN'] and int(row[2]) == snapshot['sequence']]
    projection, = [row for row in rows if row[:3] == ['PROJECTION','BEGIN',probe[2]]]
    field = np.asarray([point['field'] for point in snapshot['points']],float)
    assert field.shape == (snapshot['count'],60) and field.shape[0] >= 2
    factors = np.asarray([point['lambda'] for point in snapshot['points']])
    fractions = np.asarray([point['fraction'] for point in snapshot['points']])
    assert np.all(np.isfinite(field)) and np.all(np.isfinite(factors))
    control = np.asarray(projection[3:],float)
    assert control.shape == (60,) and float(probe[4]) == factors[-1]
    target, seed_factor = float(probe[5]), float(probe[3])
    spacing = float(control @ (field[-1]-field[-2]))
    ratio = (target-float(control @ field[-1]))/spacing
    assert spacing > 0 and ratio > 0
    assert abs(seed_factor-(factors[-1]+ratio*(factors[-1]-factors[-2]))) < 1e-10
    data = bytearray(b'NDMREVTEST000001')
    data.extend(struct.pack('<i',60))
    def TestVector(values):
        """Записывает уже восстановленные значения как явные little-endian Double с ограниченным размером."""
        data.extend(struct.pack('<i',len(values)))
        data.extend(struct.pack('<'+'d'*len(values),*values))
    for key in ('calibration','options','offset','direction'):
        TestVector(old[key])
    data.extend(struct.pack('<i',len(field)))
    for fraction,factor,values in zip(fractions,factors,field):
        data.extend(struct.pack('<dd',fraction,factor)); TestVector(values)
    for values in (field[-2],field[-1],field[-1],old['offset']+factors[-1]*old['direction'],control,
        [target,seed_factor,0.,0.],[0.,0.,0.,0.]):
        TestVector(values)
    name = 'Input/Reverse1.bin'
    (root / name).write_bytes(data)
    prep['inputs'][name] = sha256(data).hexdigest().upper()
    for source_path,name in ((snapshot_path,'Input/OriginalCheckpoint.csv'),(probes_path,'Input/OriginalProbes.csv')):
        (root / name).write_bytes(source_path.read_bytes())
        prep['inputs'][name] = sha256((root / name).read_bytes()).hexdigest().upper()
    name = 'tests/modTestCoupledElasticState.bas'
    code = (root / name).read_text('utf-8-sig')
    body, = re.findall(r'(?ms)^Public Sub TestRunReversePolishModel\(.*?^End Sub',code)
    original = body
    anchor = '    If Not solver.ReplayAcceptedFields(fields, fields.Count) Then Err.Raise 5, , "Не восстановлена исходная история живого перехода."'
    assert body.count(anchor) == 1
    body = body.replace(anchor, '''    Set mReverseFieldsTest = fields: Set mReverseTraceTest = trace.CopyPrefix(trace.Count)
    TestExportReverseCandidate 0
''' + anchor,1)
    old_block = '''    If Not solver.SelectSpatialDirectionControl(previous, accepted, movement, projection, coordinate, spacing) Then _
        Err.Raise 5, , "Не построена исходная пространственная мера."'''
    assert body.count(old_block) == 1
    body = body.replace(old_block, '''    For degree = 1 To UBound(projection)
        coordinate = coordinate + projection(degree) * accepted(degree)
        spacing = spacing + projection(degree) * movement(degree)
    Next degree
    If spacing <= 0# Then Err.Raise 5, , "Не восстановлен положительный интервал исходного управления."
    Dim ratio As Double: ratio = (context(1) - coordinate) / spacing''',1)
    assert body.count('seedFactor = currentFactor + 0.5 * (currentFactor - previousFactor)') == 1
    body = body.replace('seedFactor = currentFactor + 0.5 * (currentFactor - previousFactor)', 'seedFactor = context(2)',1)
    assert body.count('accepted(degree) + 0.5 * movement(degree)') == 1
    body = body.replace('accepted(degree) + 0.5 * movement(degree)','accepted(degree) + ratio * movement(degree)',1)
    assert body.count('coordinate + 0.5 * spacing') == 2
    body = body.replace('coordinate + 0.5 * spacing','context(1)')
    code = code.replace(original,body,1)
    (root / name).write_bytes(code.replace('\n','\r\n').encode('utf-8-sig'))
    prep['inputs'][name] = sha256((root / name).read_bytes()).hexdigest().upper()
    name = Path(__file__).name
    (root / name).write_bytes(Path(__file__).read_bytes()); prep['inputs'][name] = sha256((root / name).read_bytes()).hexdigest().upper()
    prep['directed_scope'].update(case='ResumeLastCompleteDecimalCheckpoint', decimal_checkpoint_reconstructed=True,
        original_checkpoint_bits_preserved=False, original_material_accuracy=True,
        native_equilibrium_required_for_every_replayed_state=True, requested_probe=int(probe[2]),
        checkpoint_sequence=snapshot['sequence'], checkpoint_states=snapshot['count'],
        current_lambda=float(factors[-1]), seed_factor=seed_factor, requested_coordinate=target,
        predictor_ratio=ratio, physical_capacity_certified=False, installed_workbook_changed=False)
    (root / 'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
    print('DECIMAL CHECKPOINT REPLAY AND ACTUAL NEXT CONTROL TARGET READY',snapshot['count'],factors[-1],ratio)


if __name__ == '__main__':
    parser = ArgumentParser(description=__doc__)
    parser.add_argument('--repo', type=Path, default=Path.cwd())
    parser.add_argument('--root', type=Path, required=True)
    parser.add_argument('--source', type=Path, required=True)
    parser.add_argument('--checkpoint', type=Path, required=True)
    args = parser.parse_args()
    TestStage(args.repo.resolve(),args.root.resolve(),args.source.resolve(),args.checkpoint.resolve())
