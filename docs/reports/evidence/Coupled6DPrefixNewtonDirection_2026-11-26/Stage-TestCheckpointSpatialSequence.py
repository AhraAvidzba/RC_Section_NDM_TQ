"""ДЛЯ ТЕСТОВ. Запускает ограниченное продолжение пространственной секущей с проверкой обеих оценок точности lambda."""

from argparse import ArgumentParser
from hashlib import sha256
import json
from pathlib import Path
import re
from runpy import run_path


def TestStage(repo, root, source, checkpoint, binary=None):
    """Переиспользует исходный материал и решатель; управляющий цикл существует лишь в тестовом модуле."""
    here = Path(__file__).resolve().parent
    run_path(str(here / 'Stage-TestCheckpointResume.py'))['TestStage'](repo,root,source,checkpoint)
    prep = json.loads((root / 'Preparation.json').read_bytes())
    if binary is not None:
        data = run_path(str(here / 'Check-TestActualReverseCapture.py'))['TestReadBinary'](binary)
        name = 'Input/Reverse1.bin'
        (root / name).write_bytes(binary.read_bytes())
        prep['inputs'][name] = sha256((root / name).read_bytes()).hexdigest().upper()
        prep['directed_scope'].update(current_input_states=len(data['fields']),current_lambda=float(data['lambdas'][-1]),
            accepted_tail_binary_double_values_preserved=True,binary_input_origin=str(binary),
            original_decimal_checkpoint_prefix_states=29)
    name = 'tests/modTestCoupledElasticState.bas'
    code = (root / name).read_text('utf-8-sig')
    body = (here / 'TestNativeCheckpointSpatialSequence.bas.txt').read_text('utf-8-sig')
    code, count = re.subn(r'(?ms)^Public Sub TestRunReversePolishModel\(.*?^End Sub',lambda _: body,code)
    assert count == 1 and max(map(len,code.splitlines())) < 1024
    (root / name).write_bytes(code.replace('\n','\r\n').encode('utf-8-sig'))
    prep['inputs'][name] = sha256((root / name).read_bytes()).hexdigest().upper()
    for name in (Path(__file__).name,'TestNativeCheckpointSpatialSequence.bas.txt'):
        (root / name).write_bytes((here / name).read_bytes()); prep['inputs'][name] = sha256((root / name).read_bytes()).hexdigest().upper()
    prep['directed_scope'].update(case='ReplayedCheckpointSpatialSecantSequence',
        maximum_new_accepted_steps=12, maximum_attempts=24, original_material_accuracy=True,
        actual_candidate_and_current_lambda_bounds_required=True, abs_gain_required_above_sum_of_bounds_and_resolution=True,
        immutable_binary_accepted_snapshots=True, accepted_native_double_fields_unchanged=True,
        working_sources_changed=False, installed_workbook_changed=False, physical_capacity_certified=False)
    (root / 'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
    print('STRICT SPATIAL SEQUENCE CURRENT AND CANDIDATE UNCERTAINTY WITH IMMUTABLE BINARY SNAPSHOTS READY')


if __name__ == '__main__':
    parser = ArgumentParser(description=__doc__)
    parser.add_argument('--repo', type=Path, default=Path.cwd())
    parser.add_argument('--root', type=Path, required=True)
    parser.add_argument('--source', type=Path, required=True)
    parser.add_argument('--checkpoint', type=Path, required=True)
    parser.add_argument('--binary', type=Path)
    args = parser.parse_args()
    TestStage(args.repo.resolve(),args.root.resolve(),args.source.resolve(),args.checkpoint.resolve(),args.binary.resolve() if args.binary else None)
