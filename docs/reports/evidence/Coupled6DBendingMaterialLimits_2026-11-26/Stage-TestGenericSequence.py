"""ДЛЯ ТЕСТОВ. Проверяет общий адаптер в принятой истории и возврат к штатному управлению деформацией."""

from argparse import ArgumentParser
from hashlib import sha256
import json
from pathlib import Path
from runpy import run_path


def TestStage(repo, root, source):
    """Крупная отдельная цель не подменяется успешной; проверяется иной алгоритмический путь с последовательным принятием истории."""
    here = Path(__file__).resolve().parent
    run_path(str(here / 'Stage-TestYieldProjectedSequence.py'))['TestStage'](repo, root, source, parent='Stage-TestGenericConstraint.py')
    prep = json.loads((root / 'Preparation.json').read_bytes())
    name = 'tests/modTestCoupledElasticState.bas'
    code = (root / name).read_text('utf-8-sig')
    old = '        If probeCase = 0 Then target = 0.001 Else target = 0.01'
    assert code.count(old) == 1
    code = code.replace(old, '', 1)
    old = '    ReturnPrecisionFactorTest = 0#\n    Close #handle: mPolishExecutedTest = True'
    assert code.count(old) == 1
    code = code.replace(old, '    TestReturnToOrdinaryControl solver, handle\n' + old, 1)
    old = '    If caseIndex = 999 Then path = mHardPreparedPathTest & ".Sequence"'
    assert code.count(old) == 1
    code = code.replace(old, old + '\n    If caseIndex = 998 Then path = mHardPreparedPathTest & ".Bridge"', 1)
    code += '\n' + (here / 'TestNativeOrdinaryBridge.bas.txt').read_text('utf-8-sig')
    assert max(map(len, code.splitlines())) < 1024
    (root / name).write_bytes(code.replace('\n', '\r\n').encode('utf-8-sig'))
    prep['inputs'][name] = sha256((root / name).read_bytes()).hexdigest().upper()
    for filename in (Path(__file__).name, 'TestNativeOrdinaryBridge.bas.txt', 'Stage-TestYieldProjectedSequence.py'):
        (root / filename).write_bytes((here / filename).read_bytes())
        prep['inputs'][filename] = sha256((root / filename).read_bytes()).hexdigest().upper()
    prep['directed_scope'].update({'case': 'GenericProjectionSequenceAndOrdinaryBridge',
        'event_target_MPa': .001, 'event_targets_by_case_MPa': None,
        'case_ids': list(range(4)), 'step_limit': 4, 'bridge_step_limit': 2,
        'both_complete_Newton_and_existing_hybrid_attempted': False,
        'generic_projection_used_only_for_first_four_steps': True,
        'ordinary_TryDeformationStep_without_scalar_adapter_for_bridge': True,
        'ordinary_bridge_uses_original_equilibrium_tolerances': True,
        'same_two_seeds_and_original_history_for_both_numerical_engines': False,
        'ordinary_hybrid_fallback_allowed_for_bridge': True,
        'new_working_classes_not_enabled_in_ordinary_solver': True,
        'installed_workbook_changed': False})
    (root / 'Preparation.json').write_text(json.dumps(prep, indent=2, sort_keys=True)+'\n', encoding='utf-8')
    print('GENERIC PROJECTION SEQUENCE AND ORDINARY CONTINUATION BRIDGE READY')


if __name__ == '__main__':
    parser = ArgumentParser(description=__doc__)
    parser.add_argument('--repo', type=Path, default=Path.cwd())
    parser.add_argument('--root', type=Path, required=True)
    parser.add_argument('--source', type=Path, required=True)
    args = parser.parse_args()
    TestStage(args.repo.resolve(), args.root.resolve(), args.source.resolve())
