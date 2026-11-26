"""ДЛЯ ТЕСТОВ. Готовит направленные проверки максимума ветви, закрытой пределом стали."""

from argparse import ArgumentParser
from hashlib import sha256
import json
from pathlib import Path
from runpy import run_path
from subprocess import check_output


def TestStage(repo, root, before):
    """Закрепляет реальные рабочие тесты и, для отрицательного контроля, точное ядро контрольной точки."""
    here = Path(__file__).resolve().parent
    run_path(str(here / 'Stage-TestSteelHistory.py'))['TestStage'](repo, root)
    prep = json.loads((root / 'Preparation.json').read_bytes())
    if before:
        name = 'src/Solver/CCoupledDeformationContinuation.cls'
        raw = check_output(['git', 'cat-file', 'blob', '20ecfaa:' + name], cwd=repo)
        raw = raw.replace(b'\r\n', b'\n').replace(b'\n', b'\r\n')
        (root / name).write_bytes(raw)
        prep['inputs'][name] = sha256(raw).hexdigest().upper()
    name = 'Run-TestWorkingASD.ps1'
    code = (root / name).read_text('utf-8-sig')
    code = code.replace("'Compile','SteelCrossing'", "'Compile','ClosedPeak','PeakRegression','SteelCrossing'", 1)
    needle = "        if ($Phase -eq 'SteelCrossing') {"
    assert code.count(needle) == 1
    code = code.replace(needle, """        if ($Phase -eq 'ClosedPeak') {
            $arguments+=@('-Macro','modTestCoupledElasticState.RunTestCoupledClosedBranchPeak','-TimeoutSeconds','165')
        } elseif ($Phase -eq 'PeakRegression') {
            $suite=@('RunTestCoupledDeformationControl','RunTestCoupledRetainedShoulders','RunTestCoupledDesignWitness',
                'RunTestCoupledDiscardedShoulder','RunTestCoupledSteelAtDesign','RunTestCoupledMinimumLoadStep','RunTestCoupledControlProgress')
            $arguments+=@('-Macro',(($suite|ForEach-Object {'modTestCoupledElasticState.'+$_}) -join ','),'-TimeoutSeconds','240')
        } elseif ($Phase -eq 'SteelCrossing') {""")
    (root / name).write_bytes(code.replace('\n', '\r\n').encode('utf-8-sig'))
    script = Path(__file__)
    (root / script.name).write_bytes(script.read_bytes())
    for relative in (name, script.name):
        prep['inputs'][relative] = sha256((root / relative).read_bytes()).hexdigest().upper()
    prep['directed_scope'] = {'reference_kernel': '20ecfaa' if before else 'working',
        'phase': 'ClosedPeak', 'analytical_closed_cases': 3, 'open_numerical_case': 1,
        'global_probe_limit': 96, 'requested_lambda_tolerance': .0001,
        'full_acceptance_completed': False}
    (root / 'Preparation.json').write_text(json.dumps(prep, indent=2, sort_keys=True) + '\n', encoding='utf-8')
    print('CLOSED BRANCH PEAK DIRECTED FIXTURE READY', 'Before' if before else 'Working', flush=True)


if __name__ == '__main__':
    parser = ArgumentParser(description=__doc__)
    parser.add_argument('--repo', type=Path, default=Path.cwd())
    parser.add_argument('--root', type=Path, required=True)
    parser.add_argument('--before', action='store_true')
    args = parser.parse_args()
    TestStage(args.repo.resolve(), args.root.resolve(), args.before)
