"""ДЛЯ ТЕСТОВ: готовит независимый контроль согласованности принятых полей и следа после уточнения предела стали."""
from argparse import ArgumentParser
from hashlib import sha256
import json
from pathlib import Path
from runpy import run_path


def TestStage(repo,root):
    """Использует рабочий модуль аналитических тестов; не меняет модель, допуски или численные бюджеты."""
    here=Path(__file__).resolve().parent
    run_path(str(here/'Stage-TestWorkingASD.py'))['TestStage'](repo,root)
    prep=json.loads((root/'Preparation.json').read_bytes())
    test_name='tests/modTestCoupledElasticState.bas'
    (root/test_name).write_bytes((repo/test_name).read_bytes())
    assert 'modTestCoupledElasticState' not in prep['modules']
    prep['modules'].append('modTestCoupledElasticState')
    prep['inputs'][test_name]=sha256((root/test_name).read_bytes()).hexdigest().upper()
    name='Run-TestWorkingASD.ps1'
    code=(root/name).read_text('utf-8-sig')
    old="[ValidateSet('Compile','Core'";assert code.count(old)==1
    code=code.replace(old,"[ValidateSet('Compile','SteelCrossing','Core'")
    old="        if ($Phase -eq 'Core') {";assert code.count(old)==1
    code=code.replace(old,"""        if ($Phase -eq 'SteelCrossing') {
            $arguments+=@('-Macro','modTestCoupledElasticState.RunTestCoupledDeformationSteelCrossing','-TimeoutSeconds','165')
        } elseif ($Phase -eq 'Core') {""")
    (root/name).write_bytes(code.encode('utf-8-sig'))
    script=Path(__file__);(root/script.name).write_bytes(script.read_bytes())
    for relative in (name,script.name):prep['inputs'][relative]=sha256((root/relative).read_bytes()).hexdigest().upper()
    prep['directed_scope']={'phase':'SteelCrossing','analytical_cases':3,'unchanged_physical_status_expectations':True,
                            'global_probe_limit':96,'physical_capacity_certified':False,'full_acceptance_completed':False}
    (root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
    print('ANALYTICAL STEEL HISTORY ALIGNMENT TEST READY; WORKING NUMERICS UNCHANGED',flush=True)


if __name__=='__main__':
    parser=ArgumentParser(description=__doc__);parser.add_argument('--repo',type=Path,default=Path.cwd())
    parser.add_argument('--root',type=Path,required=True);args=parser.parse_args();TestStage(args.repo.resolve(),args.root.resolve())
