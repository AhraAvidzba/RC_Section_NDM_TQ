"""ДЛЯ ТЕСТОВ. Готовит настоящий Trace от103 принятых состояний с текущими рабочими методами Newton и контроллером."""
from argparse import ArgumentParser
from hashlib import sha256
import json,re
from pathlib import Path
from runpy import run_path


def TestMethods(code):
    """Отделяет точные тела методов от тестовых дополнений, не нормализуя формулы и строки."""
    pattern=r'(?ms)^(?:Public|Private|Friend) (Function|Sub|Property (?:Get|Let|Set)) (\w+)\b.*?^End (?:Function|Sub|Property)[ \t]*$'
    return {(m[1],m[2]):m for m in re.finditer(pattern,code)}


def TestStage(repo,root,parent):
    """Восстанавливает140 рабочих методов; дополнительные читающие методы остаются только в частном тестовом проекте."""
    here=Path(__file__).resolve().parent
    proof=json.loads((parent/'DepartureRecoveryProof.json').read_bytes())
    assert proof['one_probe_only'] and proof['ordinary_adoption_and_spatial_checks_executed']
    run_path(str(here/'Stage-TestWorkingControllerResume.py'))['TestStage'](root,parent)
    (root/'DepartureRecoveryProof.json').rename(root/'Inherited.DepartureRecoveryProof.json')
    prep=json.loads((root/'Preparation.json').read_bytes())
    name='src/Solver/CCoupledNewtonSolver.cls'
    working=(repo/name).read_text('utf-8-sig');private=(root/name).read_text('utf-8-sig')
    before=TestMethods(private);actual=TestMethods(working)
    changed=[key for key in actual if actual[key].group()!=before[key].group()]
    assert set(changed)=={('Function','SolveEnvelopeTarget'),('Function','TryEnvelopeStep'),('Function','TryEnvelopeResultantCorrection')}
    assert len(actual)==140 and private[:min(m.start() for m in before.values())]==working[:min(m.start() for m in actual.values())]
    for key in changed:private=private.replace(before[key].group(),actual[key].group(),1)
    after=TestMethods(private)
    assert all(after[key].group()==value.group() for key,value in actual.items())
    controller=(repo/'src/Solver/CCoupledDeformationContinuation.cls').read_bytes()
    imports=['CCoupledDeformationContinuation','CCoupledNewtonSolver','modTestCoupledElasticState']
    compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
    anchor="-ModuleNames 'modTestCoupledElasticState'";assert compiler.count(anchor)==1
    compiler=compiler.replace(anchor,'-ModuleNames @('+','.join(repr(x) for x in imports)+')')
    scope=dict(existing_Newton_methods_exact_current=140,restored_methods=[list(x) for x in changed],
        added_private_test_methods=[list(x) for x in after if x not in actual],
        working_Newton_sha256=sha256((repo/name).read_bytes()).hexdigest().upper(),
        controller_exact_current=True,initial_states=103,full_cold_start=False,physical_capacity_certified=False)
    values={name:private.replace('\n','\r\n').encode('utf-8-sig'),
        'src/Solver/CCoupledDeformationContinuation.cls':controller,
        'Input/ForceWorkingNewton.cls':(repo/name).read_bytes(),
        'Compile-TestExternalPoint.ps1':compiler.replace('\n','\r\n').encode('utf-8-sig'),
        'WorkingNewtonMethodScope.json':(json.dumps(scope,indent=2,sort_keys=True)+'\n').encode(),
        Path(__file__).name:Path(__file__).read_bytes()}
    for item,raw in values.items():
        (root/item).write_bytes(raw);prep['inputs'][item]=sha256(raw).hexdigest().upper()
    prep['directed_scope'].update(case='WorkingDepartureActualTrace103',actual_import_this_pass=imports,
        current_input_states=103,actual_working_Trace_API=True,main_Trace_loop_not_executed=False,
        injected_post_H_departure_state=False,working_reserve_and_admission_methods_executed=False,
        no_prescribed_material_point_level_or_future_field=True,existing_Newton_methods_exact_current=140,
        prescribed_future_field=False,initial_history_replayed_from_zero=True,full_cold_start=False,
        actual_current_compilation_reused_without_import=False,single_test_module_compiled=False,
        fresh_compilation_claimed=True,global_probe_limit_unchanged=96,primary_iteration_budget=30,
        working_output_book_unchanged=True,physical_capacity_certified=False)
    (root/'Preparation.json').write_bytes((json.dumps(prep,ensure_ascii=False,indent=2,sort_keys=True)+'\n').encode())
    print('ACTUAL TRACE103 READY; CURRENT140 NEWTON METHODS; CURRENT CONTROLLER; THREE IMPORTS; NO FUTURE FIELD')


if __name__=='__main__':
    p=ArgumentParser(description=__doc__);p.add_argument('--repo',type=Path,default=Path.cwd())
    for name in ('root','parent'):p.add_argument('--'+name,type=Path,required=True)
    a=p.parse_args();TestStage(a.repo.resolve(),a.root.resolve(),a.parent.resolve())
