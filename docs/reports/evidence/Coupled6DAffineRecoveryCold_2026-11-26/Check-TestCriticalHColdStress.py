"""ДЛЯ ТЕСТОВ. Проверяет напряжения 56 холодных состояний с точным неизменяемым кодом материального повтора."""
from pathlib import Path
from runpy import run_path
import json,re
here=Path(__file__).resolve().parent
base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold882CriticalHColdStress_20261126'
source=base/'Cold880CriticalHCold96_20261126'
compiled=base/'Cold855BalancedInitialControl_20261125'

def TestReplayScope():
    """Отделяет материальный повтор от отличающихся алгоритмов поиска; проверяет исполняемые пути повтора целиком."""
    checked=[]
    for file in (source/'src').rglob('*'):
        if not file.is_file():continue
        if file.name in ('CCoupledNewtonSolver.cls','CCoupledArcContinuation.cls','CCoupledDeformationContinuation.cls'):continue
        rel=file.relative_to(source);assert file.read_bytes()==(compiled/rel).read_bytes(),str(rel);checked.append(str(rel))
    def method(path,name):
        code=path.read_text('utf-8-sig')
        match=re.findall(r'(?ms)^(?:Private |Friend |Public )(?:Sub|Function) '+re.escape(name)+r'\(.*?^End (?:Sub|Function)',code)
        assert len(match)==1,(str(path),name,len(match));return match[0]
    groups={
      'src/Solver/CCoupledNewtonSolver.cls':['ReplayAcceptedFields','EvaluatePointResponses','AcceptMaterialTrials','Evaluate','TestEquilibriumLinearization','PrepareRecentLoadingBoundary','PrepareAcceptedEnvelopeSegment'],
      'tests/modTestCoupledElasticState.bas':['TestRunReversePolishModel','TestExportFastFrozenReplay','TestObserveFrozenReplay','TestWriteHardPreparedStress','TestWriteHardPreparedOperators']}
    for rel,names in groups.items():
        for name in names:
            assert method(source/rel,name)==method(compiled/rel,name),(rel,name)
    return dict(unchanged_material_and_kinematic_files=checked,exact_replay_methods=groups,
                different_search_algorithms_not_called=True,entire_Newton_class_identity_not_claimed=True)

if __name__=='__main__':
    import sys
    scope=TestReplayScope()
    if '--scope-only' in sys.argv:
        (root/'ReplaySourceScope.json').write_text(json.dumps(scope,sort_keys=True,indent=2)+'\n',encoding='utf-8')
        print('EXACT MATERIAL/KINEMATICS AND REPLAY/EQUILIBRIUM METHODS;DIFFERENT SEARCH NOT CALLED')
    else:
        file=here/'Check-TestVerifiedColdStressReplay.py';code=file.read_text('utf-8-sig')
        begin=code.index('    for path in (source/\'src\').rglob(\'*\'):',code.index('if reuse:'))
        end=code.index("    code=(root/name).read_text('utf-8-sig')",begin)
        code=code[:begin]+"    assert json.loads((root/'ReplaySourceScope.json').read_bytes())==expected_replay_scope\n"+code[end:]
        code=code.replace('all_material_Newton_and_kinematic_methods_exact_source=True','material_and_kinematic_files_and_replay_methods_exact_source=True')
        saved=list(sys.argv);sys.argv=[str(file),'--root',str(root)]
        try:exec(compile(code,str(file),'exec'),{'__file__':str(file),'__name__':'critical_H_cold_stress_test','expected_replay_scope':scope})
        finally:sys.argv=saved
