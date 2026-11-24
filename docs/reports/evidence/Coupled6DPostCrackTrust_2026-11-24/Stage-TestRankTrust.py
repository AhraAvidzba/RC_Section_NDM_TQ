"""ДЛЯ ТЕСТОВ. Исправляет обработку допустимой внутренней поправки при вырожденном QR по первичному lmpar."""
from pathlib import Path
from hashlib import sha256
import json
from runpy import run_path


def TestStage(repo,root,parent):
    """Сохраняет отрицательный опыт696; новую сборку проверяет тем же тестом и независимыми линейными случаями."""
    here=Path(__file__).resolve().parent
    run_path(str(here/'Stage-TestPostCrackTrust.py'))['TestStage'](repo,root,parent)
    prep=json.loads((root/'Preparation.json').read_bytes())
    path='src/Solver/CCoupledHybridFactorization.cls';code=(root/path).read_text('utf-8-sig')
    before='''    fullRank = TryDampedStep(0#, direction, derivative)
    If fullRank Then'''
    after='''    fullRank = TryDampedStep(0#, direction, derivative)
    If Not fullRank Then
        If Not TryUndampedRankStep(direction) Then Exit Function
        norm = VectorNorm(direction, 1, mCount)
        If norm <= radius Then parameter = 0#: TryLeastSquaresRadius = True: Exit Function
    End If
    If fullRank Then'''
    assert code.count(before)==1;code=code.replace(before,after,1)
    # Нулевая нижняя граница допускает внутреннюю регуляризованную поправку, когда минимальная норма меньше радиуса.
    before='''        norm = VectorNorm(direction, 1, mCount): phi = norm - radius
        If norm = 0# Or derivative >= 0# Then Exit Function'''
    after='''        Dim previousPhi As Double: previousPhi = phi
        norm = VectorNorm(direction, 1, mCount): phi = norm - radius
        If lower = 0# And phi <= previousPhi And previousPhi < 0# Then TryLeastSquaresRadius = True: Exit Function
        If norm = 0# Or derivative >= 0# Then Exit Function'''
    assert code.count(before)==1;code=code.replace(before,after,1)
    code+='\n'+(here/'CandidateUndampedRank.cls.txt').read_text('utf-8-sig')
    imports=sorted(prep['directed_scope']['actual_import_this_pass']+['CCoupledHybridFactorization'])
    raw=code.replace('\n','\r\n').encode('utf-8-sig');(root/path).write_bytes(raw);prep['inputs'][path]=sha256(raw).hexdigest().upper()
    compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
    compiler=compiler.replace("-ModuleNames 'modTestCoupledElasticState'",'-ModuleNames @('+','.join(repr(x) for x in imports)+')')
    path='Compile-TestExternalPoint.ps1';raw=compiler.replace('\n','\r\n').encode('utf-8-sig');(root/path).write_bytes(raw);prep['inputs'][path]=sha256(raw).hexdigest().upper()
    for name in (Path(__file__).name,'CandidateUndampedRank.cls.txt'):
        raw=(here/name).read_bytes();(root/name).write_bytes(raw);prep['inputs'][name]=sha256(raw).hexdigest().upper()
    prep['modules']=list(dict.fromkeys(prep['modules']+imports));prep['directed_scope'].update(actual_import_this_pass=imports,rank_deficient_interior_step_candidate=True)
    (root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
    print('RANK-AWARE INTERIOR LM CANDIDATE STAGED; ORIGINAL TEST EXPECTATIONS')


if __name__=='__main__':
    repo=Path.cwd();base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
    TestStage(repo,base/'Cold697PostCrackRankTrust_20261124',base/'Cold691HybridArithmeticRetry_20261124')
