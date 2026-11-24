"""ДЛЯ ТЕСТОВ. Закрепляет текущий рабочий численный код,213 направленных проверок и48 независимых линейных задач."""
from pathlib import Path
from hashlib import sha256
from runpy import run_path
import json,subprocess


def TestRecord(repo,root):
    """Проверяет фактическую книгу и собственные процессы; не объявляет предел HardShear или полную приёмку."""
    here=Path(__file__).resolve().parent
    native=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative']
    regression,text=native(root,'WorkingRank');linear_native,_=native(root,'TrustStep')
    assert regression==dict(passed=213,failed=0,macros=5,timeout600=False,own_process_closed=True)
    assert linear_native==dict(passed=1,failed=0,macros=1,timeout600=False,own_process_closed=True)
    linear=json.loads((root/'RankTrustLinearProof.json').read_bytes())
    assert len(linear['original40']['cases'])==40 and len(linear['extra8'])==8
    paths=['src/Solver/CCoupledHybridFactorization.cls','src/Solver/CCoupledHybridSolver.cls',
       'tests/CHybridEquationsTest.cls','tests/modTestCoupledHybrid.bas']
    hashes={path:sha256((repo/path).read_bytes()).hexdigest().upper() for path in paths}
    for path in paths:assert (repo/path).read_bytes()==(root/path).read_bytes(),path
    # Полные тексты этих двух рабочих классов не менялись, включая численные методы и диагностику.
    for path in ('src/Solver/CCoupledNewtonSolver.cls','src/Solver/CCoupledDeformationContinuation.cls'):
        expected=subprocess.check_output(['git','show','HEAD:'+path],cwd=repo).decode('utf-8-sig').replace('\r\n','\n')
        assert (repo/path).read_text('utf-8-sig')==expected,path
    assert 'mRowCount' not in (repo/paths[0]).read_text('utf-8-sig')
    audit=json.loads((root/'PointTraceSourceComposition.json').read_bytes())
    compilation=json.loads((root/'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
    assert audit['components']==255 and len(audit['other_actual_components_exact_parent_except_identifier_case'])==250
    assert compilation['Compiled'] and not compilation['ProductionChanged'] and len(compilation['Modules'])==5
    assert sha256((root/'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper()==compilation['WorkbookSHA256']
    out=repo/'docs/reports/evidence/Coupled6DWorkingRankKernel_2026-11-24';out.mkdir(parents=True,exist_ok=False)
    files={}
    for name in ('Native/WorkingRank.txt','Native/WorkingRank.txt.ProcessClosure.json','Native/TrustStep.txt',
       'Native/TrustStep.txt.ProcessClosure.json','Native/TrustStepResults.bin','Input/TrustSteps.bin','Input/TrustStepCases.json',
       'RankTrustLinearProof.json','PointTraceSourceComposition.json','CompilePointClosure.json'):
        target=out/Path(name).name;target.write_bytes((root/name).read_bytes());files[target.name]=sha256(target.read_bytes()).hexdigest().upper()
    for name in ('Stage-TestRankTrustLinear.py','Check-TestRankTrustLinear.py','Stage-TestWorkingRankKernel.py',Path(__file__).name):
        target=out/name;target.write_bytes((here/name).read_bytes());files[name]=sha256(target.read_bytes()).hexdigest().upper()
    (out/'.gitattributes').write_text('* -text whitespace=blank-at-eol,blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8',newline='\n')
    proof=dict(working_sources_sha256=hashes,working_controller_and_Newton_unchanged=True,
       private_rectangular_QR_extension_not_promoted=True,native_regression=regression,
       native_linear_output_record=linear_native,linear_cases_independently_verified=48,
       original40_criteria_and_binary_body_unchanged=True,actual_book_components_verified=255,
       actual_book_sha256=compilation['WorkbookSHA256'],evidence_files_sha256=files,private_root=str(root),
       new_cold_HardShear_not_run=True,physical_capacity_certified=False,full_acceptance_completed=False,
       entire_goal_completed=False,primary_source='literature/Coupled6D/NumericalControl/Minpack/src/lmpar.c')
    (out/'Proof.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
    (root/'WorkingRankKernelProof.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
    print('WORKING RANK KERNEL VERIFIED;213/0 AND48 LINEAR CASES;255ACTUAL COMPONENTS; NO CAPACITY CLAIM')


if __name__=='__main__':TestRecord(Path.cwd(),Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold699WorkingRankKernel_20261124'))
