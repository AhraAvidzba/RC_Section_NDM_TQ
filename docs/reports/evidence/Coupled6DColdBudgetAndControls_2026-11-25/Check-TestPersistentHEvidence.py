"""ДЛЯ ТЕСТОВ. Проверяет фактический VBA, известный предел, отрицательные варианты и штатный исход смешанного продолжения."""
from argparse import ArgumentParser
from hashlib import sha256
from io import BytesIO
import json
from pathlib import Path
import re
from runpy import run_path
from zipfile import ZipFile


def TestRestoreAliases(root):
    """Восстанавливает только точные дубликаты внутри распакованного комплекта и проверяет обе контрольные суммы."""
    file = root / 'ExactFileAliases.json'
    if not file.is_file(): return
    aliases = json.loads(file.read_bytes())
    pins = json.loads((root / 'Files.json').read_bytes())
    for target, source in aliases.items():
        destination = (root / target).resolve(); original = (root / source).resolve()
        assert destination.is_relative_to(root.resolve()) and original.is_relative_to(root.resolve())
        assert target != source and source not in aliases and pins[target] == pins[source]
        content = original.read_bytes()
        assert sha256(content).hexdigest().upper() == pins[target]
        if destination.exists(): assert destination.read_bytes() == content
        else:
            destination.parent.mkdir(parents=True, exist_ok=True)
            with destination.open('xb') as output: output.write(content)


def TestSources(repo, root):
    """Переиспользует точную сверку CFB с отдельным контрактом добавленного тестового модуля и различным числом компонентов."""
    here = Path(__file__).resolve().parent
    source = here / 'Check-TestEnvelopeFailureEvidence.py'
    text = source.read_text('utf-8-sig'); begin = text.index('def TestProof(repo,root):'); end = text.index('    numerical={',begin)
    body = text[begin:end] + '    return sources,binaries\n'
    body = body.replace('def TestProof(repo,root):','def TestSourceProof(repo,root,expected):',1)
    begin = body.index('    expected={'); end = body.index('    for key in cases:',begin)
    body = body[:begin] + '    assert set(expected)==set(cases)\n' + body[end:]
    body = body.replace("expected.get(key,['modTestCoupledElasticState'])", 'expected[key]')
    body = body.replace("assert (audit['components'],audit['imported'],audit['unchanged'])==(254,41,213)",
        "count=254 if key in ('Recovery','Initial') else 255\n        assert (audit['components'],audit['imported'],audit['unchanged'])==(count,count-213,213)")
    body = body.replace('assert len(parent)==len(actual)==254 and set(parent)==set(actual)',
        "assert len(actual)==count and set(parent).issubset(actual)\n        assert set(actual)-set(parent)==({'modTestCoupledEnvelopeControl'} if key=='UnitOriginal' else set())")
    body = body.replace('other_components_exact_except_identifier_case=254-len(imports)',
        'other_components_exact_except_identifier_case=count-len(imports)')
    body = body.replace("assert prepared.count(line)==1;prepared=prepared.replace(line,'',1)",
        "assert prepared.count(line)<=1;prepared=prepared.replace(line,'',1)")
    namespace = run_path(str(source)); exec(compile(body,str(source),'exec'),namespace)
    trio = ['CCoupledNewtonSolver','CCoupledDeformationContinuation','modTestCoupledElasticState']
    expected = {'Recovery':trio,'Initial':trio,'UnitOriginal':['modTestCoupledEnvelopeControl'],
        'Balanced01':['CCoupledNewtonSolver','CCoupledDeformationContinuation','modTestCoupledEnvelopeControl'],
        'Bounded02':['CCoupledNewtonSolver','CCoupledDeformationContinuation','modTestCoupledEnvelopeControl'],
        'Capture':['CCoupledNewtonSolver','CCoupledDeformationContinuation','modTestCoupledEnvelopeControl'],
        'Difference':['CCoupledNewtonSolver','CCoupledDeformationContinuation','modTestCoupledEnvelopeControl','CCoupledHybridSolver','modTestCoupledHybrid'],
        'Mixed':['CCoupledNewtonSolver','CCoupledDeformationContinuation','modTestCoupledEnvelopeControl','CCoupledHybridSolver','modTestCoupledHybrid','modTestCoupledElasticState']}
    return namespace['TestSourceProof'](repo,root,expected)


def TestNative(folder, name):
    """Различает завершённые проверки, численный исход и тайм-аут, подтверждая закрытие собственного процесса."""
    file = folder / 'Native' / (name + '.txt'); text = file.read_text('utf-8-sig')
    closure = json.loads((folder / 'Native' / (name + '.txt.ProcessClosure.json')).read_bytes())
    totals = re.findall(r'(?m)^TOTAL[^\r\n]*passed=(\d+); failed=(\d+)',text)
    timeout = 'WATCHDOG_FAILURE: timeout=600 sec;' in text
    if timeout:
        assert 'ownExcelState=Absent; exactCleanupRequested=True' in text and not totals
        assert closure['Stage']=='TIMEOUT' and not closure['Passed'] and closure['Identity']['Valid']
        assert closure['Before']['State']=='Live' and closure['After']['State']=='Absent'
        assert closure['Identity']['PID']==closure['Before']['PID']==closure['After']['PID']
        assert str(closure['Identity']['Ticks'])==closure['Before']['StartTicksUTC']==closure['After']['StartTicksUTC']
        assert closure['ExactOwnExcelStopRequestedOnlyByExplicitTimeout'] and closure['ExcelStopError'] is None
    else:
        assert closure['ChildHasExited'] and closure['Passed']
        assert totals and 'SOURCE_UNCHANGED: True' in text and 'RUNTIME ERROR' not in text
        assert closure['ChildExitCode'] == (1 if any(int(failed) for _,failed in totals) else 0)
    return dict(passed=sum(int(passed) for passed,_ in totals),failed=sum(int(failed) for _,failed in totals),
        macros=len(totals),timeout600=timeout,own_process_closed=True),text


def TestStock(repo, folder, compiled):
    """Сверяет штатный исход, сохранённые Results и точные исходники фактически повторно открытой книги."""
    native,text = TestNative(folder,'HBackendStock')
    if native['timeout600']: return native
    assert native['failed'] == 0
    state, = re.findall(r'(?m)^TIMING_STATES: direct=(\w+); capacity=(\w+); confirmed lower=([^;]*); limit=([^;]*); code=(\w+)$',text)
    before,after = re.findall(r'(?m)^RESULTS_SAVE_REOPEN: before=(\w+); after=(\w+); equal=True;',text)[0]
    assert before == after
    saved_file = folder / 'SavedWorkbook.xlsm'
    with ZipFile(saved_file) as book: saved = book.read('xl/vbaProject.bin')
    here = Path(__file__).resolve().parent
    helper,_,reader = run_path(str(here / 'Inspect-TestResolutionVBA.py'))['TestHelpers'](repo)
    _,original = reader(helper['TestContainer'](compiled)); _,reopened = reader(helper['TestContainer'](saved))
    assert original == reopened
    return dict(native=native,direct=state[0],capacity=state[1],lower=float(state[2]) if state[2] else None,
        limit=state[3],code=state[4],seconds=float(re.findall(r'(?m)^TIMING_SECONDS: ([^ |]+)',text)[0]),
        results_checksum=before,all_saved_sources_exact=len(original),saved_book_sha256=sha256(saved_file.read_bytes()).hexdigest().upper())


def TestProof(repo, root):
    """Сохраняет известные неудачи, не заменяя ими физический предел; полнота всего ТЗ не объявляется."""
    TestRestoreAliases(root)
    sources,binaries = TestSources(repo,root)
    cases = root / 'Cases'; results = {}
    for key in ('Recovery','Initial','Difference','Mixed'):
        regression,_ = TestNative(cases / key,'HBackendRegression')
        assert regression == dict(passed=540,failed=0,macros=14,timeout600=False,own_process_closed=True)
        results[key] = dict(regression=regression,stock=TestStock(repo,cases / key,binaries[key]))
    assert results['Recovery']['stock']['capacity']=='NumFail' and abs(results['Recovery']['stock']['lower']-1.10786819990302)<1e-13
    assert results['Initial']['stock']['timeout600']
    assert results['Difference']['stock']['capacity']=='NumFail' and abs(results['Difference']['stock']['lower']-1.05220775947834)<1e-13
    for key in ('UnitOriginal','Balanced01','Bounded02','Capture'):
        result,_ = TestNative(cases / key,'PersistentHUnit')
        assert (result['passed'],result['failed']) == (23,1)
        results[key] = result
    for key in ('Difference','Mixed'):
        hybrid,_ = TestNative(cases / key,'HBackendUnit'); physical,text = TestNative(cases / key,'PersistentHUnit')
        assert (hybrid['passed'],hybrid['failed'])==(75,0) and (physical['passed'],physical['failed'])==(26,0)
        best = float(re.findall(r'наибольший допустимый lambda: ([0-9.E+-]+)',text)[0])
        accuracy = float(re.findall(r'Оценка погрешности lambda: ([0-9.E+-]+)',text)[0])
        assert abs(best-0.9)<=accuracy and accuracy<=0.001
        results[key].update(hybrid=hybrid,known_concrete_peak=physical,retained_lambda=best,accuracy=accuracy,reference_lambda=0.9)
    rank = run_path(str(Path(__file__).resolve().parent / 'Check-TestHAdjointRank.py'))['TestRank'](cases / 'Capture')
    assert rank['first_failure_stage'][1]=='MaterialTangent' and not rank['linear_matrix_captured']
    mechanics,fields = run_path(str(Path(__file__).resolve().parent / 'Validate-TestHardPrepared.py'))['TestValidate'](cases / 'Recovery','IntegratedEvent.Sequence')
    assert len(fields)==294
    direct, _ = TestNative(cases / 'Recovery','PersistentHController')
    failed, _ = TestNative(cases / 'Initial','PersistentHController'); assert failed['timeout600']
    mixed = (cases / 'Mixed/src/Solver/CCoupledDeformationContinuation.cls').read_text('utf-8-sig')
    line = '    modTestCoupledElasticState.TestCaptureStockContinuation Me, Details()\n'
    assert mixed.count(line)==1
    assert mixed.replace(line,'',1)==(root / 'WorkingSources/CCoupledDeformationContinuation.cls').read_text('utf-8-sig')
    return dict(sources=sources,results=results,captured_first_failure=rank,recovery_independent_equilibrium=mechanics,
        known_concrete_fcA_peak_verified=True,full_acceptance_completed=False,installed_workbook_changed=False,
        physical_capacity_of_entire_goal_certified=False,target10seconds_met=False,
        external_material_reintegration_completed=False)


if __name__ == '__main__':
    parser = ArgumentParser(description=__doc__)
    parser.add_argument('--repo',type=Path,default=Path.cwd()); parser.add_argument('--root',type=Path,required=True)
    args = parser.parse_args(); result = TestProof(args.repo.resolve(),args.root.resolve())
    (args.root / 'PersistentHEvidenceProof.json').write_bytes((json.dumps(result,ensure_ascii=False,indent=2,sort_keys=True)+'\n').encode())
    print('PERSISTENT H EVIDENCE VERIFIED; KNOWN fc*A PEAK PASSES; FULL COUPLED GOAL OPEN')
