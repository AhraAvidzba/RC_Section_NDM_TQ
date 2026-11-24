"""ДЛЯ ТЕСТОВ. Закрепляет устранение аналитического отказа и фактический исход посттрещинной ветви перед паузой."""
from hashlib import sha256
import json,difflib,math
from pathlib import Path
from runpy import run_path


def TestCheck(repo):
    """Сравнивает неизменные тесты, закрытие собственных процессов и все сохранённые уравнения; не объявляет полную приёмку."""
    here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
    before=base/'Cold696PostCrackFullTrust_20261124';after=base/'Cold697PostCrackRankTrust_20261124'
    parent=base/'Cold691HybridArithmeticRetry_20261124'
    native=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative']
    old,oldtext=native(before,'LeastSquaresTrust');new,newtext=native(after,'LeastSquaresTrust')
    assert (old['passed'],old['failed'])==(19,1) and (new['passed'],new['failed'])==(20,0)
    assert [v for v in oldtext.splitlines() if v.startswith('FAIL:')]==['FAIL: LM rank deficient existing roots']
    paths=['src/Solver/CCoupledDeformationContinuation.cls','src/Solver/CCoupledHybridSolver.cls',
      'src/Solver/CCoupledNewtonSolver.cls','tests/CHybridEquationsTest.cls','tests/modTestCoupledHybrid.bas','tests/modTestCoupledElasticState.bas']
    for path in paths:assert (before/path).read_bytes()==(after/path).read_bytes(),path
    assert (before/'src/Solver/CCoupledHybridFactorization.cls').read_bytes()!=(after/'src/Solver/CCoupledHybridFactorization.cls').read_bytes()
    for root,imports in ((before,6),(after,7)):
        compiled=json.loads((root/'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
        audit=json.loads((root/'PointTraceSourceComposition.json').read_bytes())
        assert compiled['Compiled'] and not compiled['ProductionChanged'] and len(compiled['Modules'])==imports
        assert sha256((root/'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper()==compiled['WorkbookSHA256']
        assert audit['components']==255 and len(audit['other_actual_components_exact_parent_except_identifier_case'])==255-imports
        assert json.loads((root/'CompilePointClosure.json').read_bytes())['State']=='Absent'
    capacity_native,_=native(after,'PostCrackTension')
    if capacity_native['timeout600']:
        import numpy as np
        read=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary']
        reference=read(base/'Cold692PostCrackArithmeticRetry_20261124/Native/PostCrackPrescribed.Accepted11.bin')
        current=read(after/'Native/PostCrackPrescribed.Accepted2.bin')
        for key in ('calibration','options','offset','direction'):assert np.array_equal(current[key],reference[key]),key
        for key in ('fields','fractions','lambdas'):assert np.array_equal(current[key],reference[key][:2]),key
        phase='PostCrackPrescribed.Accepted'
        part=run_path(str(here/'Check-TestPostCrackFields.py'))['TestAnalyse'](after,phase)
        mechanical=dict(phases={phase:part},automatic_continuation_not_exported_before_timeout=True,
           only_two_prescribed_fields_verified=True,capacity_certified=False,entire_goal_completed=False)
        (after/'PostCrackFieldsProof.json').write_text(json.dumps(mechanical,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
        physical=dict(native=capacity_native,particular_uniaxial_capacity_passed=False,
           known_maximum_lambda=(21168*3*math.exp(-3*125*(.002-.0001)/.12)+432*400)/120000,
           same_original_two_cold_fields_and_load_bits=True,returned_peak=None,reported_accuracy=None,
           full_prescribed_curve_not_checked_in_this_run=True,arbitrary_section_all6_capacity_certified=False,
           entire_goal_completed=False)
        (after/'PostCrackCapacityProof.json').write_text(json.dumps(physical,indent=2,sort_keys=True)+'\n',encoding='utf-8')
    else:
        physical=run_path(str(here/'Check-TestPostCrackCapacity.py'))['TestCheck'](after,base/'Cold692PostCrackArithmeticRetry_20261124')
    proof=dict(analytic_baseline=old,analytic_rank_repair=new,
       same_six_other_sources_and_tests_exact=True,only_rank_factorization_changed_between_analytic_runs=True,
       actual_VBA_components_per_book=255,physical_capacity=physical,
       working_code_contains_new_LM_candidate=False,entire_goal_completed=False,
       next_checks=['40 original linear SVD cases and new interior rank-deficient cases',
         'full prescribed real CDP and steel curve','cold HardShear and all LoadPath',
         'finite concrete physical-limit contract ONQ-008','full original workbook acceptance and performance'])
    out=repo/'docs/reports/evidence/Coupled6DPostCrackTrust_2026-11-24';out.mkdir(parents=True,exist_ok=True)
    paths.insert(0,'src/Solver/CCoupledHybridFactorization.cls')
    for name,left,right in (('Candidate696.patch',parent,before),('Candidate697RankCorrection.patch',before,after)):
        parts=[]
        for path in paths:
            a=(left/path).read_text('utf-8-sig').splitlines(True);b=(right/path).read_text('utf-8-sig').splitlines(True)
            parts.extend(difflib.unified_diff(a,b,fromfile='a/'+path,tofile='b/'+path))
        (out/name).write_text(''.join(parts),encoding='utf-8',newline='\n')
    # Первичные журналы и сведения об исходниках сохраняются также в Git; частные книги остаются на D.
    files={}
    for name in ('Stage-TestPostCrackTrust.py','Stage-TestRankTrust.py','TestLeastSquaresTrust.bas.txt','CandidateUndampedRank.cls.txt',Path(__file__).name):
        target=out/name;target.write_bytes((here/name).read_bytes());files[target.name]=sha256(target.read_bytes()).hexdigest().upper()
    for label,root in (('696',before),('697',after)):
        for name in ('Native/LeastSquaresTrust.txt','Native/LeastSquaresTrust.txt.ProcessClosure.json','CompilePointClosure.json','PointTraceSourceComposition.json'):
            target=out/(label+'_'+Path(name).name);target.write_bytes((root/name).read_bytes());files[target.name]=sha256(target.read_bytes()).hexdigest().upper()
    for name in ('Native/PostCrackTension.txt','Native/PostCrackTension.txt.ProcessClosure.json','PostCrackCapacityProof.json','PostCrackFieldsProof.json'):
        target=out/('697_'+Path(name).name);target.write_bytes((after/name).read_bytes());files[target.name]=sha256(target.read_bytes()).hexdigest().upper()
    proof['evidence_files_sha256']=files
    proof['candidate_sources_sha256']={path:sha256((after/path).read_bytes()).hexdigest().upper() for path in paths}
    proof['candidate_workbook_sha256']=sha256((after/'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper()
    proof['private_root']=str(after)
    proof['private_parent_root']=str(parent)
    proof['released_files_sha256']={p:sha256((repo/p).read_bytes()).hexdigest().upper() for p in ('workbook/output/RC_Section_NDM.xlsm','workbook/output/VBA_All_Code.txt')}
    for target in (out/'Proof.json',after/'PostCrackTrustStageProof.json'):
        target.write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
    mechanical=json.loads((after/'PostCrackFieldsProof.json').read_bytes())
    states=mechanical['phases'].get('PostCrackContinuation.Accepted',{}).get('independent_equilibrium',{}).get('states',[])
    highest=format(max(v['lambda'] for v in states),'.15g') if states else 'не получена: история продолжения не выгружена до тайм-аута'
    physical_native=physical['native']
    outcome='тайм-аут 600 с; итоговые assertions не получены' if physical_native['timeout600'] else str(physical_native['passed'])+'/'+str(physical_native['failed'])
    report=f'''# Регуляризованная посттрещинная поправка: точка для запрошенной паузы

Пользователь попросил завершить текущий этап и поставить цель на паузу.
Этап ограничен опытами 696–697, переданными на выполнение до просьбы;
дополнительные опыты после неё не добавлялись. Аналитическая проверка
и контроль сечения 697 исполняются последовательно одним запуском.
Основное ТЗ остаётся незавершённым.

## Найденная ошибка и аналитическая проверка

Существующий `TrySolveLeastSquaresTrust` проверен на известных корнях,
отсутствии корня, недоступном отклике, ошибках VBA 6/11, отмене,
неверных размерах и общих бюджетах. Частная версия 696 дополнительно
использует обработчик арифметически невычислимой пробной невязки.
Она дала 19/1: вырожденная линейная система с существующими корнями
не решилась. Исходный тест сохранён без ослабления ожиданий.

В `TryLeastSquaresRadius` отсутствовало принятие внутренней поправки
Гаусса—Ньютона при неполном ранге. Вместо неё начинался поиск
регуляризации для границы радиуса. Кандидат 697 добавляет базовое
решение по ведущему рангу переставленной R, с нулевыми свободными
координатами. Это не объявляется решением минимальной нормы.
Допустимая внутренняя поправка принимается без принудительного
растяжения до радиуса. Дополнительно учтён внутренний регуляризованный
шаг при нулевой нижней границе параметра, как в первичном `lmpar`.
Источник уже сохранён: `literature/Coupled6D/NumericalControl/Minpack/src/lmpar.c`.

Тот же аналитический набор в 697 дал **20/0**. Шесть остальных
подготовленных исходников, включая оба тестовых компонента, побитно
совпадают между 696 и 697; отличается только факторизация.
Фактические VBA-проекты обеих книг проверены: 255 компонентов,
соответственно шесть и семь импортов. Собственные процессы закрыты.

## Реальный бетон и арматура

Контроллер частной версии включает полный LM-резерв после установленного
локального максимума CDP. Положительная опорная lambda и специальный
переход после H сохраняют свою отдельную ветвь. Материал, силовые
допуски, 30 первичных итераций и общий бюджет 96 проб не меняются.
Начало состоит из тех же двух дотрещинных состояний, что в 692;
двоичные поля, параметры и нагрузки сверены точно.

Нативный поиск предела 697: **{outcome}**.
Наибольшая принятая lambda: **{highest}**.
Известный одноосный максимум: **{physical['known_maximum_lambda']:.15g}**.
Подтверждение данного частного предела: **{physical['particular_uniaxial_capacity_passed']}**.
Независимо проверены слабое равновесие, связи и шесть усилий всех
выгруженных состояний; при тайм-ауте это только две начальные точки,
а история автоматического продолжения остаётся непроверенной.
Интегрирование CDP остаётся нативным, поэтому
эта проверка не является независимой реализацией закона материала.
Полная предписанная материальная кривая в этом запуске не проверялась.

## Сохранность и продолжение

Рабочий код контрольной точки `02622b5` не заменён частными кандидатами.
Выпускная книга и `VBA_All_Code.txt` сохраняют прежние SHA-256.
`main` и наблюдение `ndm` не менялись. Полные долгие наборы не запускались.

После возобновления сначала проверить исправление на исходных 40
линейных случаях с независимым SVD и добавить случаи внутреннего
вырожденного шага. Затем завершить полную посттрещинную кривую и
поиск предела, проверить холодный HardShear. ONQ-008, все LoadPath,
внешние контрольные расчёты, выпускная книга, CAD, полная приёмка
и целевое время всего сочетания остаются открытыми. Пауза не означает
выполнения ТЗ.

Доказательства, точные журналы, контрольные суммы и патчи двух
кандидатов: [Proof.json](evidence/Coupled6DPostCrackTrust_2026-11-24/Proof.json).
Полные частные книги и исходные поля сохранены на D в папках 696–697.
'''
    (repo/'docs/reports/Coupled6DPostCrackTrustPause_2026-11-24.md').write_text(report,encoding='utf-8')
    print('ANALYTIC RANK REPAIR VERIFIED20/0; PHYSICAL',physical['native'],'CAPACITY_CONFIRMED',physical['particular_uniaxial_capacity_passed'])


if __name__=='__main__':TestCheck(Path.cwd())
