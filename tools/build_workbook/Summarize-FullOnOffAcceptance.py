"""Проверяет полноту обоих режимов по журналам, хешам и завершению каждой suite."""
from __future__ import annotations
import argparse
import hashlib
import importlib.util
import json
import re
import subprocess
from datetime import datetime
from pathlib import Path

FULL = (
    'modTestGeometry.RunGeometryTests',
    'modTestMaterialDiagrams.RunMaterialDiagramTests',
    'modTestSectionSolver.RunSectionSolverTests',
    'modTestCapacitySolver.RunCapacitySolverTests',
    'modTestCrackWidth.RunCrackWidthTests',
    'modTestBatchCalculation.RunBatchCalculationTests',
    'modTestWorkbookInterface.RunWorkbookInterfaceTests',
    'modTestRegressionBaseline.RunRegressionBaselineTests',
)
GROUP_COUNTS = {'Full': 8, 'Geometry': 26, 'SP35': 6, 'Recipes': 26}
CURRENT_GROUP_COUNTS = {**GROUP_COUNTS, 'Geometry': 30, 'SavedOffsets': 2}
INDIVIDUAL_ROW_GROUP_COUNTS = {**CURRENT_GROUP_COUNTS, 'Geometry': 36}
INDIVIDUAL_ROW_MACROS = {
    'modTestIndividualContours.RunIndividualContourTests',
    'modTestIndividualContours.RunIndividualContourFailureTests',
    'modTestIndividualContours.RunIndividualContourWorkbookTests',
    'modTestIndividualContours.RunUserRoundedContourTests',
    'modTestIndividualContours.RunOffsetNeighborMergeTests',
    'modTestIndividualContours.RunIndividualContourPerformanceTests',
}
SNAPSHOT_ACCELERATION_GROUP_COUNTS = {**INDIVIDUAL_ROW_GROUP_COUNTS, 'CurrentFeatures': 3}
SNAPSHOT_ACCELERATION_MACROS = {
    'modTestContourMesh.RunAutomaticMeshTests',
    'modTestCalculationSnapshot.RunCalculationSnapshotTests',
    'modTestDiscreteForces.RunDiscreteForceTests',
}

MESSAGE_PRESENTATION_GROUP_COUNTS = {**INDIVIDUAL_ROW_GROUP_COUNTS, 'CurrentFeatures': 6}
MESSAGE_PRESENTATION_MACROS = SNAPSHOT_ACCELERATION_MACROS | {
    'modTestAudit02Api.RunProgramMessageTests',
    'modTestAudit02Api.RunProgramMessagePresentationTests',
    'modTestAudit02Api.RunInputOnlyDialogTests',
}

CONTOUR_ELEMENT_GROUP_COUNTS = {**INDIVIDUAL_ROW_GROUP_COUNTS, 'CurrentFeatures': 11}
CONTOUR_ELEMENT_MACROS = MESSAGE_PRESENTATION_MACROS | {
    'modTestContourMesh.RunAdaptiveBoundaryMeshTests',
    'modTestElementGeometry.RunElementGeometryTests',
    'modTestElementGeometry.RunContourCellMeshTests',
    'modTestCalculationSnapshot.RunConcreteGeometrySnapshotTests',
    'modTestElementGeometryViews.RunElementGeometryViewTests',
}

ACTION_COMPLETION_GROUP_COUNTS = {**CONTOUR_ELEMENT_GROUP_COUNTS, 'CurrentFeatures': 13}
ACTION_COMPLETION_MACROS = CONTOUR_ELEMENT_MACROS | {
    'modTestCalculationSnapshot.RunGeometryHashTests',
    'modTestCalculationSnapshot.RunWorkbookActionCompletionTests',
}


COUPLED_6D_TOTALS = {
    'modTestWorkbookInterface.RunTestCoupledLoadInput': 'TOTAL_COUPLED_LOAD_INPUT',
    'modTestAutoCADConfig.RunTestCoupledFieldProfiles': 'TOTAL_COUPLED_FIELD_PROFILES',
    'modTestCoupledElasticState.RunTestCoupledContinuation': 'TOTAL_COUPLED_CONTINUATION',
    'modTestCoupledElasticState.RunTestCoupledLoadMatrix': 'TOTAL_COUPLED_LOAD_MATRIX',
    'modTestCoupledElasticState.RunTestCoupledBatchRoute': 'TOTAL_COUPLED_BATCH_ROUTE',
    'modTestCoupledElasticState.RunTestCoupledTraditionalChecks': 'TOTAL_COUPLED_TRADITIONAL',
    'modTestCoupledElasticState.RunTestCoupledWorkbookInput': 'TOTAL_COUPLED_WORKBOOK_INPUT',
    'modTestCoupledElasticState.RunTestCoupledHoopCapacity': 'TOTAL_COUPLED_HOOP_CAPACITY',
}
COUPLED_6D_GROUP_COUNTS = {**ACTION_COMPLETION_GROUP_COUNTS, 'Coupled6D': len(COUPLED_6D_TOTALS)}
COUPLED_NATIVE_MACRO = 'RunRealAutoCADCoupledTransverseTests'
COUPLED_NATIVE_TOTAL = 'TOTAL_NATIVE_COUPLED_TRANSVERSE'


def coverage_errors(manifest):
    """Исторические запуски читаются в прежнем объёме; новый полный набор обязателен целиком."""
    if manifest.get('CoverageDefinition') != 'AllActiveAcceptanceIncludingFocused':
        return []
    errors=[]
    version=manifest.get('CoverageVersion',1)
    counts_by_version = {
        1: GROUP_COUNTS, 2: CURRENT_GROUP_COUNTS, 3: INDIVIDUAL_ROW_GROUP_COUNTS,
        4: SNAPSHOT_ACCELERATION_GROUP_COUNTS, 5: MESSAGE_PRESENTATION_GROUP_COUNTS,
        6: CONTOUR_ELEMENT_GROUP_COUNTS, 7: ACTION_COMPLETION_GROUP_COUNTS, 8: COUPLED_6D_GROUP_COUNTS,
    }
    if type(version) is not int or version not in counts_by_version:
        return ['Unsupported full coverage version']
    counts=counts_by_version[version]
    if manifest.get('IncludesFocusedAcceptance') is not True:
        errors.append('The current full definition requires focused acceptance')
    expected={(group,mode) for group in counts for mode in ('No','Yes')}
    jobs=manifest.get('Jobs',[])
    actual=[(job['Group'],job['Mode']) for job in jobs]
    if set(actual)!=expected or len(actual)!=len(expected):
        errors.append('The full plan requires every declared canonical ON/OFF group')
    for job in jobs:
        if job.get('ExpectedSuites')!=counts.get(job['Group']):
            errors.append('Canonical suite coverage differs: '+job['Group'])
        if version >= 3 and job['Group'] == 'Geometry':
            macros = job.get('Macro', [])
            if len(macros) != counts['Geometry'] or len(macros) != len(set(macros)) or not INDIVIDUAL_ROW_MACROS <= set(macros):
                errors.append('Individual row coverage is incomplete or duplicated')
        if version >= 4 and job['Group'] == 'CurrentFeatures':
            macros = job.get('Macro', [])
            required = ACTION_COMPLETION_MACROS if version >= 7 else CONTOUR_ELEMENT_MACROS if version == 6 else MESSAGE_PRESENTATION_MACROS if version == 5 else SNAPSHOT_ACCELERATION_MACROS
            if len(macros) != len(required) or set(macros) != required:
                errors.append('Every feature declared for this coverage version is mandatory')
        if job['Group'] == 'Coupled6D' and tuple(job.get('Macro', [])) != tuple(COUPLED_6D_TOTALS):
            errors.append('Every declared coupled 6D suite is mandatory in both modes')
    return errors


def runner_coverage_errors(manifest, runner_text):
    """Связывает новую версию с реально сохранённым runner, не меняя исторические планы."""
    versions = {int(value) for value in re.findall(r'\bCoverageVersion\s*=\s*(\d+)\b', runner_text)}
    declared = manifest.get('CoverageVersion', 1)
    if declared not in (6,7,8) and not any(version >= 6 for version in versions):
        return []
    errors = []
    if versions != {declared}:
        errors.append('Full runner snapshot and manifest coverage versions differ')
    if manifest.get('CoverageDefinition') != 'AllActiveAcceptanceIncludingFocused':
        errors.append(f'The version {declared} runner requires the complete current coverage definition')
    return errors

def has_exact_pass_total(text, label):
    """Требует единственный собственный итог с выполненными проверками."""
    totals=re.findall(r'(?m)^TOTAL[^\r\n]*',text)
    return (len(totals)==1
            and re.fullmatch(re.escape(label)+r': passed=[1-9]\d*; failed=0',totals[0]) is not None
            and re.search(r'(?m)^FAIL(?::| )|RUNTIME ERROR',text) is None)


def coupled_totals_errors(text):
    """Привязывает итог каждого макроса к его секции журнала."""
    errors=[]
    for block in re.split(r'(?m)^SUITE_STARTED: ',text)[1:]:
        name=block.partition(';')[0]
        if name not in COUPLED_6D_TOTALS or not has_exact_pass_total(block,COUPLED_6D_TOTALS[name]):
            errors.append('Coupled suite lacks its exact successful total: '+name)
    return errors


def coupled_native_errors(directory, stage_result, source):
    """Подтверждает отдельный нативный этап по точному журналу и входной книге."""
    if not isinstance(stage_result,dict) or stage_result.get('Passed') is not True or stage_result.get('ExitCode') != 0:
        return ['Mandatory coupled native CAD stage did not finish successfully']
    folder=directory/'CoupledNativeCAD'
    case=folder/COUPLED_NATIVE_MACRO
    def read(path):
        return json.loads(path.read_text(encoding='utf-8-sig'))
    def sha(path):
        return hashlib.sha256(path.read_bytes()).hexdigest()

    try:
        stage_path=folder/'Stage.json'
        stage=read(stage_path)
        native_runner=sha(folder/'NativeRunner.ps1')
        invocation=read(case/'NativeInvocation.json')
        report=case/'NativeCAD.txt'
        report_hash=sha(report)
        results=stage.get('Results',[])
        if (sha(stage_path)!=str(stage_result.get('StageSHA256','')).lower()
                or stage.get('Scope')!='Coupled6D' or stage.get('Passed') is not True
                or stage.get('UserDocumentsUsed') is not False or len(results)!=1
                or results[0].get('Macro')!=COUPLED_NATIVE_MACRO or results[0].get('Passed') is not True
                or str(stage.get('RunnerSHA256','')).lower()!=native_runner
                or str(results[0].get('ReportSHA256','')).lower()!=report_hash
                or invocation.get('Macro')!=COUPLED_NATIVE_MACRO or invocation.get('Passed') is not True
                or invocation.get('SourceUnchanged') is not True
                or str(invocation.get('SourceSHA256','')).lower()!=source
                or str(invocation.get('RunnerSHA256','')).lower()!=native_runner
                or not invocation.get('OwnedCadPID') or invocation.get('OwnedCadPID')!=invocation.get('LaunchedCadPID')
                or not has_exact_pass_total(report.read_text(encoding='utf-8-sig'),COUPLED_NATIVE_TOTAL)):
            return ['Coupled native CAD report, source workbook, or ownership evidence differs']
    except (OSError, UnicodeError, ValueError, AttributeError, TypeError):
        return ['Coupled native CAD evidence is missing or unreadable']
    return []


def load_logs():
    path=Path(__file__).with_name('Summarize-GeometryFull.py')
    spec=importlib.util.spec_from_file_location('full_logs',path)
    module=importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module

def summarize(directory):
    logs=load_logs()
    manifest_path=directory/'Run.json'
    manifest=json.loads(manifest_path.read_text(encoding='utf-8-sig'))
    source=manifest['SourceSHA256'].lower()
    jobs=[]
    errors=coverage_errors(manifest)
    completed={(j['Group'],j['Mode']):j for j in manifest['CompletedJobs']}
    for job in manifest['Jobs']:
        path=directory/(job['Group']+'_'+job['Label']+'.txt')
        report=logs.inspect(path,source,job['ExpectedSuites'])
        expected=FULL if job['Group']=='Full' else tuple(job['Macro'])
        actual=tuple(s['name'] for s in report['suites'])
        text=path.read_text(encoding='utf-8-sig') if path.exists() else ''
        report['group']=job['Group']
        report['mode']=job['Mode']
        report['expectedMacros']=expected
        report['printedSuccessfulAssertions']=len(re.findall(r'(?m)^(?:OK(?::|\|)|PASS )',text))
        report['failures']=re.findall(r'(?m)^FAIL(?::| ).*$',text)
        if actual!=expected:
            report.setdefault('errors',[]).append('Suite names/order differ from the declared plan')
            report['status']='INCOMPLETE_OR_FAILED'
        expected_mode='GLOBAL_MODE: General.DiagramExtension='+job['Mode']+';'
        if expected_mode not in text:
            report.setdefault('errors',[]).append('Requested global mode is not confirmed')
            report['status']='INCOMPLETE_OR_FAILED'
        finish=completed.get((job['Group'],job['Mode']))
        if not finish or finish['ExitCode']!=0:
            report.setdefault('errors',[]).append('Successful driver completion is missing')
            report['status']='INCOMPLETE_OR_FAILED'
        if 'WATCHDOG_FAILURE:' in text or re.search(r'(?m)^(?:SCRIPT ERROR|CLEANUP ERROR|RUN_STOPPED):',text):
            report.setdefault('errors',[]).append('Runner failure or interruption')
            report['status']='INCOMPLETE_OR_FAILED'
        if job['Group']=='Coupled6D':
            invalid_totals=coupled_totals_errors(text)
            if invalid_totals:
                report.setdefault('errors',[]).extend(invalid_totals)
                report['status']='INCOMPLETE_OR_FAILED'
        if report['status']!='PASS':errors.append(path.name)
        jobs.append(report)
    if manifest['Status']!='COMPLETED_PENDING_LOG_REVIEW':errors.append('Driver is not successfully complete')
    if manifest.get('SourceChanged'):errors.append('Frozen input changed')
    if manifest.get('RunnerChanged'):errors.append('Test runners changed during execution')
    coupled_native=manifest.get('CoupledNativeCAD')
    if manifest.get('CoverageVersion')==8:
        errors.extend(coupled_native_errors(directory,coupled_native,source))
    focused = manifest.get('FocusedAcceptance')
    if manifest.get('IncludesFocusedAcceptance'):
        if not focused or not focused.get('Passed') or focused.get('ExitCode') != 0:
            errors.append('Mandatory focused acceptance did not finish successfully')
        else:
            focused_path=Path(__file__).resolve().parents[2]/focused['Directory']/'Acceptance.json'
            if not focused_path.is_file():
                errors.append('Focused acceptance evidence is missing')
            else:
                accepted=json.loads(focused_path.read_text(encoding='utf-8-sig'))
                if (hashlib.sha256(focused_path.read_bytes()).hexdigest()!=focused.get('AcceptanceSHA256','').lower()
                    or accepted.get('Status')!='PASS' or accepted.get('SourceSHA256','').lower()!=source
                    or accepted.get('IndividualCases')!=122):
                    errors.append('Focused acceptance content or provenance differs')
    for name,key in (('FullRunner.ps1','RunnerSHA256'),('ModeRunner.ps1','ModeRunnerSHA256')):
        snapshot=directory/name
        if not snapshot.is_file() or hashlib.sha256(snapshot.read_bytes()).hexdigest()!=manifest.get(key,'').lower():
            errors.append('Executed runner snapshot does not match: '+name)
        if name == 'FullRunner.ps1' and snapshot.is_file():
            errors.extend(runner_coverage_errors(manifest, snapshot.read_text(encoding='utf-8-sig')))
    finished=manifest.get('FinishedUTC')
    if not finished or datetime.fromisoformat(finished)>datetime.fromisoformat(manifest['Deadline']):
        errors.append('Completion before the requested deadline is not confirmed')
    frozen=Path(manifest['FrozenWorkbook'])
    provenance='frozen file'
    if frozen.exists():
        input_bytes=frozen.read_bytes()
    else:
        # После очистки собственных временных копий точный вход восстанавливается из коммита.
        repo=Path(__file__).resolve().parents[2]
        try:
            relative=Path(manifest['SourceWorkbook']).relative_to(repo).as_posix()
            input_bytes=subprocess.check_output(['git','show',manifest['SourceCommit']+':'+relative],cwd=repo)
            provenance='committed workbook'
        except (ValueError,subprocess.CalledProcessError):
            input_bytes=b''
    if hashlib.sha256(input_bytes).hexdigest()!=source:
        errors.append('Frozen or committed workbook hash does not match the manifest')
    total_suites=sum(len(j['suites']) for j in jobs)
    expected_suites=sum(j['ExpectedSuites'] for j in manifest['Jobs'])
    return {
        'status':'PASS' if not errors else 'INCOMPLETE_OR_FAILED',
        'sourceCommit':manifest['SourceCommit'], 'sourceSHA256':source,
        'runnerSHA256':manifest['RunnerSHA256'],
        'modeRunnerSHA256':manifest.get('ModeRunnerSHA256'),
        'coverageDefinition':manifest.get('CoverageDefinition','HistoricalCoreSuitesOnly'),
        'coverageVersion':manifest.get('CoverageVersion',1),
        'includesFocusedAcceptance':bool(manifest.get('IncludesFocusedAcceptance')),
        'focusedAcceptance':focused,
        'coupledNativeCAD':coupled_native,
        'inputProvenance':provenance,
        'startedUTC':manifest['StartedUTC'], 'finishedUTC':finished,
        'deadline':manifest['Deadline'], 'expectedSuites':expected_suites,
        'completedSuites':sum(bool(s.get('finished')) for j in jobs for s in j['suites']),
        'observedSuites':total_suites,
        'printedSuccessfulAssertions':sum(j['printedSuccessfulAssertions'] for j in jobs),
        'failedAssertions':sum(len(j['failures']) for j in jobs),
        'errors':errors, 'jobs':jobs,
    }

def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--directory',type=Path,required=True)
    parser.add_argument('--report',type=Path)
    parser.add_argument('--check',action='store_true')
    args=parser.parse_args()
    result=summarize(args.directory)
    output=args.report or args.directory/'Summary.json'
    output.write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    print(json.dumps({k:v for k,v in result.items() if k!='jobs'},ensure_ascii=False))
    if args.check and result['status']!='PASS':raise SystemExit(1)

if __name__=='__main__':main()
