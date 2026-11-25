"""ДЛЯ ТЕСТОВ. Встраивает подтверждённую дуговую операцию в ядро и контроллер; сначала проверяет одну операцию на точной33-точечной истории."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json
import re

here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent = base / 'Cold784AutomaticEventRefinement_20261125'
root = base / 'Cold785OperationalArcController_20261125'
assert json.loads((parent / 'ConfirmedArcWalkProof.json').read_bytes())['accepted_states'] == 53
run_path(str(here / 'Stage-TestEnergyNumerics.py'))['TestClone'](root, parent, 1e-6, 1, 1e-6, 0)
archive = root / 'OperationalArcCompilationAncestor'; archive.mkdir()
for name in ('CompilePoint.log', 'CompilePointClosure.json'): (root / name).rename(archive / name)
prep = json.loads((root / 'Preparation.json').read_bytes()); values = {}
name = 'src/Solver/CCoupledNewtonSolver.cls'; code = (root / name).read_text('utf-8-sig')
helper = (here / 'TestArcYieldBoundary.cls.txt').read_text('utf-8-sig')
assert code.endswith(helper)
core = helper.replace("' -------------------- ДЛЯ ТЕСТОВ --------------------\n' На фактически отклонённом полном шаге выбирает первое новое начало пластичности CDP и решает существующие уравнения события.",
    "' На отклонённом равновесном поле выбирает первое новое начало пластического нагружения CDP и решает его без принятия истории.")
core = core.replace('TestLocateArcYieldBoundary', 'TryLoadingYieldBoundary')
wrapper = '''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет тестовый вход к рабочему обнаружению события без отдельной реализации поиска.
Friend Function TestLocateArcYieldBoundary(ByRef baseLoad() As Double, ByRef direction() As Double, ByRef ends() As Double, _
        ByVal fromLambda As Double, ByVal toLambda As Double, ByVal absoluteTolerance As Double, ByVal relativeTolerance As Double, _
        ByVal maxIterations As Long, ByRef result() As Double, ByRef resultLambda As Double, ByRef eventInfo As Variant) As Boolean
    TestLocateArcYieldBoundary = TryLoadingYieldBoundary(baseLoad, direction, ends, fromLambda, toLambda, _
        absoluteTolerance, relativeTolerance, maxIterations, result, resultLambda, eventInfo)
End Function

' Возвращает модуль, используемый для общей нормы пространственного равновесия и управляющих координат.
Friend Property Get ReferenceModulus() As Double
    ReferenceModulus = mReferenceModulus
End Property

' Вычисляет допуск заданной нагрузки без изменения текущего пробного контекста или материалной истории.
Friend Function ControlledLoadTolerance(ByRef baseLoad() As Double, ByRef direction() As Double, ByVal factor As Double, _
        ByVal absoluteTolerance As Double, ByVal relativeTolerance As Double) As Double
    Dim load(1 To COMPONENT_COUNT) As Double, equations() As Double, norm As Double, i As Long
    For i = 1 To COMPONENT_COUNT: load(i) = baseLoad(i) + factor * direction(i): Next i
    mSection.LoadEquations load(1), load(2), load(3), load(4), load(5), load(6), equations
    For i = 1 To mGaugeStart - 1: norm = GeomMax(norm, Abs(equations(i))): Next i
    ControlledLoadTolerance = absoluteTolerance + relativeTolerance * norm
End Function
'''
code = code[:-len(helper)] + core + '\n' + wrapper
values[name] = code.replace('\n', '\r\n').encode('utf-8-sig')
name = 'src/Solver/CCoupledArcContinuation.cls'
values[name] = (here / 'CandidateArcContinuation.cls.txt').read_text('utf-8-sig').replace('\n','\r\n').encode('utf-8-sig')
name = 'src/Solver/CCoupledDeformationContinuation.cls'; code = (root / name).read_text('utf-8-sig')
anchor = 'Private mGuideRecoveryProbes As Long, mGuideRecoverySolutions As Long'
position = code.index('\n', code.index(anchor))
code = code[:position] + '\nPrivate mSignedArcRevision As Long, mSignedArcAttempts As Long, mSignedArcRatio As Double\nPrivate mSignedArcProbes As Long, mSignedArcRecovered As Long' + code[position:]
anchor = '    mGuideRecoveryProbes = 0: mGuideRecoverySolutions = 0\n'; assert code.count(anchor) == 1
code = code.replace(anchor, anchor + '    mSignedArcRevision = -1: mSignedArcAttempts = 0: mSignedArcRatio = 1#: mSignedArcProbes = 0: mSignedArcRecovered = 0\n', 1)
anchor = '    Dim targetCoordinate As Double\n'; assert code.count(anchor) == 1
code = code.replace(anchor, '    Dim arcRecovered As Boolean\n' + anchor, 1)
code = code.replace('        hybridRecovery = False: envelopeRecovered = False: departureRecovered = False\n',
    '        hybridRecovery = False: envelopeRecovered = False: departureRecovered = False: arcRecovered = False\n', 1)
anchor = '        If increment < minimum Then\n            If mUsesEnvelopeControl Then\n'; assert code.count(anchor) == 1
code = code.replace(anchor, '''        If increment < minimum Then
            If Not unresolvedSteel Then arcRecovered = TryArcContinuationProbe(history, current, trial, nominal, increment, minimum)
            If arcRecovered Then
                lastProbeFailed = False: lastNoProgress = False: lastLoadUnresolved = False: smoothSteps = 0
            ElseIf mUsesEnvelopeControl Then
''', 1)
code = code.replace('        If envelopeRecovered Or departureRecovered Then\n', '        If envelopeRecovered Or departureRecovered Or arcRecovered Then\n', 1)
code = code.replace('        changedControl = envelopeRecovered: smoothTrial = False\n', '        changedControl = envelopeRecovered Or arcRecovered: smoothTrial = False\n', 1)
# Явная проверка стали остаётся ниже; дуга не должна немедленно заменяться ещё одной непроверенной H-пробой.
code = code.replace('        If probed And reinforced And Not envelopeRecovered And Not mUsesEnvelopeControl Then\n',
    '        If probed And reinforced And Not envelopeRecovered And Not arcRecovered And Not mUsesEnvelopeControl Then\n', 1)
anchor = 'Private Function Details() As String\n'; assert code.count(anchor) == 1
code = code.replace(anchor, anchor + '    If mSignedArcProbes > 0 Then Details = "; вызовов проверенной знаковой дуги " & CStr(mSignedArcProbes) & "; принятых направлений " & CStr(mSignedArcRecovered)\n', 1)
# Начальное присваивание Details внутри существующего метода не должно стирать новые счётчики.
start = code.index(anchor); end = code.index('End Function', start)
method = code[start:end]
assignments = list(re.finditer(r'(?m)^    Details = (?!Details)', method))
assert len(assignments) == 1
at = start + assignments[0].start()
code = code[:at] + code[at:].replace('    Details = ', '    Details = Details & ', 1)
code += '\n' + (here / 'CandidateArcController.cls.txt').read_text('utf-8-sig') + '\n'
# Variant из Collection сначала переносится в массив: ByRef Double() не принимает выражение Item.
code = code.replace('    If Not mSolver.SpatialIncrementAdvances(history(current.HistoryIndex - 1), accepted, field, nominal, increment) Then',
    '    Dim previous() As Double: previous = history(current.HistoryIndex - 1)\n    If Not mSolver.SpatialIncrementAdvances(previous, accepted, field, nominal, increment) Then', 1)
values[name] = code.replace('\n', '\r\n').encode('utf-8-sig')
name = 'tests/modTestCoupledElasticState.bas'; code = (root / name).read_text('utf-8-sig')
assert code.count('    TestRunArcYieldSplit\n') == 1
code = code.replace('    TestRunArcYieldSplit\n', '    TestRunOperationalArc\n', 1)
code += '''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Вызывает только рабочую дуговую операцию; положительный результат экспортируется для независимой сборки равновесия.
Private Sub TestRunOperationalArc()
    Dim arc As New CCoupledArcContinuation, solved As Boolean, factor As Double, projection() As Double
    Dim revision As Long: revision = mReverseSolverTest.AcceptedRevision
    arc.Initialize mReverseSolverTest, mReverseBaseTest, mReverseDirectionTest, mReverseOptionsTest(3), mReverseOptionsTest(4), CLng(mReverseOptionsTest(2))
    mReverseSolverTest.DiscardStep
    solved = arc.TryStep(mReverseFieldsTest, mReverseTraceTest, 2#, 6, factor, projection)
    If mReverseSolverTest.AcceptedRevision <> revision Or arc.Calls > 6 Then Err.Raise 5, , "Рабочая дуга нарушила историю или бюджет."
    Dim certificate As Variant: certificate = arc.Verification
    If IsEmpty(certificate) Then Err.Raise 5, , "Рабочая дуга не сформировала проверочный контекст."
    Dim handle As Integer: handle = FreeFile
    Open mPolishRootTest & "\\Native\\OperationalArc.Result.bin" For Binary Access Write As #handle
    Dim magic As String * 16: magic = "NDMOPERARC000001"
    Put #handle, , magic
    Dim i As Long
    For i = LBound(certificate) To UBound(certificate): TestWriteReverseBinaryVector handle, certificate(i): Next i
    Dim field() As Double, actual() As Double, effort As Variant
    If solved Then
        mReverseSolverTest.CopyTrial field, actual
        TestWriteReverseBinaryVector handle, field: TestWriteReverseBinaryVector handle, actual: TestWriteReverseBinaryVector handle, projection
    End If
    effort = mReverseSolverTest.ReadPolishEffortTest: TestWriteReverseBinaryVector handle, effort
    Close #handle
    handle = FreeFile
    Open mPolishRootTest & "\\Native\\OperationalArc.Outcome.txt" For Output As #handle
    Print #handle, "SOLVED;" & CStr(solved) & ";LAMBDA;" & CStr(factor) & ";CALLS;" & CStr(arc.Calls)
    Print #handle, arc.Details
    Close #handle
    mReverseSolverTest.DiscardStep
    If solved Then TestExportSignedArcTrial "OperationalArc.Full", field, factor
End Sub
'''
values[name] = code.replace('\n', '\r\n').encode('utf-8-sig')
imports = ['CCoupledArcContinuation','CCoupledNewtonSolver','CCoupledDeformationContinuation','modTestCoupledElasticState']
compiler = (here / 'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
values.update({'Compile-TestExternalPoint.ps1': compiler.replace("-ModuleNames 'modTestCoupledElasticState'", '-ModuleNames @(' + ','.join(repr(n) for n in imports) + ')', 1).replace('\n','\r\n').encode('utf-8-sig'),
    'ParentCodeUpdate.json': (parent / 'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json': (parent / 'CompilePointClosure.json').read_bytes(), Path(__file__).name: Path(__file__).read_bytes()})
for name, data in values.items():
    if name.endswith(('.cls','.bas')): assert max(map(len,data.decode('utf-8-sig').splitlines())) < 1024
    (root / name).write_bytes(data); prep['inputs'][name] = sha256(data).hexdigest().upper()
prep['modules'].append('CCoupledArcContinuation')
prep['directed_scope'] = dict(case='OperationalSignedArcAndStalledController', actual_import_this_pass=imports,
    actual_component_count=256, source33_exact784=True, first_test_calls_operational_class_only=True,
    controller_retains_original_steel_and_physical_status_acceptance=True, global96_call_budget_preserved=True,
    pair_and_event_each_share30_primary_budget=True, cold_search_not_yet_run=True,
    no_working_promotion=True, physical_capacity_certified=False)
(root / 'Preparation.json').write_text(json.dumps(prep, indent=2, sort_keys=True) + '\n', encoding='utf-8')
print('785 OPERATIONAL ARC CLASS/NEWTON API/STALLED CONTROLLER;4 IMPORTS/256 COMPONENTS; WARM SINGLE OPERATION THEN COLD REQUIRED')
