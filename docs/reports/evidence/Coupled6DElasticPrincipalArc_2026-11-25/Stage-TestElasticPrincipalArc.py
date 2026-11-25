"""ДЛЯ ТЕСТОВ. Подключает проверенную упругую спектральную границу к общей дуге и продолжает точную историю86 без задания будущих полей."""
from pathlib import Path
from hashlib import sha256
from runpy import run_path
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold837ElasticPrincipalEvent_20261125';root=base/'Cold838ElasticPrincipalArc_20261125'
assert json.loads((parent/'ElasticPrincipalEventProof.json').read_bytes())['solved_events']==6
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'ElasticPrincipalArcCompilationAncestor';archive.mkdir()
for n in ('CompilePoint.log','CompilePointClosure.json'):(root/n).rename(archive/n)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
def TestReplace(code,old,new):
    """Сохраняет точный контракт единственной ожидаемой замены."""
    assert code.count(old)==1,old[:160]
    return code.replace(old,new,1)
name='src/Solver/CCoupledNewtonSolver.cls';code=(root/name).read_text('utf-8-sig')
start=code.index('Friend Function TryElasticPrincipalBoundary(');stop=code.index('\nEnd Function',start)
method=code[start:stop]
begin=method.index('    Dim strain(1 To 6) As Double');end=method.index('    Next point\n',begin)+len('    Next point\n')
loop=method[begin:end]
method=method[:begin]+'    SelectElasticPrincipalStation ends, eventRadius, resolution, station, nearest, selected\n'+method[end:]
code=code[:start]+method+code[stop:]
code+='''
' Перечисляет упругие спектральные переходы вместе с уже выбранным пластическим событием; сохраняет самое раннее разрешимое.
' Все уровни читаются на одной принятой истории; это выбор численной границы, а не проверка разрушения.
Private Sub SelectElasticPrincipalStation(ByRef ends() As Double, ByVal eventRadius As Double, ByVal resolution As Double, _
        ByRef station As TYieldStation, ByRef nearest As Double, ByRef selected As Boolean)
    Dim candidate As TYieldStation, fraction As Double, point As Long, side As Long, principal As Long
'''+loop+'''End Sub
'''
anchor='    eventInfo = Array(Abs(CLng(selected)), station.Point, station.Side, station.FromValue, station.ToValue, nearest, mAcceptedRevision)'
code=TestReplace(code,anchor,'    SelectElasticPrincipalStation ends, eventRadius, resolution, station, nearest, selected\n'+anchor+'''
    If station.ElasticPrincipal Then _
        eventInfo = Array(Abs(CLng(selected)), station.Point, station.Side, station.FromValue, station.ToValue, nearest, mAcceptedRevision, station.PrincipalIndex, 1#)''')
code=TestReplace(code,"' На отклонённом равновесном поле выбирает первое новое начало пластического нагружения CDP и решает его без принятия истории.",
    "' На отклонённом поле выбирает первое разрешимое начало пластического нагружения или упругий спектральный переход CDP; историю не принимает.")
values[name]=code
name='src/Solver/CCoupledArcContinuation.cls';code=(root/name).read_text('utf-8-sig')
code=TestReplace(code,"' допускается отдельное решение первого начала пластического нагружения.",
    "' допускается отдельное решение начала пластического нагружения или упругой спектральной границы CDP.")
code=TestReplace(code,'beforeSurface = mSolver.TryReadLoadingSurface(CLng(eventInfo(1)), CLng(eventInfo(2)), half, halfYield)',
    'beforeSurface = ReadEventSurface(eventInfo, half, halfYield)')
code=TestReplace(code,'And beforeSurface And halfYield < 0# Then','And beforeSurface And IsBeforeEvent(eventInfo, halfYield) Then')
code=TestReplace(code,'If Not mSolver.TryReadLoadingSurface(CLng(eventInfo(1)), CLng(eventInfo(2)), candidate, candidateYield) Then verified = False',
    'If Not ReadEventSurface(eventInfo, candidate, candidateYield) Then verified = False')
code=TestReplace(code,'If Not mSolver.TryReadLoadingSurface(CLng(eventInfo(1)), CLng(eventInfo(2)), probe, probeYield) Then verified = False',
    'If Not ReadEventSurface(eventInfo, probe, probeYield) Then verified = False')
code=TestReplace(code,'verified = verified And candidateYield < 0# And probeYield < 0#',
    'verified = verified And IsBeforeEvent(eventInfo, candidateYield) And IsBeforeEvent(eventInfo, probeYield)')
code+='''
' Читает именно выбранную численную границу, сохраняя семипольный контракт пластического события.
Private Function ReadEventSurface(ByRef eventInfo As Variant, ByRef field() As Double, ByRef value As Double) As Boolean
    If UBound(eventInfo) >= 8 Then
        If eventInfo(8) = 1# Then
            ReadEventSurface = mSolver.TryReadElasticPrincipalSurface(CLng(eventInfo(1)), CLng(eventInfo(2)), CLng(eventInfo(7)), field, value)
            Exit Function
        End If
    End If
    ReadEventSurface = mSolver.TryReadLoadingSurface(CLng(eventInfo(1)), CLng(eventInfo(2)), field, value)
End Function

' Проверяет подход с исходной стороны: пластическая поверхность отрицательна, спектральная координата сохраняет исходный знак.
Private Function IsBeforeEvent(ByRef eventInfo As Variant, ByVal value As Double) As Boolean
    If UBound(eventInfo) >= 8 Then
        If eventInfo(8) = 1# Then
            IsBeforeEvent = value * CDbl(eventInfo(3)) > 0#
            Exit Function
        End If
    End If
    IsBeforeEvent = value < 0#
End Function
''';values[name]=code
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
code=TestReplace(code,'    TestReadElasticPrincipalEvent\n',
    '    TestRunContinuedOutgoingGuard\n    TestRunAutomaticLoadingWalk\n    TestRecentLoadingResetGuard\n')
code=code.replace('Исходный69-точечный путь неожиданно заканчивается новой границей нагружения.',
    'Контрольный путь неожиданно заканчивается новой границей нагружения.')
values[name]=code
imports=['CCoupledNewtonSolver','CCoupledArcContinuation','modTestCoupledElasticState']
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
values['Compile-TestExternalPoint.ps1']=TestReplace(compiler,"-ModuleNames 'modTestCoupledElasticState'",'-ModuleNames @('+','.join(repr(n) for n in imports)+')')
values.update({'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for n,data in values.items():
    if isinstance(data,str):
        if n.endswith(('.cls','.bas')):assert max(map(len,data.splitlines()))<1024
        data=data.replace('\n','\r\n').encode('utf-8-sig')
    (root/n).write_bytes(data);prep['inputs'][n]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='IntegratedCDPElasticPrincipalArcFrom86',actual_import_this_pass=sorted(imports),actual_component_count=256,
    material_and_scalar_constraint_exact837=True,negative_yield_required_at_endpoints_and_solved_face=True,
    earliest_resolvable_material_event_selected=True,maximum_attempts32=True,maximum_calls96=True,
    full_half_and_lambda_and_direction_criteria_unchanged=True,existing30_iteration_phase_budgets=True,
    no_prescribed_future_field=True,not_cold_capacity_search=True,physical_capacity_certified=False,no_working_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print(root)
