"""ДЛЯ ТЕСТОВ. Подготавливает общий контроль предельных деформаций стали трёх станций без изменения диаграмм и численных допусков."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent
base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold878CriticalHRecoveryPairs_20261126'
root=base/'Cold881SpatialSteelLimits_20261126'
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'SpatialSteelCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
prep=json.loads((root/'Preparation.json').read_bytes());values={}

def replace(code,old,new):
    """Проверяет точный контракт заменяемого фрагмента."""
    assert code.count(old)==1,(old[:100],code.count(old))
    return code.replace(old,new,1)

name='src/Materials/CCoupledReinforcedResponse.cls';code=(root/name).read_text('utf-8-sig')
code=replace(code,'Private mConcrete As ICoupledPointResponse','''Private mConcrete As ICoupledPointResponse
Private mSteelPositive As CCoupledReinforcedResponse, mSteelNegative As CCoupledReinforcedResponse
Private mSteelHalfStep As Double''')
code=replace(code,'    Set mConcrete = Nothing','''    Set mConcrete = Nothing
    Set mSteelPositive = Nothing: Set mSteelNegative = Nothing: mSteelHalfStep = 0#''')
anchor="' Показывает наличие продольной или поперечной арматуры в балансе сечения."
methods='''
' Подключает только читающий контроль двух независимых боковых историй.
' Принятие и сброс каждой истории остаются обязанностью владельца решателя.
Friend Sub AttachSteelStations(ByVal positive As CCoupledReinforcedResponse, _
        ByVal negative As CCoupledReinforcedResponse, ByVal halfStep As Double)
    If (positive Is Nothing) Xor (negative Is Nothing) Then Err.Raise 5, , "Для контроля стали требуются обе боковые станции."
    If Not positive Is Nothing Then
        If positive Is Me Or negative Is Me Or positive Is negative Or halfStep <= 0# Then _
            Err.Raise 5, , "Для контроля стали нужны независимые станции и положительный шаг."
    End If
    Set mSteelPositive = positive: Set mSteelNegative = negative: mSteelHalfStep = halfStep
End Sub

' Возвращает максимум заданной группы без повторного расчёта материала.
' При равенстве приоритет имеет центральная станция; csuNone означает всю сталь.
Friend Function ReadSteelLimit(ByVal accepted As Boolean, ByVal family As ECoupledSteelUncertainty, _
        ByRef point As Long, ByRef side As Long) As Double
    Dim localPoint As Long, ratio As Double, station As Long, response As CCoupledReinforcedResponse
    side = 0: point = 0
    ReadSteelLimit = LocalSteelLimit(accepted, family, point)
    For station = -1 To 1 Step 2
        Set response = SteelStation(station)
        If Not response Is Nothing Then
            ratio = response.LocalSteelLimit(accepted, family, localPoint)
            If localPoint <> 0 And (point = 0 Or ratio > ReadSteelLimit) Then
                ReadSteelLimit = ratio: point = localPoint: side = station
            End If
        End If
    Next station
End Function

' Читает только собственную станцию, исключая рекурсивное объединение историй.
Friend Function LocalSteelLimit(ByVal accepted As Boolean, ByVal family As ECoupledSteelUncertainty, ByRef point As Long) As Double
    Select Case family
        Case csuNone
            If accepted Then
                LocalSteelLimit = mAcceptedSteelUtilization: point = mAcceptedSteelPoint
            Else
                LocalSteelLimit = mTrialSteelUtilization: point = mTrialSteelPoint
            End If
        Case csuLongitudinal
            If accepted Then
                LocalSteelLimit = mAcceptedLongitudinalUtilization: point = mAcceptedLongitudinalPoint
            Else
                LocalSteelLimit = mTrialLongitudinalUtilization: point = mTrialLongitudinalPoint
            End If
        Case csuTransverse
            If accepted Then
                LocalSteelLimit = mAcceptedTransverseUtilization: point = mAcceptedTransversePoint
            Else
                LocalSteelLimit = mTrialTransverseUtilization: point = mTrialTransversePoint
            End If
        Case Else: Err.Raise 5, , "Неизвестная группа контроля стали."
    End Select
End Function

' Возвращает материал выбранной станции; ноль соответствует центральной.
Private Function SteelStation(ByVal side As Long) As CCoupledReinforcedResponse
    Select Case side
        Case -1: Set SteelStation = mSteelNegative
        Case 0: Set SteelStation = Me
        Case 1: Set SteelStation = mSteelPositive
        Case Else: Err.Raise 5, , "Номер станции стали должен быть -1, 0 или 1."
    End Select
End Function

' Возвращает знаковое смещение станции в безразмерных продольных координатах.
Friend Function SteelStationOffset(ByVal side As Long) As Double
    If Abs(side) > 1 Or (side <> 0 And mSteelPositive Is Nothing) Then Err.Raise 5, , "Станция стали не подготовлена."
    SteelStationOffset = side * mSteelHalfStep
End Function

' Читает фактическое принятое напряжение именно выбранной станции.
Friend Function AcceptedStationSteelStress(ByVal point As Long, ByVal side As Long) As Double
    Dim response As CCoupledReinforcedResponse: Set response = SteelStation(side)
    If response Is Nothing Then Err.Raise 5, , "Боковая история стали не подготовлена."
    AcceptedStationSteelStress = response.AcceptedSteelStress(point)
End Function

'''
code=replace(code,anchor,methods+anchor)
anchor="' Возвращает контрольную точку принятого состояния по всей арматуре."
code=replace(code,anchor,''' ' Возвращает станцию точки максимального использования последней полной пробы.
Friend Property Get TrialSteelCriticalSide() As Long
    Dim point As Long, ratio As Double
    ratio = ReadSteelLimit(False, csuNone, point, TrialSteelCriticalSide)
End Property

'''.lstrip()+anchor)
for prop,accepted,family,local,ispoint in [
('TrialSteelUtilization',False,'csuNone','mTrialSteelUtilization',False),
('AcceptedSteelUtilization',True,'csuNone','mAcceptedSteelUtilization',False),
('TrialSteelCriticalPoint',False,'csuNone','mTrialSteelPoint',True),
('AcceptedSteelCriticalPoint',True,'csuNone','mAcceptedSteelPoint',True),
('AcceptedTransverseCriticalPoint',True,'csuTransverse','mAcceptedTransversePoint',True),
('AcceptedTransverseSteelUtilization',True,'csuTransverse','mAcceptedTransverseUtilization',False),
('AcceptedLongitudinalUtilization',True,'csuLongitudinal','mAcceptedLongitudinalUtilization',False),
('AcceptedLongitudinalPoint',True,'csuLongitudinal','mAcceptedLongitudinalPoint',True)]:
    new=f'    Dim point As Long, side As Long, ratio As Double\n    ratio = ReadSteelLimit({str(accepted)}, {family}, point, side)\n    {prop} = '+('point' if ispoint else 'ratio')
    code=replace(code,f'    {prop} = {local}',new)
# Проекция и оценка арифметической неопределённости используют ту же кинематику станции.
start=code.index('Friend Function SteelBoundaryWithinCorrection(');end=code.index("' Меняет одну координату поля",start)
old=code[start:end]
local=old.replace('Friend Function SteelBoundaryWithinCorrection','Friend Function LocalSteelBoundaryWithinCorrection',1)
local=replace(local,'ByRef unresolved As ECoupledSteelUncertainty) As Boolean',
    'ByRef unresolved As ECoupledSteelUncertainty, ByVal selectedFamily As Long, ByVal offset As Double, ByRef definitelyExceeded As Boolean) As Boolean')
local=replace(local,'        ReadSteelStrain operators, point, field, strain, limit','''        If family <> selectedFamily Then GoTo NextSteelPoint
        ReadSteelStrain operators, point, field, strain, limit, offset''')
local=replace(local,'        operators.CopySteelStrainProjection point, projection','        operators.CopySteelStrainProjection point, projection, offset')
local=replace(local,'    Next point\n    unresolved = csuNone','''NextSteelPoint:
    Next point
    definitelyExceeded = exceeded(selectedFamily)
    unresolved = csuNone''')
local=local.replace('    SteelBoundaryWithinCorrection = (governingPoint <> 0)','    LocalSteelBoundaryWithinCorrection = (governingPoint <> 0)')
wrapper='''Friend Function SteelBoundaryWithinCorrection(ByVal operators As CCoupledElasticState, ByRef field() As Double, _
        ByRef correction() As Double, ByRef governingPoint As Long, ByRef unresolved As ECoupledSteelUncertainty, _
        Optional ByRef governingSide As Long = 0) As Boolean
    Dim family As Long, side As Long, candidate As Long, chosen As Long, chosenSide As Long
    Dim response As CCoupledReinforcedResponse, localUnresolved As ECoupledSteelUncertainty
    Dim exceeded As Boolean, familyExceeded As Boolean, group As ECoupledSteelUncertainty
    governingPoint = 0: governingSide = 0: unresolved = csuNone
    For family = 0 To 1
        chosen = 0: chosenSide = 0: familyExceeded = False
        For side = -1 To 1
            Set response = SteelStation(side)
            If Not response Is Nothing Then
                exceeded = False: candidate = 0
                If response.LocalSteelBoundaryWithinCorrection(operators, field, correction, candidate, localUnresolved, _
                        family, side * mSteelHalfStep, exceeded) Then
                    If chosen = 0 Or side = 0 Then chosen = candidate: chosenSide = side
                End If
                familyExceeded = familyExceeded Or exceeded
            End If
        Next side
        If chosen <> 0 And Not familyExceeded Then
            group = csuLongitudinal
            If family = 1 Then group = csuTransverse
            unresolved = unresolved Or group
            If governingPoint = 0 Then governingPoint = chosen: governingSide = chosenSide
        End If
    Next family
    SteelBoundaryWithinCorrection = (governingPoint <> 0)
End Function

' Проверяет одну станцию и одну группу; подтверждённое превышение группы
' исключает неопределённость от других точек той же группы.
'''
code=code[:start]+wrapper+local+code[end:]
code=replace(code,'ByRef field() As Double) As Boolean\n    Dim strain As Double, limit As Double, projection() As Double',
    'ByRef field() As Double, Optional ByVal side As Long = 0) As Boolean\n    Dim offset As Double: offset = SteelStationOffset(side)\n    Dim strain As Double, limit As Double, projection() As Double')
begin=code.index('Friend Function ProjectSteelBoundary(');end=code.index('End Function',begin)
block=code[begin:end].replace('field, strain, limit','field, strain, limit, offset').replace('point, projection','point, projection, offset')
code=code[:begin]+block+code[end:]
code=replace(code,'ByRef field() As Double, ByRef strain As Double, ByRef limit As Double)',
    'ByRef field() As Double, ByRef strain As Double, ByRef limit As Double, Optional ByVal offset As Double = 0#)')
code=replace(code,'    operators.PointStrain point, field, tensor\n    strain = 0#','''    operators.PointStrain point, field, tensor
    If offset <> 0# Then
        Dim gradient(1 To COMPONENT_COUNT) As Double
        operators.AxialGradientAtPoint point, field, gradient
        For i = 1 To COMPONENT_COUNT: tensor(i) = tensor(i) + offset * gradient(i): Next i
    End If
    strain = 0#''')
values[name]=code

name='src/Solver/CCoupledElasticState.cls';code=(root/name).read_text('utf-8-sig')
code=replace(code,'Friend Sub CopySteelStrainProjection(ByVal point As Long, ByRef projection() As Double)',
    'Friend Sub CopySteelStrainProjection(ByVal point As Long, ByRef projection() As Double, Optional ByVal offset As Double = 0#)')
begin=code.index('Friend Sub CopySteelStrainProjection(');end=code.index('End Sub',begin)
block=code[begin:end]+'''    If offset <> 0# Then
        projection(mGradientBase + 1) = projection(mGradientBase + 1) + offset * mSteelDirection(3, point)
        projection(mGradientBase + 2) = projection(mGradientBase + 2) + offset * mY(point) * mSteelDirection(3, point)
        projection(mGradientBase + 3) = projection(mGradientBase + 3) + offset * mX(point) * mSteelDirection(3, point)
        projection(mGradientBase + 4) = projection(mGradientBase + 4) + offset * _
            (-mY(point) * mSteelDirection(5, point) + mX(point) * mSteelDirection(6, point))
    End If
'''
code=code[:begin]+block+code[end:];values[name]=code

name='src/Solver/CCoupledNewtonSolver.cls';code=(root/name).read_text('utf-8-sig')
anchor='    mUseSpatialHistory = Not (plusSide Is Nothing)'
code=replace(code,anchor,anchor+'''
    If TypeOf response Is CCoupledReinforcedResponse Then
        Dim steelCenter As CCoupledReinforcedResponse, steelPositive As CCoupledReinforcedResponse, steelNegative As CCoupledReinforcedResponse
        Set steelCenter = response
        If mUseSpatialHistory Then
            If Not TypeOf plusSide Is CCoupledReinforcedResponse Or Not TypeOf minusSide Is CCoupledReinforcedResponse Then _
                Err.Raise 5, , "Для армированного отклика нужны одинаковые типы трёх станций."
            Set steelPositive = plusSide: Set steelNegative = minusSide
        End If
        steelCenter.AttachSteelStations steelPositive, steelNegative, SPATIAL_HALF_STEP
    End If''')
begin=code.index('Friend Function VerifyDirectSteelTrial(');end=code.index('End Function',begin)
block=code[begin:end]
block=replace(block,'correction() As Double, point As Long','correction() As Double, point As Long, side As Long')
block=replace(block,'correction, point, unresolved)','correction, point, unresolved, side)')
block=replace(block,'refinedCorrection, point, refinedUncertainty)','refinedCorrection, point, refinedUncertainty, side)')
block=replace(block,'ProjectSteelBoundary(mSection, point, boundary)','ProjectSteelBoundary(mSection, point, boundary, side)')
code=code[:begin]+block+code[end:];values[name]=code

name='src/Solver/CCoupledDeformationContinuation.cls';code=(root/name).read_text('utf-8-sig')
code=replace(code,'    CriticalPoint As Long\nEnd Type','    CriticalPoint As Long\n    CriticalSide As Long\nEnd Type')
anchor='    eventState.CriticalPoint = mReinforcement.TrialSteelCriticalPoint'
assert code.count(anchor)==3
code=code.replace(anchor,anchor+'\n    eventState.CriticalSide = mReinforcement.TrialSteelCriticalSide')
code=replace(code,'        "]. Точка арматуры: " & CStr(eventState.CriticalPoint) & "; отношения деформаций к пределам: " & _',
    '        "]. Точка арматуры: " & CStr(eventState.CriticalPoint) & "; станция: " & CStr(eventState.CriticalSide) & "; отношения деформаций к пределам: " & _')
values[name]=code

# Поля хомута должны соответствовать станции максимального использования.
name='src/Solver/CCoupledSectionAnalysis.cls';code=(root/name).read_text('utf-8-sig')
anchor='    If point = 0 Then point = reinforcement.AcceptedTransverseCriticalPoint'
code=replace(code,anchor,'''    Dim stationPoint As Long, side As Long, utilization As Double
    utilization = reinforcement.ReadSteelLimit(True, csuTransverse, stationPoint, side)
    If point = 0 Then point = stationPoint''')
anchor='    mOperators.PointStrain point, field, tensor\n    mOperators.SteelProjection point, direction'
code=replace(code,anchor,'''    mOperators.PointStrain point, field, tensor
    mOperators.SteelProjection point, direction''')
# Выбранная предельная точка может отличаться от текущего максимума семейства.
anchor='    strainLimit = steel.SignedStrainLimit(strain)'
assert code.count(anchor)==1
code=replace(code,anchor,'''    Dim gradient(1 To 6) As Double, candidateStrain As Double, station As Long, ratio As Double
    mOperators.AxialGradientAtPoint point, field, gradient
    utilization = steel.StrainUtilization(strain): side = 0
    For station = -1 To 1 Step 2
        candidateStrain = 0#
        For i = 1 To 6
            candidateStrain = candidateStrain + (tensor(i) + reinforcement.SteelStationOffset(station) * gradient(i)) * direction(i)
        Next i
        ratio = steel.StrainUtilization(candidateStrain)
        If ratio > utilization Then utilization = ratio: side = station
    Next station
    If side <> 0 Then
        strain = 0#
        For i = 1 To 6
            strain = strain + (tensor(i) + reinforcement.SteelStationOffset(side) * gradient(i)) * direction(i)
        Next i
    End If
    strainLimit = steel.SignedStrainLimit(strain)''')
code=replace(code,'    state.SetCoupledTransverseControl SteelControlDescription(point), _\n        steel.StrainUtilization(strain), reinforcement.AcceptedSteelStress(point), strain, strainLimit, criterion',
'''    state.SetCoupledTransverseControl SteelControlDescription(point) & ", станция = " & CStr(side), _
        steel.StrainUtilization(strain), reinforcement.AcceptedStationSteelStress(point, side), strain, strainLimit, criterion''')
values[name]=code

for name,code in values.items():
    assert max(map(len,code.splitlines()))<1024
    data=code.replace('\n','\r\n').encode('utf-8-sig');(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='SpatialSteelLimits',actual_import_this_pass=[Path(n).stem for n in values],actual_component_count=256,
    all_three_station_steel_limits=True,material_laws_and_force_tolerances_unchanged=True,physical_capacity_certified=False,no_release_promotion=True)
for name,data in {'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
                  'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()}.items():
    (root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'SPATIAL STEEL WORKING CANDIDATE STAGED;NATIVE TARGETED TEST ENTRY STILL TO ADD')
