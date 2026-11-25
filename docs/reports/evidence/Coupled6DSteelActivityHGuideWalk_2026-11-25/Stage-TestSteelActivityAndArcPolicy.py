"""ДЛЯ ТЕСТОВ. Добавляет читающий признак пластической работы стали и ограничение первого шага после новой границы нагружения."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,numpy as np

here=Path(__file__).resolve().parent;repo=Path.cwd();base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold844AutomaticHGuideArc_20261125';gated=base/'Cold831GatedLoadingArc_20261125';root=base/'Cold845SteelActivityArcPolicy_20261125'
assert json.loads((parent/'AutomaticHGuideArcProof.json').read_bytes())['allowance6_and96_return_bit_identical_event']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'SteelActivityPolicyCompilationAncestor';archive.mkdir()
for n in ('CompilePoint.log','CompilePointClosure.json'):(root/n).rename(archive/n)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
def TestReplace(code,old,new):
    """Заменяет только один заранее ожидаемый фрагмент исходника."""
    assert code.count(old)==1,old[:160]
    return code.replace(old,new,1)
name='src/Solver/CCoupledNewtonSolver.cls';code=(root/name).read_text('utf-8-sig')
code+='''
' Читает наличие нового пластического приращения любого продольного или поперечного стержня от принятой истории.
' Проверяет все действующие станции; геометрия, поля, буферы материала и принятая ревизия не меняются.
Friend Function HasSteelPlasticIncrement(ByRef field() As Double) As Boolean
    If mSection Is Nothing Then Err.Raise 5, "CCoupledNewtonSolver", "Для чтения работы стали подготовьте сечение и материалы."
    If mPointCount <= mSection.ConcretePointCount Then Exit Function
    If LBound(field) <> 1 Or UBound(field) <> mDegreeCount Then Err.Raise 5, "CCoupledNewtonSolver", "Поле для чтения работы стали не соответствует сечению."
    Dim firstSide As Long, lastSide As Long, side As Long, point As Long
    If mUseSpatialHistory Then firstSide = -1: lastSide = 1
    Dim selected As ICoupledPointResponse, response As CCoupledReinforcedResponse
    Dim strain(1 To 6) As Double, axial As Double, stress As Double, tangent As Double
    For side = firstSide To lastSide
        Select Case side
            Case -1: Set selected = mMinusResponse
            Case 0: Set selected = mResponse
            Case 1: Set selected = mPlusResponse
        End Select
        If TypeOf selected Is CCoupledReinforcedResponse Then
            Set response = selected
            If response.HasSteel Then
                For point = mSection.ConcretePointCount + 1 To mPointCount
                    YieldStrain point, side, field, strain
                    If response.ReadSteelBranch(point, strain, axial, stress, tangent) <> 0 Then HasSteelPlasticIncrement = True: Exit Function
                Next point
            End If
        End If
    Next side
End Function
''';values[name]=code
name='src/Solver/CCoupledArcContinuation.cls';code=(root/name).read_text('utf-8-sig')
code=TestReplace(code,'Private mEnvelopeGuide As Variant','Private mEnvelopeGuide As Variant\nPrivate mNextStepRatioLimit As Double')
code=TestReplace(code,'    mMaximumIterations = maximumIterations: mModulus = solver.ReferenceModulus',
    '    mMaximumIterations = maximumIterations: mModulus = solver.ReferenceModulus: mNextStepRatioLimit = 2#')
code=TestReplace(code,'    Set mApproaches = New Collection: mEnvelopeGuide = Empty',
    '    Set mApproaches = New Collection: mEnvelopeGuide = Empty: mNextStepRatioLimit = 2#')
code=TestReplace(code,'    If baseline <= 0# Then Exit Function\n    If Not mSolver.TryEnvelopeCoordinateAt(field, coordinate) Then Exit Function',
    '    If baseline <= 0# Then Exit Function\n    If mSolver.HasSteelPlasticIncrement(field) Then Exit Function \' Замороженная H бетона не исключает пластическое продвижение стали.\n    If Not mSolver.TryEnvelopeCoordinateAt(field, coordinate) Then Exit Function')
code=TestReplace(code,'    resultLambda = candidateLambda\n    mCertificate =',
    '''    If restored And selectedKind = 2 Then
        If UBound(eventInfo) = 6 Then mNextStepRatioLimit = 1# ' После новой границы пластического нагружения сначала проверяется исходящая касательная без увеличения шага.
    End If
    resultLambda = candidateLambda
    mCertificate =''')
code+='''
' Ограничивает первую попытку после нового начала нагружения; спектральные переходы и обычная дуга сохраняют свой предел.
' Ограничение применяется владельцем только после принятия проверенного кандидата и не меняет критерии допуска следующего решения.
Friend Property Get NextStepRatioLimit() As Double
    NextStepRatioLimit = mNextStepRatioLimit
End Property
''';values[name]=code
name='src/Solver/CCoupledDeformationContinuation.cls'
raw=(gated/name).read_bytes();pinned=repo/'docs/reports/evidence/Coupled6DSpectralPause_2026-11-25/831'/name
assert raw==pinned.read_bytes()
code=raw.decode('utf-8-sig').replace('\r\n','\n')
code=TestReplace(code,'mSignedArcNextRatio = GeomMin(2#, 0.8 * Sqr(0.005 / GeomMax(arc.FieldError, 0.00000001)))',
    'mSignedArcNextRatio = GeomMin(arc.NextStepRatioLimit, 0.8 * Sqr(0.005 / GeomMax(arc.FieldError, 0.00000001)))')
values[name]=code
source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](root/'Input/Reverse1.bin')
ops=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](parent/'Native/EnvelopeEnergy.Accepted.Operators.csv',60,43)
g,L,w,v,coords,axis,b,gradient=ops;op=np.array([b+s*.001*gradient for s in (-1,0,1)])
A=op[:,:,:,:56].reshape(-1,56);fields=[source['fields'][-1],np.zeros(60)];labels=['AcceptedElasticHistory','Zero'];errors=[0.,0.]
specs=[('LongT',[0,0,.003,0,0,0]),('LongC',[0,0,-.003,0,0,0]),('HoopT',[.002,.002,0,0,0,0]),
    ('HoopC',[-.002,-.002,0,0,0,0]),('LongPlusOnly',None),('LongMinusOnly',None),
    ('LongNearT',[0,0,.0019,0,0,0]),('LongNearC',[0,0,-.0019,0,0,0]),
    ('HoopNearT',[.0011,.0011,0,0,0,0]),('HoopBeyondT',[.0013,.0013,0,0,0,0]),
    ('CombinedElastic',[.001,.001,.0016,0,0,0])]
for label,tensor in specs:
    if tensor is None:
        target=np.zeros((3,43,6));target[:,:,2]=np.array([.0005,.0015,.0025] if label=='LongPlusOnly' else [.0025,.0015,.0005])[:,None]
    else:target=np.broadcast_to(tensor,(3,43,6))
    f=np.linalg.lstsq(A,target.reshape(-1),rcond=1e-12)[0];error=float(np.max(abs(A@f-target.reshape(-1))));assert error<1e-12
    fields.append(np.append(f,np.zeros(4)));labels.append(label);errors.append(error)
vector=run_path(str(here/'Stage-TestColdRawReplay.py'))['TestVector']
values['Input/SteelActivity.Fields.bin']=vector([len(fields)])+b''.join(vector(f) for f in fields)
values['Input/SteelActivity.Cases.json']=(json.dumps([dict(case=i+1,label=label,kinematic_fit_error=error) for i,(label,error) in enumerate(zip(labels,errors))],indent=2)+'\n').encode()
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
code=TestReplace(code,'    TestRunAutomaticHGuideArc\n','    TestReadSteelPlasticActivity\n    TestTrainedSteelPlasticActivity\n    TestNoSteelPlasticActivity\n    TestRunAutomaticHGuideArc\n')
start=code.index('Private Sub TestRunAutomaticHGuideArc()');stop=code.index('\nEnd Sub',start)
method=code[start:stop];method=TestReplace(method,'    For caseIndex = 1 To 3','    For caseIndex = 3 To 3')
method=TestReplace(method,'        certificate = arc.Verification: info = certificate(0): guide = arc.EnvelopeGuideVerification',
    '''        certificate = arc.Verification: info = certificate(0): guide = arc.EnvelopeGuideVerification
        Dim limitHandle As Integer: limitHandle = FreeFile
        Open mPolishRootTest & "\\Native\\ArcNextRatio.csv" For Output As #limitHandle
        Print #limitHandle, "LIMIT;" & CStr(caseIndex) & ";" & TestOnsetNumber(arc.NextStepRatioLimit)
        Close #limitHandle''')
code=code[:start]+method+code[stop:]
code=TestReplace(code,'ratio = GeomMin(2#, 0.8 * Sqr(0.005 / GeomMax(arc.FieldError, 0.00000001)))',
    'ratio = GeomMin(arc.NextStepRatioLimit, 0.8 * Sqr(0.005 / GeomMax(arc.FieldError, 0.00000001)))')
code+='''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Читает ветви стали в заданных совместных полях без решения НДС и без изменения принятой истории70.
Private Sub TestReadSteelPlasticActivity()
    Dim inputHandle As Integer: inputHandle = FreeFile
    Open mPolishRootTest & "\\Input\\SteelActivity.Fields.bin" For Binary Access Read As #inputHandle
    Dim header As Variant: header = TestReadPolishVector(inputHandle)
    Dim handle As Integer: handle = FreeFile
    Open mPolishRootTest & "\\Native\\SteelActivity.csv" For Output As #handle
    Dim revision As Long: revision = mReverseSolverTest.AcceptedRevision
    Dim accepted() As Double: mReverseSolverTest.CopyAccepted accepted
    Dim caseIndex As Long, point As Long, side As Long, i As Long, raw As Variant, field() As Double, active As Boolean
    Dim response As CCoupledReinforcedResponse, strain(1 To 6) As Double, slope(1 To 6) As Double, axial As Double, stress As Double, tangent As Double, branch As Long
    Dim before(1 To 3, 26 To 43) As Double
    For side = -1 To 1
        Select Case side
            Case -1: Set response = mReverseSidesTest(2)
            Case 0: Set response = mReverseReinforcedTest
            Case 1: Set response = mReverseSidesTest(4)
        End Select
        For point = 26 To 43: before(side + 2, point) = response.AcceptedSteelStress(point): Next point
    Next side
    For caseIndex = 1 To CLng(header(1))
        raw = TestReadPolishVector(inputHandle): ReDim field(1 To UBound(raw))
        For i = 1 To UBound(raw): field(i) = raw(i): Next i
        active = mReverseSolverTest.HasSteelPlasticIncrement(field)
        Print #handle, "CASE;" & CStr(caseIndex) & ";" & CStr(active) & ";" & CStr(revision) & ";" & CStr(mReverseSolverTest.AcceptedRevision)
        For side = -1 To 1
            Select Case side
                Case -1: Set response = mReverseSidesTest(2)
                Case 0: Set response = mReverseReinforcedTest
                Case 1: Set response = mReverseSidesTest(4)
            End Select
            For point = 26 To 43
                mReverseOperatorsTest.PointStrain point, field, strain: mReverseOperatorsTest.AxialGradientAtPoint point, field, slope
                For i = 1 To 6: strain(i) = strain(i) + side * 0.001 * slope(i): Next i
                branch = response.ReadSteelBranch(point, strain, axial, stress, tangent)
                Print #handle, "BAR;" & CStr(caseIndex) & ";" & CStr(side) & ";" & CStr(point) & ";" & CStr(branch) & _
                    ";" & TestOnsetNumber(axial) & ";" & TestOnsetNumber(stress) & ";" & TestOnsetNumber(tangent)
                If response.AcceptedSteelStress(point) <> before(side + 2, point) Then Err.Raise 5, , "Чтение ветви изменило принятую сталь."
            Next point
        Next side
        Dim checked() As Double: mReverseSolverTest.CopyAccepted checked
        If mReverseSolverTest.AcceptedRevision <> revision Then Err.Raise 5, , "Чтение стали изменило принятую ревизию."
        For i = 1 To UBound(accepted)
            If accepted(i) <> checked(i) Then Err.Raise 5, , "Чтение стали изменило принятое поле."
        Next i
    Next caseIndex
    Close #inputHandle: Close #handle
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Подтверждает настоящий одноосный НДС с пластической сталью и различает её дальнейшее нагружение, упругую разгрузку и обратную текучесть.
Private Sub TestTrainedSteelPlasticActivity()
    Dim section As CSectionModel: Set section = TestHardShearSection()
    Dim operators As New CCoupledElasticState: operators.PrepareModel section, 1000#, 0#, 1, 1, 5, steelModulus:=200000#
    Dim concrete As New CCoupledConcreteResponse: concrete.Initialize operators.PointCount, 1000#, 0#, 16#, 30#, 10#, 0.06, 0.12, 80#, 125#
    Dim steel As New CCoupledSteelModel: steel.Initialize 200000#, 400#, 350#, 1500#, 1000#, 0.025, 0.018
    Dim models As Object: Set models = CreateObject("Scripting.Dictionary"): models.Add "A400", steel
    Dim response As New CCoupledReinforcedResponse: response.Initialize operators, concrete, models
    Dim solver As New CCoupledNewtonSolver: solver.Initialize operators, response, 1000#
    Dim cx As Double, cy As Double, area As Double: section.ConcreteMeshCentroid cx, cy, area
    Dim barArea As Double: barArea = GEOM_PI * 36#
    Dim stress As Double, tangent As Double, plastic As Double, backstress As Double
    steel.EvaluateTrial 0.003, 0#, 0#, stress, tangent, plastic, backstress
    Dim force As Double: force = (area - 2# * barArea) * 1000# * 0.003 + 2# * barArea * stress
    Dim elasticForce As Double: elasticForce = (area * 1000# + 2# * barArea * (200000# - 1000#)) * 0.003
    Dim deformation() As Double, resultants() As Double, pointStress() As Double, rawField As Variant
    If Not operators.Solve(elasticForce, 0#, 0#, 0#, 0#, 0#, deformation, resultants, pointStress, rawField) Then Err.Raise 5, , "Не подготовлено одноосное поле проверки пластической истории стали."
    Dim initial As Variant: initial = rawField
    If Not solver.TryStep(force, 0#, 0#, 0#, 0#, 0#, 0.00000001, 0.00000001, 30, initialField:=initial) Then Err.Raise 5, , "Не подтверждено одноосное НДС с пластической сталью."
    Dim accepted() As Double, actual() As Double: solver.CopyTrial accepted, actual
    solver.AcceptStep
    Dim revision As Long: revision = solver.AcceptedRevision
    Dim handle As Integer: handle = FreeFile
    Open mPolishRootTest & "\\Native\\SteelActivity.Trained.csv" For Output As #handle
    Print #handle, "TRAIN" & TestHardPreparedVector(Array(force, actual(1), area, barArea, stress, plastic, backstress, revision))
    Dim level As Variant, field() As Double, i As Long, active As Boolean
    For Each level In Array(0.0025, 0.0035, -0.003)
        field = accepted
        For i = 1 To UBound(field): field(i) = accepted(i) * CDbl(level) / 0.003: Next i
        active = solver.HasSteelPlasticIncrement(field)
        Print #handle, "READ" & TestHardPreparedVector(Array(CDbl(level), Abs(CLng(active)), revision, solver.AcceptedRevision))
    Next level
    Close #handle
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет применимость читающего признака на подготовленном бетонном сечении без арматуры.
Private Sub TestNoSteelPlasticActivity()
    Dim section As CSectionModel: Set section = TestBuildSection(TestFixtureBox(180#, 120#))
    Dim operators As New CCoupledElasticState: operators.PrepareModel section, 30000#, 0.2, 1, 1, 5
    Dim concrete As New CCoupledConcreteResponse: concrete.Initialize operators.PointCount, 30000#, 0.2, 16#, 30#, 3#, 0.002, 0.12, 80#, 125#
    Dim response As New CCoupledReinforcedResponse: response.Initialize operators, concrete, Nothing
    Dim solver As New CCoupledNewtonSolver: solver.Initialize operators, response, 30000#
    Dim field() As Double: ReDim field(1 To operators.DegreeCount)
    Dim handle As Integer: handle = FreeFile
    Open mPolishRootTest & "\\Native\\SteelActivity.None.csv" For Output As #handle
    Print #handle, "NONE;" & CStr(operators.PointCount) & ";" & CStr(operators.ConcretePointCount) & ";" & CStr(response.HasSteel) & ";" & CStr(solver.HasSteelPlasticIncrement(field))
    Close #handle
End Sub
''';values[name]=code
imports=['CCoupledNewtonSolver','CCoupledArcContinuation','CCoupledDeformationContinuation','modTestCoupledElasticState']
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
values['Compile-TestExternalPoint.ps1']=TestReplace(compiler,"-ModuleNames 'modTestCoupledElasticState'",'-ModuleNames @('+','.join(repr(n) for n in imports)+')')
values.update({'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for n,data in values.items():
    if isinstance(data,str):
        if n.endswith(('.cls','.bas')):assert max(map(len,data.splitlines()))<1024
        data=data.replace('\n','\r\n').encode('utf-8-sig')
    (root/n).write_bytes(data);prep['inputs'][n]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='ReadOnlySteelActivityAndNewLoadingStepPolicy',actual_import_this_pass=sorted(imports),actual_component_count=256,
    material_laws_unchanged=True,steel_activity_uses_existing_ReadSteelBranch=True,three_stations_and_both_bar_families=True,
    loaded_and_unloaded_plastic_history_verified_in_real_axial_NDS=True,noSteel_case=True,
    retained_controller_exact_pinned831_except_next_ratio_cap=True,first_phi_boundary_next_ratio_limit1=True,
    ordinary_and_spectral_ratio_limit2_unchanged=True,automatic_guide_single_allowance96_unchanged_reference70=True,
    physical_capacity_certified=False,no_working_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print(root,'STEEL KINEMATIC CASES',len(fields),'TRAINED PLASTIC HISTORY;NO STEEL;ACTUAL H GUIDE')
