"""ДЛЯ ТЕСТОВ. Сравнивает существующие меры по реальному приращению вместо безусловного предпочтения среднего сдвига."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold851LoadingSecantReparameterization_20261125';root=base/'Cold855BalancedInitialControl_20261125'
assert json.loads((base/'Cold853LoadingSecantCold_20261125/LoadingSecantColdProof.json').read_bytes())['cold832_fields_fractions_lambdas_bit_identical']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'BalancedInitialControlCompilationAncestor';archive.mkdir()
for n in ('CompilePoint.log','CompilePointClosure.json'):(root/n).rename(archive/n)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
name='src/Solver/CCoupledDeformationContinuation.cls';code=(root/name).read_text('utf-8-sig')
old='''    If mSolver.UsesSpatialHistory Then selected = mSolver.SelectTransverseDeformationControl(previousField, field, _
        mDirection, mProjection, current.Coordinate, nominal, mControlName)
    mUsesTransverseControl = selected
    If Not selected Then selected = mSolver.SelectDeformationControl(previousField, field, mProjection, current.Coordinate, nominal, mControlName)'''
assert code.count(old)==1
code=code.replace(old,'    selected = SelectInitialDeformationControl(previousField, field, current.Coordinate, nominal)',1)
code+='''
' Выбирает существующую меру с большим фактическим приращением в безразмерном нормировании ядра.
' Средний сдвиг по направлению Q сохраняет предпочтение при равенстве; нагрузки, знаки и условия последующей приёмки не меняются.
Private Function SelectInitialDeformationControl(ByRef previous() As Double, ByRef current() As Double, _
        ByRef coordinate As Double, ByRef spacing As Double) As Boolean
    Dim selected As Boolean
    selected = mSolver.SelectDeformationControl(previous, current, mProjection, coordinate, spacing, mControlName)
    mUsesTransverseControl = False
    If mSolver.UsesSpatialHistory Then
        Dim transverseProjection() As Double, transverseCoordinate As Double, transverseSpacing As Double, transverseName As String
        If mSolver.SelectTransverseDeformationControl(previous, current, mDirection, transverseProjection, _
                transverseCoordinate, transverseSpacing, transverseName) Then
            If Not selected Or transverseSpacing >= spacing Then
                mProjection = transverseProjection: coordinate = transverseCoordinate: spacing = transverseSpacing
                mControlName = transverseName: selected = True: mUsesTransverseControl = True
            End If
        End If
    End If
    SelectInitialDeformationControl = selected
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает решение читающего выбора меры на заданной паре полей, не решая НДС и не меняя историю материала.
Friend Function TestInitialControlSelection(ByVal solver As CCoupledNewtonSolver, ByRef previous() As Double, _
        ByRef current() As Double, ByRef direction() As Double) As Variant
    Set mSolver = solver: mDirection = direction
    Dim coordinate As Double, spacing As Double, selected As Boolean, projection As Variant
    selected = SelectInitialDeformationControl(previous, current, coordinate, spacing)
    projection = Array(0#)
    If selected Then projection = mProjection
    TestInitialControlSelection = Array(Abs(CLng(selected)), Abs(CLng(mUsesTransverseControl)), coordinate, spacing, projection, solver.AcceptedRevision)
End Function
''';values[name]=code
source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](parent/'Input/Reverse1.bin')
g,L,w,v,coords,axis,b,gradient=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](base/'Cold835GatedColdReplay_20261125/Native/EnvelopeEnergy.Accepted.Operators.csv',60,43)
op=np.array([b+s*.001*gradient for s in (-1,0,1)])
A=op[:,:,:,:56].reshape(-1,56);q=source['direction'][4:];q=q/np.linalg.norm(q)
target=np.zeros((3,43,6));target[:,:,4]=1e-4*q[0];target[:,:,5]=1e-4*q[1]
shear=np.append(np.linalg.lstsq(A,target.reshape(-1),rcond=1e-12)[0],np.zeros(4))
assert np.max(abs(A@shear[:56]-target.reshape(-1)))<1e-12
zero=np.zeros(60);twist=zero.copy();twist[3]=1e-4;axial=zero.copy();axial[0]=1e-4;bending=zero.copy();bending[1]=1e-4
cases=[('ActualInitialHardShear',source['fields'][2],source['fields'][3]),('Zero',zero,zero),
       ('PureShearDirection',zero,shear),('PureTwist',zero,twist),('PureAxial',zero,axial),
       ('PureBendingX',zero,bending),('ShearDominatesTwist',zero,shear+.3*twist)]
vector=run_path(str(here/'Stage-TestColdRawReplay.py'))['TestVector']
values['Input/InitialControl.Fields.bin']=vector([len(cases)])+b''.join(vector(previous)+vector(current) for _,previous,current in cases)
values['Input/InitialControl.Cases.json']=(json.dumps([dict(case=i+1,label=label) for i,(label,_,_) in enumerate(cases)],indent=2)+'\n').encode()
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
old='    TestDiagnoseNewLoadingCoordinate\n';assert code.count(old)==1
code=code.replace(old,'    TestReadInitialControlCases\n',1)
code+='''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет выбор начальной меры на семи кинематических парах; принятые поле и ревизия должны остаться прежними.
Private Sub TestReadInitialControlCases()
    Dim inputHandle As Integer: inputHandle = FreeFile
    Open mPolishRootTest & "\\Input\\InitialControl.Fields.bin" For Binary Access Read As #inputHandle
    Dim header As Variant: header = TestReadPolishVector(inputHandle)
    Dim handle As Integer: handle = FreeFile
    Open mPolishRootTest & "\\Native\\InitialControl.bin" For Binary Access Write As #handle
    Dim magic As String * 16: magic = "NDMINITCTRL00001"
    Put #handle, , magic
    Dim before() As Double: mReverseSolverTest.CopyAccepted before
    Dim revision As Long: revision = mReverseSolverTest.AcceptedRevision
    Dim caseIndex As Long, i As Long, raw As Variant, previous() As Double, current() As Double, result As Variant
    For caseIndex = 1 To CLng(header(1))
        raw = TestReadPolishVector(inputHandle): ReDim previous(1 To UBound(raw))
        For i = 1 To UBound(raw): previous(i) = raw(i): Next i
        raw = TestReadPolishVector(inputHandle): ReDim current(1 To UBound(raw))
        For i = 1 To UBound(raw): current(i) = raw(i): Next i
        Dim subject As New CCoupledDeformationContinuation
        result = subject.TestInitialControlSelection(mReverseSolverTest, previous, current, mReverseDirectionTest)
        TestWriteReverseBinaryVector handle, Array(caseIndex, result(0), result(1), result(2), result(3), revision, mReverseSolverTest.AcceptedRevision)
        TestWriteReverseBinaryVector handle, result(4)
        If mReverseSolverTest.AcceptedRevision <> revision Then Err.Raise 5, , "Выбор начальной меры принял историю."
        Dim after() As Double: mReverseSolverTest.CopyAccepted after
        For i = 1 To UBound(before)
            If before(i) <> after(i) Then Err.Raise 5, , "Выбор начальной меры изменил принятое поле."
        Next i
    Next caseIndex
    Close #handle: Close #inputHandle
End Sub
''';values[name]=code
imports=['CCoupledDeformationContinuation','modTestCoupledElasticState']
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig');old="-ModuleNames 'modTestCoupledElasticState'";assert compiler.count(old)==1
values['Compile-TestExternalPoint.ps1']=compiler.replace(old,'-ModuleNames @('+','.join(repr(n) for n in imports)+')',1)
values.update({'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for name,data in values.items():
    if isinstance(data,str):
        if name.endswith(('.cls','.bas')):assert max(map(len,data.splitlines()))<1024
        data=data.replace('\n','\r\n').encode('utf-8-sig')
    (root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='ResponseBalancedInitialControl',actual_import_this_pass=imports,actual_component_count=256,
    old_SelectDeformation_and_SelectTransverse_methods_reused=True,only_initial_numerical_selector_changed=True,
    seven_readonly_kinematic_cases=True,all_loads_materials_force_and_physical_criteria_unchanged=True,
    default_global96_and_primary30_unchanged=True,cold_workflow_not_yet_run=True,physical_capacity_certified=False,no_working_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print(root,'SEVEN READONLY CASES;EXISTING SELECTORS;96/30ANDMATERIALSUNCHANGED')
