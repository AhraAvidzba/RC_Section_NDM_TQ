"""ДЛЯ ТЕСТОВ. Добавляет короткий читающий вход к существующей уточнённой касательной, не заменяя большой тестовый модуль."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold858BalancedInitialStressReplay_20261125';root=base/'Cold864Cold19BranchTangents_20261125';verified=base/'Cold859SpatialDeclineStressReplay_20261125'
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'BranchTangentCompilationAncestor';archive.mkdir()
for n in ('CompilePoint.log','CompilePointClosure.json'):(root/n).rename(archive/n)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
assert 'Public Function TestReadColdBranchTangents(' not in code
code+='''
' FOR TESTS: baseline entry point only; existing methods retained.
Public Function TestReadColdBranchTangents() As String
    ' -------------------- ДЛЯ ТЕСТОВ --------------------
    ' Читает уточнённые касательные около принятого поля; равновесие не решается, материал и история не принимаются.
    Dim accepted() As Double, previous() As Double, field() As Double
    mReverseSolverTest.CopyAccepted accepted: previous = mReverseFieldsTest(mReverseFieldsTest.Count - 1)
    Dim revision As Long: revision = mReverseSolverTest.AcceptedRevision
    Dim fraction As Double, reference As Double
    mReverseTraceTest.ReadPoint mReverseTraceTest.Count, fraction, reference
    Dim levels As Variant: levels = Array(0#, 0.0009765625, 0.00390625, 0.015625)
    Dim handle As Integer: handle = FreeFile
    Open mPolishRootTest & "\\Native\\ColdBranchTangents.bin" For Binary Access Write As #handle
    Dim magic As String * 16: magic = "NDMCOLDTANGENT01"
    Put #handle, , magic
    TestWriteReverseBinaryVector handle, Array(reference, revision, mReverseModulusTest, 60#, 56#)
    Dim level As Variant, i As Long, degree As Long
    For Each level In levels
        field = accepted
        For i = 1 To UBound(field): field(i) = accepted(i) + CDbl(level) * (accepted(i) - previous(i)): Next i
        mReverseSolverTest.DiscardStep
        Dim projection() As Double, zero() As Double, residual() As Double, matrix() As Double, scales As Variant, ready As Boolean
        ReDim projection(1 To 60): ReDim zero(1 To 61)
        ready = mReverseSolverTest.BeginNewtonLineTest(mReverseBaseTest, mReverseDirectionTest, projection, 0#, field, reference)
        If ready Then ready = mReverseSolverTest.TryReadBranchLinearization(zero, residual, matrix)
        TestWriteReverseBinaryVector handle, Array(CDbl(level), Abs(CLng(ready)), revision, mReverseSolverTest.AcceptedRevision)
        TestWriteReverseBinaryVector handle, field
        If ready Then
            scales = mReverseSolverTest.ReadNewtonLineScalesTest
            TestWriteReverseBinaryVector handle, scales: TestWriteReverseBinaryVector handle, residual
            Dim column() As Double: ReDim column(1 To 61)
            For degree = 1 To 61
                For i = 1 To 61: column(i) = matrix(degree, i): Next i
                TestWriteReverseBinaryVector handle, column
            Next degree
        End If
        mReverseSolverTest.EndNewtonLineTest
        If mReverseSolverTest.AcceptedRevision <> revision Then Err.Raise 5, , "Чтение касательной приняло материальную историю."
    Next level
    Close #handle
    Dim checked() As Double: mReverseSolverTest.CopyAccepted checked
    For i = 1 To UBound(accepted)
        If checked(i) <> accepted(i) Then Err.Raise 5, , "Чтение касательной изменило принятое поле."
    Next i
    TestReadColdBranchTangents = "TOTAL_COLD_BRANCH_TANGENTS_TEST: passed=1; failed=0; no equilibrium or capacity search"
End Function
'''
marker="\n' FOR TESTS: baseline entry point only; existing methods retained.\nPublic Function TestReadColdBranchTangents()"
position=code.index(marker);baseline=code[:position];entry=code[position+1:]
prefix,entry=entry.split('\n',1)
anchor='Private Const HARD_PREPARED_HALF_TEST As Double = 0.001\n';assert baseline.count(anchor)==1
code=baseline.replace(anchor,anchor+prefix+'\n\n\n'+entry,1)  # AddFromString вставляет новую процедуру перед существующими, сохраняя их тела.
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
anchor="-ModuleNames 'modTestCoupledElasticState'";assert compiler.count(anchor)==1
values['Compile-TestExternalPoint.ps1']=compiler.replace(anchor,anchor+" -EntryPointNames 'TestReadColdBranchTangents'",1).replace('\n','\r\n').encode('utf-8-sig')
name='tools/build_workbook/Run-Audit02ModeTests.ps1';script=(root/name).read_text('utf-8-sig')
anchor='        else { $result = [string]$excel.Run($macroName) }';assert script.count(anchor)==1
script=script.replace(anchor,anchor+'''
        # ДЛЯ ТЕСТОВ: исходный экспорт закончен; отдельный читающий вход использует те же подготовленные объекты.
        if ($macro -ne 'modTestCoupledElasticState.RunTestReversePolish') {throw 'Unexpected setup macro for tangent read.'}
        $lines.Add([string]$excel.Run("'RC_Section_NDM.xlsm'!modTestCoupledElasticState.TestReadColdBranchTangents"))
        Save-Progress
''',1)
values[name]=script.replace('\n','\r\n').encode('utf-8-sig')
values.update({'Input/Reverse1.bin':(verified/'Input/Reverse1.bin').read_bytes(),
 'SourceColdStressProof.json':(verified/'VerifiedColdStressReplayProof.json').read_bytes(),
 'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
 'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='Cold19ReadOnlyBranchTangents',actual_import_this_pass=['modTestCoupledElasticState'],
 actual_component_count=256,only_short_test_entry_appended_without_full_module_replacement=True,
 existing_private_members_and_all_working_methods_exact858=True,source_states=19,
 refinement_uses_existing_TryReadBranchLinearization=True,levels=[0,1/1024,1/256,1/64],
 no_new_equilibrium_or_material_acceptance=True,physical_capacity_certified=False,no_working_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print(root,'SHORT READONLY ENTRY APPEND;NO LARGE MODULE REPLACEMENT;ALL WORKING METHODS EXACT858')
