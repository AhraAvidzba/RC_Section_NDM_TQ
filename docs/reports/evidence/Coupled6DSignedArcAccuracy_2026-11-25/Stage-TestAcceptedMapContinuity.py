"""ДЛЯ ТЕСТОВ. Исследует предел материальной невязки при стремлении приращения реального принятого поля к нулю."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import struct
import json

here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent = base / 'Cold761SignedArcHybrid_20261125'
root = base / 'Cold765AcceptedMapContinuity_20261125'
run_path(str(here / 'Stage-TestEnergyNumerics.py'))['TestClone'](root, parent, 1e-6, 1, 1e-6, 0)
archive = root / 'ContinuityCompilationAncestor'
archive.mkdir()
for name in ('CompilePoint.log', 'CompilePointClosure.json'):
    (root / name).rename(archive / name)
prep = json.loads((root / 'Preparation.json').read_bytes())
name = 'tests/modTestCoupledElasticState.bas'
code = (root / name).read_text('utf-8-sig')
assert code.count('    TestRunSignedArcPair\n') == 1
code = code.replace('    TestRunSignedArcPair\n', '    TestRunAcceptedMapContinuity\n', 1)
code += '''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Читает слабые уравнения при всё меньших положительных/отрицательных приращениях, не принимая материал и не решая НДС.
Private Sub TestRunAcceptedMapContinuity()
    Dim accepted() As Double, previous() As Double, movement() As Double, field() As Double
    accepted = mReverseFieldsTest(mReverseFieldsTest.Count): previous = mReverseFieldsTest(mReverseFieldsTest.Count - 1): movement = accepted
    Dim i As Long
    For i = 1 To UBound(movement): movement(i) = accepted(i) - previous(i): Next i
    Dim fraction As Double, factor As Double, load(1 To 6) As Double, external() As Double
    mReverseTraceTest.ReadPoint mReverseTraceTest.Count, fraction, factor
    For i = 1 To 6: load(i) = fraction * mReverseBaseTest(i) + factor * mReverseDirectionTest(i): Next i
    mReverseOperatorsTest.LoadEquations load(1), load(2), load(3), load(4), load(5), load(6), external
    Dim inputHandle As Integer: inputHandle = FreeFile
    Open mPolishRootTest & "\\Input\\MapContinuityOptions.bin" For Binary Access Read As #inputHandle
    Dim settings As Variant: settings = TestReadPolishVector(inputHandle): Close #inputHandle
    Dim revision As Long: revision = mReverseSolverTest.AcceptedRevision
    Dim handle As Integer: handle = FreeFile
    Open mPolishRootTest & "\\Native\\MapContinuity.Result.bin" For Binary Access Write As #handle
    Dim magic As String * 16: magic = "NDMMAPLIMIT00001"
    Put #handle, , magic
    TestWriteReverseBinaryVector handle, Array(factor, revision, mReverseModulusTest, mReverseSolverTest.PhysicalDegreeCount, UBound(accepted))
    TestWriteReverseBinaryVector handle, accepted: TestWriteReverseBinaryVector handle, previous: TestWriteReverseBinaryVector handle, external
    Dim cases As Long: cases = UBound(settings) - LBound(settings) + 1
    Put #handle, , cases
    Dim index As Long, eta As Double, ready As Boolean, weak() As Double, matrix() As Double, changed As Long
    For index = LBound(settings) To UBound(settings)
        eta = settings(index): field = accepted: changed = 0
        For i = 1 To mReverseSolverTest.PhysicalDegreeCount
            field(i) = accepted(i) + eta * movement(i)
            If field(i) <> accepted(i) Then changed = changed + 1
        Next i
        ready = mReverseSolverTest.TestEquilibriumLinearization(field, False, True, weak, matrix)
        If Not ready Then ReDim weak(1 To 1)
        TestWriteReverseBinaryVector handle, Array(eta, Abs(CLng(ready)), changed)
        TestWriteReverseBinaryVector handle, field: TestWriteReverseBinaryVector handle, weak
    Next index
    TestWriteReverseBinaryVector handle, Array(mReverseSolverTest.AcceptedRevision)
    Close #handle
    If mReverseSolverTest.AcceptedRevision <> revision Then Err.Raise 5, , "Читающие пробы непрерывности изменили принятую историю."
End Sub
'''
assert max(map(len, code.splitlines())) < 1024
etas = [0.] + [sign * 10.**(-k) for k in range(9) for sign in (1, -1)] + [0.]
values = {name: code.replace('\n', '\r\n').encode('utf-8-sig'),
    'Input/MapContinuityOptions.bin': struct.pack('<i' + 'd' * len(etas), len(etas), *etas),
    'Compile-TestExternalPoint.ps1': (here / 'Compile-TestExternalPoint.ps1').read_bytes(),
    'ParentCodeUpdate.json': (parent / 'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json': (parent / 'CompilePointClosure.json').read_bytes(), Path(__file__).name: Path(__file__).read_bytes()}
for name, data in values.items():
    (root / name).write_bytes(data)
    prep['inputs'][name] = sha256(data).hexdigest().upper()
prep['directed_scope'].update(case='AcceptedMapContinuityActual35', actual_import_this_pass=['modTestCoupledElasticState'],
    numerical_core_exact761=True, actual35_history_unchanged=True, no_new_equilibrium_search=True,
    exact_zero_and_both_signed_tiny_increments=True, queries_do_not_record_trial_or_accept_history=True,
    fresh_compilation_claimed=True, physical_capacity_certified=False)
(root / 'Preparation.json').write_text(json.dumps(prep, indent=2, sort_keys=True) + '\n', encoding='utf-8')
print('765 READONLY MAP LIMIT AT ACCEPTED35,20 QUERIES; SAME761 KERNEL, NO NEW EQUILIBRIUM SEARCH')
