"""ДЛЯ ТЕСТОВ. Сохраняет точные поля бокового стального предела для независимой сборки и холодного поиска эталонного корня."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import ast,json
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold881SpatialSteelLimits_20261126';root=base/'Cold885SpatialSteelRawControl_20261126'
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'SpatialRawCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
tree=ast.parse((here/'Stage-TestSpatialSteelBoundary.py').read_text('utf-8-sig'))
entry,=[ast.literal_eval(n.value) for n in tree.body if isinstance(n,ast.Assign) and any(isinstance(t,ast.Name) and t.id=='entry' for t in n.targets)]
def change(old,new):
    """Не допускает применения к другому тестовому входу."""
    global entry
    assert entry.count(old)==1,(old[:80],entry.count(old));entry=entry.replace(old,new,1)
change('Public Function RunTestSpatialSteelBoundary() As String','Public Function RunTestSpatialSteelRawControl(ByVal root As String) As String')
entry=entry.replace('RunTestSpatialSteelBoundary =','RunTestSpatialSteelRawControl =')
change('линейный бетон является аналитическим контролем численного алгоритма.','заданный кубический отклик матрицы служит численным контролем и не описывает реальный бетон.')
anchor='    Dim low As Double, high As Double, onOffset As Boolean, history As Collection, meta As CResultMeta'
change(anchor,anchor+'''
    TestWriteHardPreparedOperators operators, root & "\\Native\\SpatialSteelRaw.Operators.csv"
    Dim handle As Integer: handle = FreeFile
    Open root & "\\Native\\SpatialSteelRaw.bin" For Binary Access Write As #handle
    Dim magic As String * 16: magic = "NDMSTEELSIDE0001"
    Put #handle, , magic
    TestWriteReverseBinaryVector handle, Array(operators.DegreeCount, operators.ConcretePointCount, operators.PointCount, _
        30000#, 15000#, 1E+10, 200000#, 1#, 1#, 199000#, 199000#, 0.0001, 0.00013, 0.001, 0.00000001, 0.00000001)
''')
anchor='        central = material(0).LocalSteelLimit(False, csuNone, localPoint)'
change(anchor,anchor+'''
        TestWriteReverseBinaryVector handle, Array(CDbl(sign), 0#, 1#, ratio, central, station)
        TestWriteReverseBinaryVector handle, field: TestWriteReverseBinaryVector handle, actual''')
anchor='            TestRecord ratio <= 1# And ratio > 0.99 And station <> 0 And central < ratio, "admissible bracket checks all stations " & sign, passed, failed, report'
change(anchor,anchor+'''
            solver.CopyAccepted field
            TestWriteReverseBinaryVector handle, Array(CDbl(sign), 1#, low, ratio, central, station)
            TestWriteReverseBinaryVector handle, field: TestWriteReverseBinaryVector handle, actual
            If Not solver.TryStep(0#, 0#, 0#, 0#, CDbl(sign) * high * 30000000#, 0#, 0.00000001, 0.00000001, 60) Then _
                Err.Raise 5, , "Не повторяется превышенное плечо бокового предела."
            solver.CopyTrial field, actual
            ratio = material(0).ReadSteelLimit(False, csuNone, point, station)
            central = material(0).LocalSteelLimit(False, csuNone, localPoint)
            TestWriteReverseBinaryVector handle, Array(CDbl(sign), 2#, high, ratio, central, station)
            TestWriteReverseBinaryVector handle, field: TestWriteReverseBinaryVector handle, actual
            TestRecord ratio > 1# And station <> 0, "exceeded bracket independently repeatable " & sign, passed, failed, report
            solver.DiscardStep''')
change('    Next sign\n    RunTestSpatialSteelRawControl =','    Next sign\n    Close #handle\n    RunTestSpatialSteelRawControl =')
change('TOTAL_SPATIAL_STEEL_BOUNDARY_TEST','TOTAL_SPATIAL_STEEL_RAW_CONTROL_TEST')
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
prefix="' FOR TESTS: baseline entry point only; existing methods retained.\n"
anchor='Private Const HARD_PREPARED_HALF_TEST As Double = 0.001\n';assert code.count(anchor)==1
code=code.replace(anchor,anchor+prefix+'\n\n'+entry,1);values[name]=code
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
values['Compile-TestExternalPoint.ps1']=compiler.replace("-ModuleNames 'modTestCoupledElasticState'","-ModuleNames 'modTestCoupledElasticState' -EntryPointNames 'RunTestSpatialSteelRawControl'",1)
runner=(parent/'Run-TestSpatialSteelLimits.ps1').read_text('utf-8-sig').replace('RunTestSpatialSteelLimits','RunTestSpatialSteelRawControl').replace('SpatialSteel.txt','SpatialSteelRaw.txt')
runner=runner.replace("-Macro 'modTestCoupledElasticState.RunTestSpatialSteelRawControl'","-Macro 'modTestCoupledElasticState.RunTestSpatialSteelRawControl' -MacroArgument1 $Root")
values['Run-TestSpatialSteelRawControl.ps1']=runner
for name,code in values.items():
 data=code.replace('\n','\r\n').encode('utf-8-sig');(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
for name,data in {'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()}.items():
 (root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='SpatialSteelRawControl',actual_import_this_pass=['modTestCoupledElasticState'],actual_component_count=256,
    all_working_sources_exact881=True,only_test_entry_added=True,exact_Double_fields=True,
    matrix_law_cubic_not_real_concrete=True,independent_boundary_root_still_required=True,no_release_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'SIX RAW FIELDS;SAME WORKING STEEL KERNEL881;CORRECT TEST LAW DESCRIPTION')
