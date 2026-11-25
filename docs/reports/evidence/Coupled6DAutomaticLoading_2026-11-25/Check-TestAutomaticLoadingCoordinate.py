"""ДЛЯ ТЕСТОВ. Проверяет автоматический выбор станции и независимо пересобирает оба исходящих равновесия."""
from pathlib import Path
from runpy import run_path
import json,sys
import numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold818AutomaticLoadingCoordinate_20261125';parent=base/'Cold815ActiveLoadingNeighborhood_20261125'
name='src/Solver/CCoupledNewtonSolver.cls';code=(root/name).read_text('utf-8-sig');before=(parent/name).read_text('utf-8-sig')
header="' -------------------- ДЛЯ ТЕСТОВ --------------------\n' Использует существующую материальную координату"
expected=before.split(header)[0].rstrip()
recovered=code.split("' Сохраняет единственную новую границу")[0].rstrip()
for value in (
    'Private mRecentLoadingPoint As Long, mRecentLoadingSide As Long, mRecentLoadingCount As Long\nPrivate mRecentLoadingRevision As Long, mRecentLoadingFrom As Double, mRecentLoadingTo As Double\n',
    '    mRecentLoadingPoint = 0: mRecentLoadingCount = 0: mRecentLoadingRevision = -1\n',
    '    PrepareRecentLoadingBoundary acceptedField\n',
    '        PrepareRecentLoadingBoundary field\n'):
    assert recovered.count(value)==1;recovered=recovered.replace(value,'',1)
assert recovered==expected
driver=(root/'tests/modTestCoupledElasticState.bas').read_text('utf-8-sig').split('Private Sub TestRunActiveLoadingCoordinate()')[1].split('End Sub')[0]
assert 'point = 25' not in driver and 'side = 1' not in driver
rows=run_path(str(here/'Validate-TestHardPrepared.py'))['TestRows'](root/'Native/RecentLoading.Selection.csv')
assert len(rows)==1 and rows[0][0]=='RECENT'
selection=np.array(rows[0][1:],float)
assert selection[0]==1 and selection[1]==25 and selection[2]==1 and selection[5]==1 and selection[6]==45
assert abs(selection[3])<=32e-12 and selection[4]-selection[3]>64e-12
for name in ('Input/Reverse1.bin','Input/SourceColdRawHistory.bin','Native/EnvelopeEnergy.Accepted.csv','Native/EnvelopeEnergy.Accepted.Operators.csv','Native/ActiveLoading.Result.bin'):
    assert (root/name).read_bytes()==(parent/name).read_bytes(),name
checker=here/'Check-TestActiveLoadingCoordinate.py';test=checker.read_text('utf-8-sig')
start="name='src/Solver/CCoupledNewtonSolver.cls';code="
end="Reader=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader']"
a=test.index(start);b=test.index(end,a);test=test[:a]+test[b:]
test=test.replace('all_working_methods_unchanged=True,only_test_coordinate_entry_added=True,',
    'existing_coordinate_force_and_material_methods_unchanged=True,recent_boundary_state_and_working_coordinate_entry_added=True,')
test=test.replace("root/'ActiveLoadingCoordinateProof.json'","root/'AutomaticLoadingCoordinateProof.json'")
test=test.replace('all129station_surface_values_captured=True,','all129station_surface_values_captured=True,automatic_unique_recent_station=True,')
sys.argv=[str(checker),'--root',str(root),'--original-predictor-neighborhood']
exec(compile(test,str(checker),'exec'),{'__file__':str(checker),'__name__':'automatic_loading_coordinate_test'})
proof=json.loads((root/'AutomaticLoadingCoordinateProof.json').read_bytes())
proof['automatic_selection']=selection.tolist()
proof['all_numerical_root_fields_efforts_and_original_stresses_bitexact815']=True
proof['recent_context_negative_tests_not_yet_run']=True
(root/'AutomaticLoadingCoordinateProof.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('AUTOMATIC UNIQUE RECENT STATION; ALL ROOT DOUBLES/EFFORTS EXACT815; NO PHYSICAL LIMIT')
