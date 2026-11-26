"""ДЛЯ ТЕСТОВ. Читает односторонние напряжения у поверхности положительной станции в фактическом отказе929."""
from pathlib import Path
from hashlib import sha256
import json
here=Path(__file__).resolve().parent;file=here/'Stage-TestRecordedTailStressScan.py'
proofpath=Path.cwd()/'docs/reports/evidence/Coupled6DRecordedNewtonTail_2026-11-26/FineLineSurfaceOriginProof.json'
proof=json.loads(proofpath.read_bytes());lo,hi=proof['plus_station_crossing_fraction_bracket'];center=(lo+hi)/2
fractions=[None,0.]+[2.**(-k) for k in range(17)]+[2.**(-20),2.**(-24),-2.**(-16),-2.**(-20),-2.**(-24)]
fractions += [center+s*10.**(-k) for k in range(6,13) for s in (-1,1)]
assert len(fractions)==38
code=file.read_text('utf-8-sig')
for old,new in [('Cold923RecordedNewtonCold96_20261126','Cold929RecordedTailFineLine_20261126'),
 ('Cold926RecordedTailStressScan_20261126','Cold930FineLinePlusSurface_20261126'),
 ('NewtonRejected.00054.bin','NewtonRejected.00001.bin'),('NewtonRejected.00055.bin','NewtonRejected.00002.bin')]:
    assert old in code,old;code=code.replace(old,new)
start=code.index('fractions=[');finish=code.index('\npredictions=',start)
code=code[:start]+'fractions='+repr(fractions)+code[finish:]
scope={'__file__':str(file),'__name__':'fine_line_plus_surface_stage_test'}
exec(compile(code,str(file),'exec'),scope)
root=scope['root'];prep=json.loads((root/'Preparation.json').read_bytes())
for name,data in ((Path(__file__).name,Path(__file__).read_bytes()),('SourceFineLineSurfaceOriginProof.json',proofpath.read_bytes())):
    (root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope'].pop('exact_history25_and_recorded54_current')
prep['directed_scope'].update(case='Actual929PositiveStationSurface',exact_history25_and_actual929_record1_current=True,
    previous54_base_guard_not_claimed_for_new_iterate=True,positive_station_crossing_fraction=center,endpoint_fractions=fractions)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('38 ENDPOINTS AROUND ACTUAL929 POINT5 PLUS CROSSING',center)
