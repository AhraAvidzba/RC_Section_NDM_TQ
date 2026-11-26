"""ДЛЯ ТЕСТОВ. Сохраняет завершённые объёмные контроли OpenSees и проверяет обратное чтение всех исходных байтов; предел не назначает."""
from pathlib import Path
from hashlib import sha256
from io import BytesIO
from datetime import datetime
import json,tarfile

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
names={993:'Cold993OpenSeesMatchedForce_20261126',994:'Cold994OpenSeesElasticBoundary_20261126',
       995:'Cold995OpenSeesElasticBoundary_20261126',996:'Cold996OpenSeesElasticBoundary_20261126',
       997:'Cold997OpenSeesElasticClamped_20261126',998:'Cold998OpenSeesBufferedCore_20261126',
       999:'Cold999OpenSeesBufferedCore_20261126'}
proof=json.loads((base/'EngineeringOpenSeesBoundaryProof_20261126.json').read_bytes())
assert set(proof['cases'])==set(map(str,names)) and not proof['physical_capacity_certified']
statepath=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';state=json.loads(statepath.read_bytes())
assert not state['active_jobs'] and state['active_exec_session'] is None
raw={}
for n,name in names.items():
    root=base/name
    for file in ('Preparation.json','Run.ndjson','Run.log','Result.json','EngineeringOpenSeesProof.json',
                 'validate_opensees_hardshear.py','Guard-IsolatedOffice.ps1'):
        path=root/file
        assert path.is_file(),path
        raw[f'{n}/{file}']=path.read_bytes()
    assert json.loads(raw[f'{n}/Result.json'])['termination'] is not None
    prep=json.loads(raw[f'{n}/Preparation.json'])
    assert sha256(raw[f'{n}/validate_opensees_hardshear.py']).hexdigest().upper()==prep['source_sha256']
raw['EngineeringOpenSeesBoundaryProof.json']=(base/'EngineeringOpenSeesBoundaryProof_20261126.json').read_bytes()
out=repo/'docs/reports/evidence/Coupled6DEngineeringOpenSees_2026-11-26';assert not out.exists();out.mkdir(parents=True)
(out/'.gitattributes').write_text('* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8')
pins={name:sha256(value).hexdigest().upper() for name,value in raw.items()};canonical={};aliases=0
archivefile=out/'RawEvidence.tar.xz'
with tarfile.open(archivefile,'w:xz',preset=6) as archive:
    for name,data in raw.items():
        info=tarfile.TarInfo(name);digest=pins[name]
        if digest in canonical:
            info.type=tarfile.LNKTYPE;info.linkname=canonical[digest];archive.addfile(info);aliases+=1
        else:
            canonical[digest]=name;info.size=len(data);archive.addfile(info,BytesIO(data))
    data=(json.dumps(pins,sort_keys=True,indent=2)+'\n').encode();info=tarfile.TarInfo('RawFiles.json')
    info.size=len(data);archive.addfile(info,BytesIO(data))
assert archivefile.stat().st_size<100*1024**2
checked={};archived_pins=None
with tarfile.open(archivefile,'r|xz') as archive:
    for info in archive:
        if info.islnk():
            assert info.linkname in checked;digest=checked[info.linkname]
        else:
            data=archive.extractfile(info).read()
            if info.name=='RawFiles.json':archived_pins=json.loads(data);continue
            digest=sha256(data).hexdigest().upper()
        assert digest==pins[info.name];checked[info.name]=digest
assert checked==pins==archived_pins
for name in ('Record-TestEngineeringOpenSees.py','Check-TestEngineeringOpenSees.py'):
    (out/name).write_bytes((here/name).read_bytes())
release={name:sha256((repo/name).read_bytes()).hexdigest().upper() for name in
         ('workbook/output/RC_Section_NDM.xlsm','workbook/output/VBA_All_Code.txt')}
assert release['workbook/output/RC_Section_NDM.xlsm']=='E47ACB7083D92310F9915FD5EB59DA6A7B6C52F352D5EE274A2DCDC6BA567450'
assert release['workbook/output/VBA_All_Code.txt']=='DB8023F9384B0059E1D1C51D59FE0E62F9B8CD94AC4D9B329DD0EE810AF676CC'
summary=dict(proof=proof,raw_files=len(raw),raw_bytes=sum(map(len,raw.values())),archive_bytes=archivefile.stat().st_size,
             exact_duplicate_hardlinks=aliases,all_raw_hashes_read_back=True,all_started_stages_terminal=True,
             release_hashes_unchanged=release,full_acceptance_not_run=True,physical_capacity_certified=False,whole_goal_completed=False)
(out/'Summary.json').write_text(json.dumps(summary,sort_keys=True,indent=2)+'\n',encoding='utf-8')
index={p.relative_to(out).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file()}
(out/'Files.json').write_text(json.dumps(index,sort_keys=True,indent=2)+'\n',encoding='utf-8')
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',pause_requested_explicitly=False,
             pause_request_revoked=True,latest_engineering_opensees_evidence=out.relative_to(repo).as_posix(),
             latest_engineering_opensees_control=proof,
             next_action='Сохранить точные книги и отрицательные силовые контроли 988–991; проверить одинаковый пространственный закон ASD для НДМ и OpenSees. Остановку по сходимости не считать пределом.',
             previous_goal_turn_classification='External engineering control: actual 3D OpenSees runs completed; support-distance sensitivity measured, all six resultants independently checked. Nonlinear stop at lambda0.699 is not certified capacity and ASD spatial law differs from NDM CDP. No production solver change or whole-goal completion.')
data=(json.dumps(state,ensure_ascii=False,sort_keys=True,indent=2)+'\n').encode('utf-8')
for path in (statepath,here/'Continuation666_Resume20261123.json'):path.write_bytes(data)
print('ARCHIVED OPENSEES993-999',len(raw),'RAW;',archivefile.stat().st_size,'BYTES;GOAL ACTIVE')
