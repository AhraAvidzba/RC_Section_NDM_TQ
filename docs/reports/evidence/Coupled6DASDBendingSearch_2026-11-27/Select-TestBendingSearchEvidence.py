"""ДЛЯ ТЕСТОВ. Публикует четыре ключевые точные книги и все исходники/протоколы; полный архив всех63 книг сохраняет на D."""
from pathlib import Path
from hashlib import sha256
from io import BytesIO
import json, tarfile, shutil

repo=Path.cwd(); base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
full=base/'BendingSearchExactArchive_20261127'; objects=base/'BendingSearchEvidencePackage_20261127/Objects'
output=repo/'docs/reports/evidence/Coupled6DASDBendingSearch_2026-11-27'
assert full.exists() and not output.exists(); output.mkdir(parents=True)
manifest=json.loads((full/'Files.json').read_bytes()); source_proof=json.loads((full/'ArchiveProof.json').read_bytes())
keep={f'{root}/docs/regression/Diagnostic.xlsm' for root in (
    'Cold1077WorkingASDBendingScales_20261127','Cold1090OwnBendingAdvance_20261127',
    'Cold1091MappedOwnBendingAdvance_20261127','Cold1094ASDAcceptedRecoverySeed_20261127')}
assert keep.issubset(manifest['files'])
selected={name:record for name,record in manifest['files'].items() if not name.lower().endswith(('.xlsm','.xlsx')) or name in keep}
keys={key for record in selected.values() for key in record['segments']}
selected_manifest=dict(manifest,files=selected,objects={k:manifest['objects'][k] for k in sorted(keys)},
    exact_books=4,logical_bytes=sum(r['bytes'] for r in selected.values()))
selected_manifest['scope']='All non-workbook files; four exact primary workbooks. Other original workbook containers remain in the full exact D archive.'
(output/'Files.json').write_text(json.dumps(selected_manifest,ensure_ascii=False,sort_keys=True,indent=2)+'\n',encoding='utf-8')
groups=[]; group=[]; total=0
for key in sorted(keys,key=lambda k:(-manifest['objects'][k],k)):
    size=1024+((manifest['objects'][key]+511)//512)*512
    if group and total+size>40*1024**2: groups.append(group); group=[]; total=0
    group.append(key); total+=size
if group: groups.append(group)
archives=[]
for index,group in enumerate(groups,1):
    path=output/f'RawEvidence{index:02d}.tar.xz'
    with tarfile.open(path,'w:xz',preset=6,format=tarfile.USTAR_FORMAT) as archive:
        for key in group:
            raw=(objects/key).read_bytes(); assert len(raw)==manifest['objects'][key] and sha256(raw).hexdigest().upper()==key
            info=tarfile.TarInfo('Objects/'+key); info.size=len(raw); info.mode=0o644; info.mtime=0
            archive.addfile(info,BytesIO(raw))
    assert path.stat().st_size<48*1024**2
    seen=[]
    with tarfile.open(path,'r:xz') as archive:
        for member in archive:
            key=member.name.removeprefix('Objects/'); raw=archive.extractfile(member).read()
            assert member.isfile() and key in group and sha256(raw).hexdigest().upper()==key and len(raw)==manifest['objects'][key]
            seen.append(key)
    assert sorted(seen)==sorted(group)
    archives.append(dict(file=path.name,bytes=path.stat().st_size,sha256=sha256(path.read_bytes()).hexdigest().upper(),objects=len(group)))
for relative,record in selected.items():
    digest=sha256(); size=0
    for key in record['segments']:
        raw=(objects/key).read_bytes(); digest.update(raw); size+=len(raw)
    assert size==record['bytes'] and digest.hexdigest().upper()==record['sha256'],relative
proof=dict(logical_files=len(selected),logical_bytes=selected_manifest['logical_bytes'],exact_books=4,
    unique_objects=len(keys),all_archive_objects_read_back_and_hashed=True,
    all_files_reconstructed_and_hashed=True,no_zip_recompression=True,archives=archives,
    full_source_archive_on_D=str(full),physical_capacity_certified=False,whole_goal_completed=False)
(output/'ArchiveProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
scope=dict(primary_exact_workbooks=sorted(keep),omitted_workbook_container_files=sorted(set(manifest['files'])-set(selected)),
    full_archive_directory=str(full),full_manifest_sha256=sha256((full/'Files.json').read_bytes()).hexdigest().upper(),
    full_archive_readback_and_all7776_file_hashes_verified=source_proof['all_files_reconstructed_and_hashed'],
    exact_portable_files=len(selected),note='Omitted containers are retained on D. Their source modules, inputs and native evidence are included. Public source workbook is the unchanged0908523 output.')
(output/'Scope.json').write_text(json.dumps(scope,ensure_ascii=False,sort_keys=True,indent=2)+'\n',encoding='utf-8')
for name in ('Restore-TestBendingSearchEvidence.py','Summary.json'): shutil.copyfile(full/name,output/name)
shutil.copyfile(full/'Files.json',output/'FullFiles.json')
shutil.copyfile(full/'ArchiveProof.json',output/'FullArchiveProof.json')
shutil.copyfile(Path(__file__),output/Path(__file__).name)
print('SELECTED EXACT EVIDENCE VERIFIED',json.dumps(proof,sort_keys=True))
