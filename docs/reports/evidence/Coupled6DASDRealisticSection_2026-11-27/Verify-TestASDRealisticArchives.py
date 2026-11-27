"""ДЛЯ ТЕСТОВ: сверяет все прямые и унаследованные payload двух новых архивов с манифестами."""
from pathlib import Path
from hashlib import sha256
from zipfile import ZipFile
from io import BytesIO
import json

root=Path.cwd()/'docs/reports/evidence/Coupled6DASDRealisticSection_2026-11-27'
verified=set();books=set();parents=set()
for part in ('Bridge','Results'):
    manifest_path=root/part/'Files.json';manifest=json.loads(manifest_path.read_bytes())
    archive=root/part/'RawArchive.zip'
    assert sha256(archive.read_bytes()).hexdigest().upper()==manifest['archive_sha256']
    for parent in manifest['parent_manifests']:
        path=Path.cwd()/parent['manifest'];assert sha256(path.read_bytes()).hexdigest().upper()==parent['sha256'];parents.add(str(path))
    for row in manifest['files']:
        path=(manifest_path.parent/row['archive']).resolve();key=(str(path),row['payload'])
        if key in verified:continue
        with ZipFile(path) as saved:raw=saved.read(row['payload'])
        assert len(raw)==row['bytes'] and sha256(raw).hexdigest().upper()==row['sha256'],row['path']
        if raw.startswith(b'PK\x03\x04'):
            with ZipFile(BytesIO(raw)) as package:
                if 'xl/vbaProject.bin' in package.namelist():assert package.testzip() is None;books.add(row['sha256'])
        verified.add(key)
proof=dict(all_direct_and_inherited_payloads_SHA256_verified=True,unique_payloads=len(verified),
    distinct_books_CRC_verified=len(books),parent_manifests_verified=len(parents),no_files_modified=True)
print(json.dumps(proof,sort_keys=True))
