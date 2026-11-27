"""ДЛЯ ТЕСТОВ. Восстанавливает точные исходные файлы/книги из проверенного архива без распаковки произвольных путей TAR."""
from pathlib import Path, PurePosixPath
from hashlib import sha256
from argparse import ArgumentParser
import json, re, tarfile

p = ArgumentParser(description=__doc__)
p.add_argument('--evidence', type=Path, default=Path(__file__).resolve().parent)
p.add_argument('--destination', type=Path, required=True)
a = p.parse_args(); destination = a.destination.resolve()
assert not destination.exists(), 'Для восстановления требуется новый каталог.'
manifest = json.loads((a.evidence / 'Files.json').read_bytes())
proof = json.loads((a.evidence / 'ArchiveProof.json').read_bytes())
assert manifest['format'] == 'NDM_EXACT_SEGMENT_ARCHIVE_1'
objects = {}
for item in proof['archives']:
    path = a.evidence / item['file']
    assert path.name == item['file'] and path.stat().st_size == item['bytes']
    assert sha256(path.read_bytes()).hexdigest().upper() == item['sha256']
    with tarfile.open(path, 'r:xz') as archive:
        for member in archive:
            assert member.isfile() and re.fullmatch(r'Objects/[0-9A-F]{64}', member.name)
            key = member.name.split('/')[1]; raw = archive.extractfile(member).read()
            assert key not in objects and len(raw) == manifest['objects'][key] and sha256(raw).hexdigest().upper() == key
            objects[key] = raw
assert set(objects) == set(manifest['objects'])
destination.mkdir(parents=True)
for relative, record in manifest['files'].items():
    parts = PurePosixPath(relative)
    assert not parts.is_absolute() and '..' not in parts.parts and '\\' not in relative and ':' not in relative
    target = destination.joinpath(*parts.parts).resolve()
    assert target.is_relative_to(destination)
    raw = b''.join(objects[key] for key in record['segments'])
    assert len(raw) == record['bytes'] and sha256(raw).hexdigest().upper() == record['sha256']
    target.parent.mkdir(parents=True, exist_ok=True)
    with target.open('xb') as stream: stream.write(raw)
    assert sha256(target.read_bytes()).hexdigest().upper() == record['sha256']
print('EXACT FILES RESTORED AND HASHED', len(manifest['files']), 'BOOKS', manifest['exact_books'])
