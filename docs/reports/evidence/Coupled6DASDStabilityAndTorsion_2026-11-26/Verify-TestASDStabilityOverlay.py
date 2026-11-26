"""ДЛЯ ТЕСТОВ. Восстанавливает книги ASD только из закреплённых архивов проекта, без диска D и без Excel."""
from pathlib import Path
from hashlib import sha256
from runpy import run_path
from argparse import ArgumentParser
import json, tarfile


def TestDigest(data):
    """Вычисляет хеш фактических байтов; представление книги не нормализуется."""
    return sha256(data).hexdigest().upper()


def TestRepositoryPath(repo, value):
    """Отклоняет ссылку метаданных, выходящую за корень сохранённого проекта."""
    result = (repo / value).resolve()
    assert result.is_relative_to(repo), value
    return result


def TestReadMember(archive, member):
    """Читает именованный файл без распаковки и изменения файловой системы."""
    with tarfile.open(archive, 'r:xz') as handle:
        return handle.extractfile(member).read()


def TestVerify(repo, folder):
    """Проверяет ссылки на родительские байты и побайтно восстанавливает все одиннадцать фактических книг."""
    pins = json.loads((folder / 'Files.json').read_bytes())
    for name, digest in pins.items():
        assert TestDigest((folder / name).read_bytes()) == digest, name
    summary = json.loads((folder / 'Summary.json').read_bytes())
    payloads = {}
    with tarfile.open(folder / 'RawOverlay.tar.xz', 'r|xz') as handle:
        for info in handle:
            assert info.isfile(), info.name
            data = handle.extractfile(info).read()
            if info.name.startswith('Payloads/'):
                digest = info.name.split('/')[-1]
                assert TestDigest(data) == digest
                payloads[digest] = data
            elif info.name == 'LogicalFiles.json':
                logical = json.loads(data)
            elif info.name == 'ParentArchive.json':
                parent = json.loads(data)
            else:
                raise AssertionError(info.name)
    assert parent == summary['parent_archive']
    parent_path = TestRepositoryPath(repo, parent['archive'])
    assert TestDigest(parent_path.read_bytes()) == parent['archive_sha256']
    raw_index = json.loads(TestReadMember(parent_path, parent['raw_index_member']))
    base_manifest_data = TestReadMember(parent_path, parent['base_book_manifest_member'])
    assert TestDigest(base_manifest_data) == raw_index[parent['base_book_manifest_member']]
    base_manifest = json.loads(base_manifest_data)
    needed = {item['blob'] for item in base_manifest['segments'] if 'blob' in item}
    for name, record in logical.items():
        if 'parent_member' in record:
            assert raw_index[record['parent_member']] == record['sha256'], name
            needed.add(record['sha256'])
        else:
            assert record['payload'] == 'Payloads/' + record['sha256']
            assert TestDigest(payloads[record['sha256']]) == record['sha256']
    parent_bytes = {}
    with tarfile.open(parent_path, 'r|xz') as handle:
        for info in handle:
            if not info.isfile() or info.name not in raw_index:
                continue  # Жёсткие ссылки разрешаются по тому же хешу канонического файла.
            digest = raw_index[info.name]
            if digest in needed and digest not in parent_bytes:
                data = handle.extractfile(info).read()
                assert TestDigest(data) == digest
                parent_bytes[digest] = data
    assert set(parent_bytes) == needed, needed - set(parent_bytes)
    ancestor = parent['base_book_reconstruction']
    ancestor_path = TestRepositoryPath(repo, ancestor['parent_archive'])
    assert TestDigest(ancestor_path.read_bytes()) == ancestor['parent_archive_sha256']
    ancestor_book = TestReadMember(ancestor_path, ancestor['parent_book_member'])
    rebuild = run_path(str(folder / 'Rebuild-TestBookSegments.py'))['TestRebuildSegments']
    base_book = rebuild(ancestor_book, base_manifest, parent_bytes)
    assert TestDigest(base_book) == parent['base_book_sha256'] == ancestor['actual_book_sha256']
    restored = {}
    for case, record in summary['book_reconstruction'].items():
        manifest_record = logical[record['manifest_member']]
        available = payloads if 'payload' in manifest_record else parent_bytes
        manifest = json.loads(available[manifest_record['sha256']])
        blobs = {}
        for item in manifest['segments']:
            if 'blob' in item:
                digest = item['blob']
                blobs[digest] = payloads[digest] if digest in payloads else parent_bytes[digest]
        book = rebuild(base_book, manifest, blobs)
        assert TestDigest(book) == record['actual_saved_book_sha256']
        restored[case] = TestDigest(book)
    assert len(restored) == 11
    return dict(all_archive_pins_verified=True, all_logical_files_resolved_and_hashed=True,
                no_scratch_disk_or_excel_used=True, all11_books_rebuilt_bitexact=True,
                logical_files=len(logical), parent_unique_bytes_verified=len(parent_bytes),
                restored_book_sha256=restored, physical_capacity_certified=False,
                whole_goal_completed=False)


if __name__ == '__main__':
    parser = ArgumentParser(description=__doc__)
    parser.add_argument('--repo', type=Path, default=Path.cwd())
    parser.add_argument('--folder', type=Path, required=True)
    args = parser.parse_args()
    proof = TestVerify(args.repo.resolve(), args.folder.resolve())
    (args.folder / 'ArchiveOnlyRecoveryProof.json').write_text(
        json.dumps(proof, sort_keys=True, indent=2) + '\n', encoding='utf-8')
    print('ARCHIVE-ONLY RECOVERY VERIFIED', proof['logical_files'], 'FILES;11 EXACT BOOKS;NO SCRATCH DEPENDENCY')
