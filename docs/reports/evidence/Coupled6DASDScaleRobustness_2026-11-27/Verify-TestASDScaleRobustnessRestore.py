"""ДЛЯ ТЕСТОВ. Независимо перечитывает фактически восстановленные файлы и ZIP-контейнеры книг общей M/N-проверки."""
from pathlib import Path, PurePosixPath
from argparse import ArgumentParser
from hashlib import sha256
from zipfile import ZipFile
from datetime import datetime, timezone
import json

parser = ArgumentParser(description=__doc__)
parser.add_argument('--evidence', type=Path, required=True)
parser.add_argument('--restored', type=Path, required=True)
args = parser.parse_args()
destination = args.restored.resolve()
manifest = json.loads((args.evidence / 'Files.json').read_bytes())
archive = json.loads((args.evidence / 'ArchiveProof.json').read_bytes())
assert manifest['format'] == 'NDM_EXACT_SEGMENT_ARCHIVE_1'
assert archive['logical_files'] == len(manifest['files']) and archive['exact_books'] == manifest['exact_books']
books, total = {}, 0
for relative, record in manifest['files'].items():
    parts = PurePosixPath(relative)
    assert not parts.is_absolute() and '..' not in parts.parts and '\\' not in relative and ':' not in relative
    path = destination.joinpath(*parts.parts).resolve()
    assert path.is_relative_to(destination)
    raw = path.read_bytes()
    assert len(raw) == record['bytes'] and sha256(raw).hexdigest().upper() == record['sha256'], relative
    total += len(raw)
    if path.suffix.lower() in ('.xlsm', '.xlsx'):
        with ZipFile(path) as book:
            assert book.testzip() is None and '[Content_Types].xml' in book.namelist()
        books[relative] = record['sha256']
assert total == manifest['logical_bytes'] == archive['logical_bytes'] and len(books) == manifest['exact_books']
for part in archive['archives']:
    path = args.evidence / part['file']
    assert path.name == part['file'] and path.stat().st_size == part['bytes']
    assert sha256(path.read_bytes()).hexdigest().upper() == part['sha256']
proof = dict(utc=datetime.now(timezone.utc).isoformat(), restored_root=str(destination),
             actual_files=len(manifest['files']), actual_bytes=total,
             all_original_file_hashes_match=True, all_archive_hashes_match=True,
             all_restored_book_zip_crc_verified=True, exact_books=books,
             verifier_sha256=sha256(Path(__file__).read_bytes()).hexdigest().upper(),
             no_new_solver_or_Excel_run=True, general_section_capacity_certified=False,
             whole_goal_completed=False)
(args.evidence / 'RestorationProof.json').write_text(json.dumps(proof, ensure_ascii=False, sort_keys=True, indent=2) + '\n', encoding='utf-8')
print('ASD SCALE EVIDENCE ACTUALLY RESTORED AND REVERIFIED', len(manifest['files']), 'FILES', len(books), 'BOOKS')
