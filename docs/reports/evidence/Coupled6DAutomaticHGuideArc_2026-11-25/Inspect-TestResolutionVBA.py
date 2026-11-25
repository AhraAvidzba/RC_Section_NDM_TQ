"""ДЛЯ ТЕСТОВ: сопоставляет фактические девять импортированных компонентов с подготовленными источниками."""

from argparse import ArgumentParser
from difflib import unified_diff
from hashlib import sha256
from io import BytesIO
import json
from pathlib import Path
from zipfile import ZipFile


def TestContainer(root, folder):
    """Восстанавливает ровно сохранённые байты CFB, проверяя каждый фрагмент и полный хеш."""
    metadata = json.loads((folder/'CompiledVBA.json').read_bytes())
    parts = []
    for key in metadata['chunks']:
        raw = (root/'VBAChunks'/(key+'.bin')).read_bytes()
        assert sha256(raw).hexdigest().upper() == key
        parts.append(raw)
    result = b''.join(parts)
    assert len(result) == metadata['bytes'] and sha256(result).hexdigest().upper() == metadata['sha256']
    return result


def TestHelpers(repo):
    """Подключает закреплённый декодер и уже проверенные правила отображения литералов VBE."""
    raw = (repo/'docs/regression/Coupled6D/KrylovFallbackNative20261117.zip').read_bytes()
    assert sha256(raw).hexdigest().upper() == 'C31C5EFABADD5CB81A790C4E5432DBCCDBAE587D3A4639D666881ACCA2BDD633'
    scope = {'__name__': '__pinned__', '__file__': 'Verify-TestKrylovFallback.py'}
    with ZipFile(BytesIO(raw)) as archive:
        exec(compile(archive.read(scope['__file__']), scope['__file__'], 'exec'), scope)
    with scope['TestArchive'](repo, scope['BASE']) as archive:
        observer, reader = scope['TestScript'](archive, 'Stage-Release.py')['parse_vba'](repo)
    return scope, observer['compiler_body'], reader


def TestInspect(repo, root):
    """Показывает только фактические различия; неизвестные нормализации здесь не принимаются."""
    helper, body, reader = TestHelpers(repo)
    for key in json.loads((root/'Cases.json').read_bytes()):
        folder = root/key
        _, actual = reader(helper['TestContainer'](TestContainer(root, folder)))
        compiler = json.loads((folder/'CodeUpdate.json').read_text('utf-8-sig'))
        for name in compiler['Modules']:
            files = [p for parent in ('src', 'tests') for p in (folder/parent).rglob(name+'.*') if p.suffix in ('.bas', '.cls')]
            assert len(files) == 1
            alias = 'modTestMVT2GaussSpan' if name.startswith('modTestMVT2') else name
            expected, _ = helper['TestMVTCompilerBody'](body, files[0].read_text('utf-8-sig'), alias)
            observed = body(actual[name], name)[0]
            if expected != observed:
                print('DIFFERENCE', key, name)
                print('\n'.join(list(unified_diff(expected.splitlines(), observed.splitlines(), lineterm=''))[:70]))
        print('DECODED', key, len(actual), flush=True)


if __name__ == '__main__':
    parser = ArgumentParser(description=__doc__)
    parser.add_argument('--repo', type=Path, default=Path.cwd())
    parser.add_argument('--root', type=Path, required=True)
    args = parser.parse_args()
    TestInspect(args.repo.resolve(), args.root.resolve())
