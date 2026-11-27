"""ДЛЯ ТЕСТОВ: потоково сверяет точные байты архива и манифестов с указанным коммитом Git."""
from pathlib import Path
from hashlib import sha256
import subprocess
import sys

repo=Path.cwd();commit=sys.argv[1];directory=sys.argv[2]
for name in ('Files.json','Summary.json','RawArchive.zip'):
    relative=directory+'/'+name;digest=sha256()
    with subprocess.Popen(['git','cat-file','blob',commit+':'+relative],cwd=repo,stdout=subprocess.PIPE) as process:
        while block:=process.stdout.read(1024*1024):digest.update(block)
        assert process.wait()==0
    assert digest.digest()==sha256((repo/relative).read_bytes()).digest(),relative
    print('EXACT_GIT_BLOB',commit,name,digest.hexdigest().upper())
