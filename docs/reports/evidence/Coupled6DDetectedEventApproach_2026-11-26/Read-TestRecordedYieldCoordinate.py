"""ДЛЯ ТЕСТОВ. Читает только завершённый двоичный протокол материальных координат без изменения численных данных."""
from pathlib import Path
from runpy import run_path

def TestRead(file):
    """Проверяет размеры четырёх исходов и всех тензоров трёх станций; незаписанный хвост не считается результатом."""
    here=Path(__file__).resolve().parent
    R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];r=R(file.read_bytes())
    assert r.raw[:16]==b'NDMPOINTYIELD001' and r.VectorTest().tolist()==[4.]
    result=[]
    for i in range(1,5):
        info=r.VectorTest();assert len(info)==15 and info[0]==i
        item={'info':info}
        if info[5]:
            item.update(field=r.VectorTest(),values=r.VectorTest(),force=r.VectorTest())
            item['stress']=[(r.VectorTest(),r.VectorTest()) for _ in range(129)]
            assert len(item['field'])==60 and len(item['values'])==61 and len(item['force'])==6
        result.append(item)
    assert r.position==len(r.raw)
    return result
