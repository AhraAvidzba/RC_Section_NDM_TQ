"""ДЛЯ ТЕСТОВ. Читает завершённые результаты исходной61-компонентной системы без преобразования нативных Double."""
from pathlib import Path
from runpy import run_path

def TestRead(file):
    """Проверяет обе записи, размеры полей и все129 тензоров каждой пригодной точки."""
    here=Path(__file__).resolve().parent;R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader']
    r=R(file.read_bytes());assert r.raw[:16]==b'NDMORIGCTRL00001' and r.VectorTest().tolist()==[2.]
    cases=[]
    for index in (1,2):
        info=r.VectorTest();assert len(info)==11 and info[0]==index;case={'info':info}
        if info[2]:
            case.update(field=r.VectorTest(),values=r.VectorTest(),forces=r.VectorTest())
            assert len(case['field'])==60 and len(case['values'])==61 and len(case['forces'])==6
            case['stress']=[(r.VectorTest(),r.VectorTest()) for _ in range(129)]
        cases.append(case)
    assert r.position==len(r.raw)
    return cases
