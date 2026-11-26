"""ДЛЯ ТЕСТОВ: читает только законченные копии принятой траектории, включая снимок перед тайм-аутом."""
from argparse import ArgumentParser
import json
from pathlib import Path
import re


def TestRead(path):
    """Игнорирует незаконченный последний блок; строки одного снимка никогда не смешиваются с другим."""
    last=None; current=None; points=[]; details=''
    for line in path.read_text('cp1251').splitlines():
        row=line.split(';')
        if row[0]=='CHECKPOINT':
            assert len(row)==4
            current={'sequence':int(row[1]),'count':int(row[2]),'result_code':int(row[3])}
            points=[]; details=''
        elif row[0]=='DETAILS' and current is not None: details=line[len('DETAILS;'):]
        elif row[0]=='P' and current is not None:
            assert int(row[1])==len(points)+1 and len(row)>4
            points.append({'fraction':float(row[2]),'lambda':float(row[3]),'field':list(map(float,row[4:]))})
        elif row[0]=='CHECKPOINT_END' and current is not None:
            assert int(row[1])==current['sequence'] and len(points)==current['count']
            assert len({len(point['field']) for point in points})==1
            last={**current,'details':details,'points':points}; current=None
    assert last is not None,'Нет законченного контрольного снимка.'
    return last


def TestLocate(root,phase):
    """Ограничивает чтение каталогом частной книги из собственного протокола запуска."""
    log=(root/'Native'/(phase+'.txt')).read_text('utf-8-sig')
    fixtures=re.findall(r'^FIXTURE: (.+)$',log,re.M); assert len(fixtures)==1
    fixture=Path(fixtures[0]).resolve(); assert fixture.is_relative_to((root/'Temp').resolve())
    return fixture.parent/'MVT2GaussSpan.csv.Checkpoint.csv'


if __name__=='__main__':
    parser=ArgumentParser(description=__doc__); parser.add_argument('--root',type=Path,required=True)
    parser.add_argument('--phase',choices=('Trace44','Trace74','Trace44Extended','Trace143','Near143'),default='Trace44'); parser.add_argument('--copy',action='store_true')
    parser.add_argument('--out',type=Path); args=parser.parse_args(); root=args.root.resolve()
    path=TestLocate(root,args.phase); result=TestRead(path)
    if args.copy:
        target=root/'Native'/(args.phase+'.Checkpoint.csv')
        with target.open('xb') as stream: stream.write(path.read_bytes())
    if args.out: args.out.write_bytes((json.dumps(result,indent=2,sort_keys=True)+'\n').encode())
    print(json.dumps({'sequence':result['sequence'],'states':result['count'],'details':result['details'],
                      'maximum_lambda_on_full_path':max(point['lambda'] for point in result['points'] if point['fraction']==1.)},ensure_ascii=False))
