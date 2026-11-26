"""ДЛЯ ТЕСТОВ. Восстанавливает точные ZIP-байты частной книги по проверенному родителю и сохранённым сегментам, без Excel."""
from pathlib import Path
from hashlib import sha256
from io import BytesIO
from zipfile import ZipFile
import json, struct

def TestDigest(data):
    """Возвращает контрольную сумму фактических байтов, без нормализации XML или ZIP."""
    return sha256(data).hexdigest().upper()

def TestSegments(data):
    """Разделяет архив на заголовки, сжатые данные и хвосты; смещения и центральный каталог сохраняет буквально."""
    parts=[]
    with ZipFile(BytesIO(data)) as book:
        items=sorted(book.infolist(),key=lambda i:i.header_offset)
        parts.append(('prefix',data[:items[0].header_offset]))
        for number,item in enumerate(items):
            offset=item.header_offset;assert data[offset:offset+4]==b'PK\x03\x04'
            name_size,extra_size=struct.unpack_from('<HH',data,offset+26)
            payload=offset+30+name_size+extra_size;end=payload+item.compress_size
            following=items[number+1].header_offset if number+1<len(items) else book.start_dir
            assert end<=following
            for kind,begin,finish in [('header',offset,payload),('payload',payload,end),('tail',end,following)]:
                parts.append((item.filename+'/'+kind,data[begin:finish]))
        parts.append(('central_directory_and_end',data[book.start_dir:]))
    assert b''.join(value for _,value in parts)==data
    return parts

def TestMakeSegments(parent,target):
    """Переиспользует только побитно совпавшие сегменты; отличия сохраняет полностью, проверяя немедленную обратную сборку."""
    parent_parts=dict(TestSegments(parent));manifest=dict(parent_sha256=TestDigest(parent),target_sha256=TestDigest(target),segments=[])
    blobs={}
    for name,data in TestSegments(target):
        digest=TestDigest(data);record=dict(name=name,sha256=digest,size=len(data))
        if name in parent_parts and parent_parts[name]==data:record['parent_segment']=name
        else:record['blob']=digest;blobs[digest]=data
        manifest['segments'].append(record)
    assert TestRebuildSegments(parent,manifest,blobs)==target
    return manifest,blobs

def TestRebuildSegments(parent,manifest,blobs):
    """Собирает сохранённую книгу и отклоняет неверный родитель, повреждённый сегмент или несовпадение итоговой суммы."""
    assert TestDigest(parent)==manifest['parent_sha256']
    parent_parts=dict(TestSegments(parent));parts=[]
    for item in manifest['segments']:
        data=parent_parts[item['parent_segment']] if 'parent_segment' in item else blobs[item['blob']]
        assert len(data)==item['size'] and TestDigest(data)==item['sha256']
        parts.append(data)
    result=b''.join(parts);assert TestDigest(result)==manifest['target_sha256']
    return result

if __name__=='__main__':
    from argparse import ArgumentParser
    parser=ArgumentParser(description=__doc__)
    parser.add_argument('--parent',type=Path,required=True);parser.add_argument('--manifest',type=Path,required=True)
    parser.add_argument('--blobs',type=Path,required=True);parser.add_argument('--output',type=Path,required=True)
    args=parser.parse_args();manifest=json.loads(args.manifest.read_bytes())
    blobs={i['blob']:(args.blobs/i['blob']).read_bytes() for i in manifest['segments'] if 'blob' in i}
    content=TestRebuildSegments(args.parent.read_bytes(),manifest,blobs)
    with args.output.open('xb') as handle:handle.write(content)
    print('EXACT SAVED BOOK RESTORED',manifest['target_sha256'])
