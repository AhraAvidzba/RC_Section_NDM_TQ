"""ДЛЯ ТЕСТОВ. Читает точные Double холодных снимков и H-запросов, различая законченные записи и незаписанный хвост."""
from argparse import ArgumentParser
from hashlib import sha256
import json
from pathlib import Path
import struct
import numpy as np


class TestReader:
    """ДЛЯ ТЕСТОВ. Ограниченный little-endian читатель; выход за конец не считается завершением процесса."""
    def __init__(self, raw):
        """Закрепляет неизменяемые байты одного чтения файла."""
        self.raw, self.position = raw, 16

    def ReadTest(self, format):
        """Не продвигает позицию при недостатке байтов текущего снимка."""
        size = struct.calcsize(format)
        if self.position + size > len(self.raw):
            raise EOFError('Incomplete native record')
        values = struct.unpack_from(format, self.raw, self.position)
        self.position += size
        return values

    def VectorTest(self):
        """Сохраняет все исходные биты численного вектора без преобразования в десятичный текст."""
        count, = self.ReadTest('<i')
        assert 0 < count <= 10000
        return np.asarray(self.ReadTest('<' + 'd' * count), dtype='<f8')


def TestHistory(file, live=False):
    """Возвращает только целые снимки; незаконченный последний снимок разрешён исключительно при живом наблюдении."""
    raw = file.read_bytes()
    assert raw[:16] == b'NDMSTOCKRAW00001'
    reader, frames = TestReader(raw), []
    while reader.position < len(raw):
        begin = reader.position
        try:
            number, count = reader.ReadTest('<ii')
            assert number == len(frames) + 1 and 2 <= count <= 1000
            fractions, factors, fields = [], [], []
            for _ in range(count):
                fraction, factor = reader.ReadTest('<dd')
                field = reader.VectorTest()
                assert len(field) == 60 and np.all(np.isfinite(field))
                assert np.isfinite(fraction) and np.isfinite(factor)
                fractions.append(fraction)
                factors.append(factor)
                fields.append(field)
        except EOFError:
            assert live
            reader.position = begin
            break
        frames.append(dict(number=number, fields=np.asarray(fields), fractions=np.asarray(fractions), lambdas=np.asarray(factors)))
    return frames, dict(bytes=len(raw), complete_bytes=reader.position, incomplete_tail_bytes=len(raw) - reader.position,
        original_binary_Double=True, sha256=sha256(raw).hexdigest().upper())


def TestRequests(file, live=False):
    """Читает точные параметры исходной H-пробы, включая фактические предыдущий вектор и прогноз."""
    raw = file.read_bytes()
    assert raw[:16] == b'NDMCOLDHREQ00001'
    reader, requests = TestReader(raw), []
    while reader.position < len(raw):
        begin = reader.position
        try:
            number, = reader.ReadTest('<i')
            values = [reader.VectorTest() for _ in range(6)]
        except EOFError:
            assert live
            reader.position = begin
            break
        info, offset, direction, previous, seed, accepted = values
        assert number == len(requests) + 1
        assert list(map(len, values)) == [11, 6, 6, 60, 60, 60]
        assert all(np.all(np.isfinite(v)) for v in values)
        assert info[7] >= 1 and info[7] <= 30 and info[8] in (0, 1) and info[10] in (0, 1)
        requests.append(dict(number=number, info=info, offset=offset, direction=direction, previous=previous, seed=seed, accepted=accepted))
    return requests, dict(bytes=len(raw), complete_bytes=reader.position, incomplete_tail_bytes=len(raw) - reader.position,
        original_binary_Double=True, sha256=sha256(raw).hexdigest().upper())


def TestSummary(root, live=False):
    """Показывает наблюдаемый путь; ни последний множитель, ни отсутствие нового снимка не определяют физический предел."""
    frames, history = TestHistory(root / 'Native/Stock.RawHistory.bin', live)
    request_file = root / 'Native/Stock.HRequests.bin'
    if request_file.exists():
        requests, goals = TestRequests(request_file, live)
    else:
        assert live
        requests, goals = [], dict(not_created_yet=True)
    assert frames
    last = frames[-1]
    result = dict(checkpoints=len(frames), states=len(last['fields']), maximum_lambda=float(max(last['lambdas'])),
        last_lambda=float(last['lambdas'][-1]), H_requests=len(requests), history_file=history, H_file=goals,
        current_run_terminal_not_proven_by_this_reader=True, physical_capacity_certified=False)
    if requests:
        request = requests[-1]
        matched = []
        for frame in frames:
            indices = np.flatnonzero(np.all(frame['fields'] == request['accepted'], axis=1) & (frame['lambdas'] == request['info'][2]))
            matched.extend((frame['number'], int(i) + 1) for i in indices)
        result['last_H_request'] = dict(number=request['number'], snapshot_number=int(request['info'][0]),
            accepted_revision=int(request['info'][1]), accepted_lambda=float(request['info'][2]),
            seed_lambda=float(request['info'][3]), target_H=float(request['info'][4]), baseline_H=float(request['info'][9]),
            maximum_iterations=int(request['info'][7]), accepted_field_and_lambda_exact_matches=matched)
    return result


if __name__ == '__main__':
    parser = ArgumentParser(description=__doc__)
    parser.add_argument('--root', type=Path, required=True)
    parser.add_argument('--live', action='store_true')
    args = parser.parse_args()
    print(json.dumps(TestSummary(args.root.resolve(), args.live), indent=2, sort_keys=True))
