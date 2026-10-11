"""Local case register with explicit sources, revisions and atomic updates."""
from contextlib import contextmanager
from copy import deepcopy
from datetime import datetime, timezone
import hashlib
import json
import os
from pathlib import Path
import re
import stat
import uuid

COLLECTIONS = ('documents', 'facts', 'tasks', 'requests', 'contradictions', 'deadlines', 'citations')
STATES = ('не начато', 'в работе', 'проект с пробелами', 'проект после проверки')


def timestamp():
    return datetime.now(timezone.utc).isoformat(timespec='seconds')


def digest(path):
    h = hashlib.sha256()
    with Path(path).open('rb') as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b''):
            h.update(block)
    return h.hexdigest()


def strict_json(path):
    def pairs(items):
        result = {}
        for key, value in items:
            if key in result:
                raise ValueError('Повторный ключ JSON: ' + key)
            result[key] = value
        return result
    return json.loads(Path(path).read_text(encoding='utf-8-sig'), object_pairs_hook=pairs,
                      parse_constant=lambda v: (_ for _ in ()).throw(ValueError(v)))


def folder(project):
    # Windows sandbox may deny GetFinalPathNameByHandle even for readable
    # directories. Check existence separately from non-strict resolution.
    project = Path(project).resolve()
    if not project.is_dir():
        raise ValueError('Проект должен быть существующей папкой')
    target = project / '_ведение_дела'
    reject_reparse(target)
    if not target.resolve().is_relative_to(project):
        raise ValueError('Папка реестра выходит за пределы проекта')
    target.mkdir(exist_ok=True)
    return target


def safe_child(base, name):
    target = base / name
    reject_reparse(base)
    reject_reparse(target)
    if not target.resolve().is_relative_to(base.resolve()):
        raise ValueError('Путь выходит за пределы реестра')
    return target


def reject_reparse(path):
    """Do not write through junctions/symlinks inside the register."""
    try:
        st = path.lstat()
    except FileNotFoundError:
        return
    if stat.S_ISLNK(st.st_mode) or getattr(st, 'st_file_attributes', 0) & getattr(stat, 'FILE_ATTRIBUTE_REPARSE_POINT', 1024):
        raise ValueError('Ссылка или перенаправление внутри реестра не допускается: ' + str(path))


@contextmanager
def lock(base):
    path = safe_child(base, 'write.lock')
    try:
        fd = os.open(path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
    except FileExistsError:
        raise ValueError('Реестр занят; проверьте текущую запись. Блокировка не удаляется автоматически.')
    try:
        os.write(fd, str(os.getpid()).encode('ascii'))
        os.close(fd)
        yield
    finally:
        path.unlink(missing_ok=True)


def atomic_json(path, value):
    temp = path.with_name(path.name + '.' + uuid.uuid4().hex + '.tmp')
    try:
        with temp.open('x', encoding='utf-8', newline='\n') as stream:
            json.dump(value, stream, ensure_ascii=False, indent=2, allow_nan=False)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temp, path)
    finally:
        temp.unlink(missing_ok=True)


def source_path(project, record):
    value = Path(record['path'])
    return value if value.is_absolute() else Path(project).resolve() / value


def require(value, message):
    if not value:
        raise ValueError(message)


def validate(data):
    require(data.get('schema_version') == 1, 'Неподдерживаемая версия реестра')
    require(isinstance(data.get('revision'), int), 'Нет номера редакции')
    require(data.get('case', {}).get('label'), 'Нет обозначения проекта')
    all_ids = set()
    for kind in COLLECTIONS:
        rows = data.get(kind)
        require(isinstance(rows, list), 'Нет списка ' + kind)
        for row in rows:
            rid = row.get('id', '')
            require(re.fullmatch(r'[A-Za-z][A-Za-z0-9_-]{0,63}', rid), 'Недопустимый ID ' + str(rid))
            require(rid not in all_ids, 'Повторный ID ' + rid)
            all_ids.add(rid)
    docs = {r['id']: r for r in data['documents']}
    facts = {r['id']: r for r in data['facts']}

    def refs(items, required=False):
        require(isinstance(items, list), 'sources должен быть списком')
        if required:
            require(items, 'Для подтверждения нужен источник')
        for item in items:
            require(isinstance(item, dict) and item.get('document_id') in docs, 'Неизвестный источник')
            require(isinstance(item.get('locator'), str) and item['locator'].strip(), 'Нет места в источнике')

    def event(value):
        if value:
            require(value.get('status') and value.get('date'), 'У события нужны статус и дата')
            datetime.fromisoformat(value['date'])
            refs(value.get('sources', []), True)

    for doc in docs.values():
        require(doc.get('path') and doc.get('title'), 'Документу нужны path и title')
        require(doc.get('status') in ('рабочая копия', 'подписанный документ', 'скан', 'сообщение пользователя', 'образец', 'не установлен'), 'Не указан статус экземпляра')
        require(doc.get('revisions') in (None, 'current', 'original', 'marked'), 'Неверный режим исправлений')
        if doc.get('revisions'):
            require(doc.get('revision_choice_source'), 'Нужно основание выбора редакции Word')
        if 'volume' in doc or 'leaf_start' in doc or 'leaf_end' in doc:
            require(all(type(doc.get(k)) is int and doc[k] > 0 for k in ('volume', 'leaf_start', 'leaf_end')), 'Том и границы листов должны быть положительными целыми')
            require(doc['leaf_start'] <= doc['leaf_end'], 'Обратный диапазон листов')
            refs(doc.get('inventory_sources', []), True)
        for fid in doc.get('fact_ids', []):
            require(fid in facts, 'Неизвестный факт ' + fid)
    for row in data['facts']:
        require(row.get('text') and row.get('status') in ('сведения документа', 'показания', 'подтверждено', 'версия', 'правовой вывод', 'неизвестно', 'противоречие'), 'Нужен текст и статус факта')
        refs(row.get('sources', []), row['status'] not in ('версия', 'неизвестно'))
    for row in data['tasks']:
        require(row.get('title') and row.get('preparation') in STATES, 'Неверный пункт плана')
        event(row.get('actual'))
        for did in row.get('document_ids', []):
            require(did in docs, 'Неизвестный результат плана')
    for row in data['requests']:
        require(row.get('document_id') in docs and row.get('title'), 'Нет документа запроса')
        event(row.get('sent'))
        items = row.get('items', [])
        require(items and len({i['id'] for i in items}) == len(items), 'Нужны уникальные пункты запроса')
        for item in items:
            require(item.get('question'), 'Нет предмета пункта запроса')
            for name in ('period', 'covered_period'):
                if item.get(name):
                    period = item[name]
                    require(isinstance(period, list) and len(period) == 2, 'Период задаётся двумя датами')
                    first, last = [datetime.fromisoformat(v) for v in period]
                    require(first <= last, 'Обратный период запроса или ответа')
            assessment = item.get('assessment')
            if assessment:
                require(assessment in ('полный', 'частичный', 'не отвечает'), 'Неверная оценка ответа')
                require(item.get('assessment_note'), 'Нужно объяснение оценки ответа')
                refs(item.get('sources', []), True)
        for answer in row.get('responses', []):
            require(answer.get('document_id') in docs and answer.get('received_on'), 'Нет источника или даты ответа')
            datetime.fromisoformat(answer['received_on'])
            for att in answer.get('attachments', []):
                require(att.get('title'), 'Нет названия приложения')
                if att.get('document_id'):
                    require(att['document_id'] in docs, 'Неизвестное приложение')
    for row in data['contradictions']:
        require(row.get('title') and row.get('status') in ('открыто', 'проверяется', 'разрешено'), 'Неверный статус противоречия')
        require(len(row.get('versions', [])) >= 2, 'Нужны как минимум две версии')
        for version in row['versions']:
            require(version.get('text'), 'Нет текста версии')
            refs(version.get('sources', []), True)
        if row['status'] == 'разрешено':
            resolution = row.get('resolution', {})
            require(resolution.get('text') and resolution.get('date'), 'Нет основания разрешения')
            refs(resolution.get('sources', []), True)
    for row in data['deadlines']:
        require(row.get('title'), 'Нет названия срока')
        refs(row.get('sources', []))
        event(row.get('completed'))
        if row.get('superseded_by'):
            require(row['superseded_by'] in {r['id'] for r in data['deadlines']} and row['superseded_by'] != row['id'], 'Неверная замена срока')
            refs(row.get('supersession_sources', []), True)
    deadlines = {r['id']: r for r in data['deadlines']}
    for rid in deadlines:
        seen = set()
        while rid:
            require(rid not in seen, 'Цикл замены сроков')
            seen.add(rid)
            rid = deadlines[rid].get('superseded_by')
    for row in data['citations']:
        require(row.get('document_id') in docs and row.get('target_id') in docs, 'Неизвестный документ ссылки')
        require(row.get('locator') and row.get('literal'), 'Нужны место и буквальный текст ссылки')
        require(all(type(row.get(k)) is int and row[k] > 0 for k in ('volume', 'leaf_start', 'leaf_end')), 'Неверные том или листы ссылки')
        require(row['leaf_start'] <= row['leaf_end'], 'Обратный диапазон ссылки')
    return data


def load(project):
    return validate(strict_json(safe_child(folder(project), 'state.json')))


def init(project, label):
    base = folder(project)
    path = safe_child(base, 'state.json')
    with lock(base):
        require(not path.exists(), 'Реестр уже существует; используйте его, не создавайте заново')
        data = dict(schema_version=1, revision=0, case={'label': label}, settings={'horizon_days': 7, 'report_format': 'docx'}, history=[])
        data.update({c: [] for c in COLLECTIONS})
        atomic_json(path, validate(data))
    return data


def apply(project, patch):
    base = folder(project)
    with lock(base):
        old = load(project)
        require(patch.get('expected_revision') == old['revision'], 'Редакция изменилась; перечитайте реестр')
        require(patch.get('reason'), 'Укажите основание изменения')
        require(set(patch) <= {'expected_revision', 'reason', 'upsert', 'acknowledge_documents'}, 'Неизвестный ключ операции')
        new = deepcopy(old)
        for kind, rows in patch.get('upsert', {}).items():
            require(kind in COLLECTIONS and isinstance(rows, list), 'Неизвестная коллекция')
            require(len({r['id'] for r in rows}) == len(rows), 'Повторный ID в операции')
            mapping = {r['id']: r for r in new[kind]}
            for row in rows:
                require(not {'sha256', 'reviewed_at', '_source_hashes'}.intersection(row), 'Контрольные поля назначает программа')
                row = deepcopy(row)
                prior = mapping.get(row['id'])
                if kind == 'documents' and prior and prior.get('path') == row.get('path'):
                    row.update({k: prior[k] for k in ('sha256', 'reviewed_at') if k in prior})
                mapping[row['id']] = row
            new[kind] = list(mapping.values())
        docs = {r['id']: r for r in new['documents']}
        for did in patch.get('acknowledge_documents', []):
            require(did in docs, 'Неизвестный документ для фиксации проверки')
            docs[did]['sha256'] = digest(source_path(project, docs[did]))
            docs[did]['reviewed_at'] = timestamp()
        # Source revision belongs to each assertion. Rechecking a document alone
        # must not silently validate old assertions, calculations or assessments.
        for kind, rows in patch.get('upsert', {}).items():
            updated = {r['id'] for r in rows}
            for row in new[kind]:
                if row['id'] in updated:
                    row['_source_hashes'] = {did: docs[did].get('sha256') for did in reference_ids(row) if did in docs}
        new['revision'] += 1
        new['history'].append({'revision': new['revision'], 'at': timestamp(), 'reason': patch['reason'], 'operation': patch})
        validate(new)
        backups = safe_child(base, 'history')
        backups.mkdir(exist_ok=True)
        backup = safe_child(backups, f"revision-{old['revision']:06d}.json")
        require(not backup.exists(), 'Историческая редакция уже существует; проверьте журнал')
        atomic_json(backup, old)
        atomic_json(safe_child(base, 'state.json'), new)
    return new


class SourceState(dict):
    """Human states plus current digests for per-assertion revision checks."""
    def __init__(self):
        super().__init__()
        self.hashes = {}


def freshness(project, data):
    result = SourceState()
    for doc in data['documents']:
        try:
            current = digest(source_path(project, doc))
            result.hashes[doc['id']] = current
            result[doc['id']] = 'не проверен' if not doc.get('sha256') else ('совпадает' if doc['sha256'] == current else 'изменён')
        except OSError as exc:
            result[doc['id']] = 'недоступен: ' + str(exc)
    return result


def reference_ids(value):
    result = set()
    if isinstance(value, dict):
        for key, item in value.items():
            if key in ('document_id', 'target_id') and isinstance(item, str):
                result.add(item)
            elif key == 'document_ids':
                result.update(item)
            else:
                result.update(reference_ids(item))
    elif isinstance(value, list):
        for item in value:
            result.update(reference_ids(item))
    return result


def stale_reasons(value, current):
    reasons = [did + ': ' + current.get(did, 'неизвестный источник') for did in sorted(reference_ids(value)) if current.get(did) != 'совпадает']
    if isinstance(value, dict) and hasattr(current, 'hashes'):
        for did in reference_ids(value):
            expected = value.get('_source_hashes', {}).get(did)
            if not expected or expected != current.hashes.get(did):
                message = did + ': версия источника не подтверждена для этой записи'
                if message not in reasons:
                    reasons.append(message)
    return reasons
