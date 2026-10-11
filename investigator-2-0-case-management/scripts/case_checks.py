"""Read-only checks for case sources, requests, citations and explicit deadline rules."""
from calendar import monthrange
from datetime import date, datetime, timedelta
from pathlib import Path
import re
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / 'investigator-2-0' / 'scripts'))
from word_text import read_docx
from case_store import reference_ids, source_path, stale_reasons, require

CITATION_RX = re.compile(r'т(?:ом)?\.?\s*(\d+)\s*[,;]?\s*л\.?\s*д\.?\s*(\d+)(?:\s*[–—-]\s*(\d+))?', re.I)


def read_text(project, doc):
    path = source_path(project, doc)
    if path.suffix.lower() == '.docx':
        return read_docx(path, doc.get('revisions'), include_auxiliary=True)[0]
    if path.suffix.lower() in ('.txt', '.md'):
        return path.read_text(encoding='utf-8-sig')
    if path.suffix.lower() == '.pdf':
        import fitz
        with fitz.open(path) as pdf:
            pages = [p.get_text() for p in pdf]
        empty = [str(i + 1) for i, page in enumerate(pages) if not page.strip()]
        if empty:
            raise ValueError('Нет текстового слоя на страницах ' + ', '.join(empty) + '; нужен локальный OCR и визуальная сверка')
        return '\n'.join(pages)
    raise ValueError('Формат требует отдельного чтения или локальной конвертации: ' + path.suffix)


def norm(text):
    return re.sub(r'\s+', ' ', text).strip().casefold()


def source_text(refs):
    return '; '.join(r['document_id'] + ' — ' + r['locator'] for r in refs) or 'источник не указан'


def impact(project, data, current, fact_id, old_value, new_value):
    facts = {r['id']: r for r in data['facts']}
    require(fact_id in facts, 'Неизвестный ID факта')
    require(old_value.strip() and new_value.strip(), 'Нужны прежнее и предложенное значения')
    lines = [f"Факт {fact_id}: {facts[fact_id]['text']}", f'Предлагаемое изменение: {old_value} → {new_value}. Реестр и документы не изменены.', 'Совпадения текста — кандидаты на проверку контекста. Отсутствие буквального совпадения не исключает смысловую связь.']
    incomplete = []
    hits = []
    for doc in data['documents']:
        linked = fact_id in doc.get('fact_ids', [])
        try:
            text = norm(read_text(project, doc))
            needle = norm(old_value)
            # Do not match 100 inside 1000; allow whitespace differences.
            pattern = r'(?<!\w)' + re.escape(needle) + r'(?!\w)'
            match = re.search(pattern, text)
            if linked or match:
                context = text[max(0, match.start()-80):match.end()+120] if match else 'Буквальное совпадение не найдено; связь зарегистрирована.'
                hits.append(doc['id'])
                lines.append(f"{doc['id']} — {doc['title']} | {doc['status']} | {current[doc['id']]} | {'связь с фактом; ' if linked else ''}{context}")
        except Exception as exc:
            incomplete.append(doc['id'] + ': ' + str(exc))
            if linked:
                hits.append(doc['id'])
                lines.append(doc['id'] + ' — связь зарегистрирована, содержание не проверено')
    for kind in ('tasks', 'requests', 'contradictions', 'deadlines', 'citations'):
        for row in data[kind]:
            if fact_id in row.get('fact_ids', []) or reference_ids(row).intersection(hits):
                lines.append(f"Связанная запись {row['id']} ({kind}): требуется оценить последствия изменения.")
    if not hits:
        lines.append('Кандидаты в прочитанной части зарегистрированных документов не найдены.')
    lines.append('Область: только зарегистрированные документы. Изменения в подписанных документах оформляются через профильный навык и надлежащий порядок, без автоматической перезаписи.')
    if incomplete:
        lines.append('Поиск неполон: ' + ' | '.join(incomplete))
    return lines


def requests_report(data, current):
    lines = []
    for row in data['requests']:
        sent = row.get('sent')
        lines.append(row['id'] + ' — ' + row['title'])
        lines.append(('Направление: ' + sent['date'] + '; ' + source_text(sent['sources'])) if sent else 'Направление не подтверждено. Дата файла не является датой отправки.')
        reasons = stale_reasons(row, current)
        if reasons:
            lines.append('Требует повторной сверки: ' + '; '.join(reasons))
        responses = row.get('responses', [])
        if not responses:
            lines.append('Поступление ответа не зарегистрировано; это не доказывает его отсутствия вне реестра.')
        for response in responses:
            lines.append('Ответ: ' + response['document_id'] + ', получен ' + response['received_on'])
            for attachment in response.get('attachments', []):
                did = attachment.get('document_id')
                lines.append('Приложение «' + attachment['title'] + '»: ' + (did + ', ' + current[did] if did else 'не представлено в реестре'))
        for item in row['items']:
            lines.append('Пункт ' + item['id'] + ': ' + item['question'] + ' — ' + item.get('assessment', 'содержательная полнота не оценена') + '. ' + item.get('assessment_note', ''))
            if item.get('sources'):
                lines.append('Основание оценки: ' + source_text(item['sources']))
            requested = item.get('period')
            covered = item.get('covered_period')
            if requested:
                lines.append('Запрошенный период: ' + ' — '.join(requested))
            if covered:
                lines.append('Период представленных сведений: ' + ' — '.join(covered))
                if requested and (date.fromisoformat(covered[0]) > date.fromisoformat(requested[0]) or date.fromisoformat(covered[1]) < date.fromisoformat(requested[1])):
                    lines.append('Период ответа не покрывает весь запрошенный период.')
    return lines or ['Запросы пока не зарегистрированы.']


def contradictions_report(data, current):
    lines = []
    for row in data['contradictions']:
        lines.append(f"{row['id']} — {row['title']} | {row['status']}")
        for version in row['versions']:
            lines.append(version['text'] + ' | ' + source_text(version['sources']))
        lines.append('Способ проверки: ' + row.get('check_action', 'не определён'))
        if row.get('resolution'):
            resolution = row['resolution']
            lines.append('Результат: ' + resolution['text'] + '; дата ' + resolution['date'] + '; ' + source_text(resolution['sources']))
        reasons = stale_reasons(row, current)
        if reasons:
            lines.append('Статус требует повторной сверки: ' + '; '.join(reasons))
    return lines or ['Противоречия пока не зарегистрированы; это не является выводом об их отсутствии в деле.']


def calculate_deadline(row):
    """Arithmetic only. No statutory duration or exception is inferred."""
    basis = row.get('basis', {})
    require(basis.get('reviewed') is True and all(basis.get(k) for k in ('rule', 'source', 'checked_on', 'applicability')), 'Нет проверенного правила, источника, даты сверки или объяснения применимости')
    date.fromisoformat(basis['checked_on'])
    require(row.get('sources'), 'Нет источника исходного события')
    rule = row.get('calculation', {})
    kind = rule.get('kind')
    trace = []
    if kind == 'date':
        require(row.get('event_on'), 'Нет даты исходного события')
        date.fromisoformat(row['event_on'])
        due = date.fromisoformat(rule['due_on'])
        trace.append('Указанная в источнике дата: ' + due.isoformat())
    else:
        amount = rule.get('amount')
        require(type(amount) is int and amount > 0, 'Длительность должна быть положительным целым числом')
        if kind == 'hours':
            start = datetime.fromisoformat(row['event_on'])
            require(start.tzinfo is not None, 'Для часов нужен часовой пояс события')
            require(rule.get('roll') == 'none', 'Перенос почасового срока не поддерживается автоматически')
            due = start + timedelta(hours=amount)
            trace.append(f'{start.isoformat()} + {amount} часов = {due.isoformat()}')
            return due, trace
        start = date.fromisoformat(row['event_on'])
        if kind == 'days':
            require(type(rule.get('include_start')) is bool, 'Нужно явно выбрать включение начального дня')
            due = start + timedelta(days=amount - int(rule['include_start']))
            trace.append(f"{start.isoformat()} + {amount} календарных дней; начальный день {'включён' if rule['include_start'] else 'исключён'} = {due.isoformat()}")
        elif kind == 'months':
            n = start.year * 12 + start.month - 1 + amount
            year, month = divmod(n, 12)
            month += 1
            last = monthrange(year, month)[1]
            require(rule.get('month_overflow') in ('last_day', 'error'), 'Нужно явное правило отсутствующего числа месяца')
            require(start.day <= last or rule['month_overflow'] == 'last_day', 'В конечном месяце нет такого числа; требуется отдельное решение')
            due = date(year, month, min(start.day, last))
            trace.append(f'{start.isoformat()} + {amount} календарных месяцев = {due.isoformat()}')
        else:
            raise ValueError('Поддержаны date, days, months, hours; специальный расчёт оформите проверенной датой')
    roll = rule.get('roll')
    require(roll in ('none', 'next_workday'), 'Нужно явное правило переноса')
    if roll == 'next_workday':
        cal = rule.get('calendar', {})
        require(cal.get('source') and cal.get('checked_on') and cal.get('start') and cal.get('end'), 'Нужен проверенный календарь и границы покрытия')
        require(isinstance(cal.get('nonworking_dates'), list), 'Нужен явный полный список нерабочих дат')
        begin, end = date.fromisoformat(cal['start']), date.fromisoformat(cal['end'])
        nonworking = {date.fromisoformat(v) for v in cal['nonworking_dates']}
        original = due
        while True:
            require(begin <= due <= end, 'Дата вышла за проверенные границы календаря')
            if due not in nonworking:
                break
            due += timedelta(days=1)
        trace.append('Перенос по представленному календарю: ' + original.isoformat() + ' → ' + due.isoformat())
    return due, trace


def deadlines_report(data, current, as_of, horizon=7):
    require(0 <= horizon <= 366, 'Горизонт должен быть от 0 до 366 дней')
    # An aware datetime permits exact hourly status; a date alone must not assert expiry within that day.
    moment = datetime.fromisoformat(as_of) if 'T' in as_of else None
    if moment:
        require(moment.tzinfo is not None, 'У текущего времени должен быть часовой пояс')
    today = moment.date() if moment else date.fromisoformat(as_of)
    end = today + timedelta(days=horizon)
    lines = [f'Дата просмотра: {as_of}. Ближайшие {horizon} дней по {end.isoformat()} включительно и все более ранние незавершённые сроки.']
    pending = []
    for row in data['deadlines']:
        reasons = stale_reasons(row, current)
        prefix = row['id'] + ' — ' + row['title']
        if reasons:
            pending.append(prefix + ': изменённые или непроверенные источники — ' + '; '.join(reasons))
            continue
        if row.get('completed'):
            continue
        if row.get('superseded_by'):
            lines.append(prefix + ': заменён записью ' + row['superseded_by'] + '; ' + source_text(row['supersession_sources']))
            continue
        try:
            due, trace = calculate_deadline(row)
            due_day = due.date() if isinstance(due, datetime) else due
            if due_day > end:
                continue
            if isinstance(due, datetime) and moment:
                overdue = due < moment
            else:
                overdue = due_day < today
            status = 'ПРОСРОЧЕН ПО ЗАРЕГИСТРИРОВАННЫМ ДАННЫМ' if overdue else ('НА СЕГОДНЯ' if due_day == today else 'ПРЕДСТОИТ')
            lines.append(prefix + ' | ' + due.isoformat() + ' | ' + status)
            lines.extend(trace)
            lines.append('Основание: ' + row['basis']['rule'] + '; ' + row['basis']['source'] + '; сверено ' + row['basis']['checked_on'] + '; применимость: ' + row['basis']['applicability'])
            lines.append('Источники события: ' + source_text(row['sources']))
            if isinstance(due, datetime) and moment is None and due_day == today:
                lines.append('Для определения истечения внутри дня нужно текущее время с часовым поясом.')
        except (ValueError, KeyError, TypeError, OverflowError) as exc:
            pending.append(prefix + ': ' + str(exc))
    if pending:
        lines.append('ТРЕБУЕТ ПРОВЕРКИ — эти записи не исключаются молча из отчёта')
        lines.extend(pending)
    if len(lines) == 1:
        lines.append('В заданном окне зарегистрированных открытых сроков нет. Полнота учёта сроков дела этим не подтверждается.')
    return lines


def citations_report(project, data, current):
    docs = {r['id']: r for r in data['documents']}
    lines = ['Проверяется зарегистрированный перечень ссылок. Совпадение диапазона с описью не доказывает содержание документа. Страницы PDF автоматически в листы дела не переводятся.']
    for row in data['citations']:
        target = docs[row['target_id']]
        reasons = stale_reasons(row, current) + stale_reasons(target, current)
        issues = []
        literal_match = CITATION_RX.fullmatch(row['literal'].strip())
        if literal_match:
            lv, ls, le = literal_match.groups()
            if (int(lv), int(ls), int(le or ls)) != (row['volume'], row['leaf_start'], row['leaf_end']):
                issues.append('буквальная ссылка не совпадает с числовыми полями записи; перепроверьте извлечение')
        else:
            issues.append('составная или нестандартная запись ссылки требует ручного разбора')
        if not all(k in target for k in ('volume', 'leaf_start', 'leaf_end')):
            issues.append('нет проверенной привязки документа к описи')
        elif row['volume'] != target['volume'] or row['leaf_start'] < target['leaf_start'] or row['leaf_end'] > target['leaf_end']:
            issues.append('том или диапазон листов не соответствует зарегистрированному документу')
        try:
            if norm(row['literal']) not in norm(read_text(project, docs[row['document_id']])):
                issues.append('буквальный текст ссылки не найден в текущем проекте')
        except Exception as exc:
            issues.append('текст проекта не проверен: ' + str(exc))
        if reasons:
            issues.append('требуется повторная сверка источников: ' + '; '.join(reasons))
        outcome = ' | '.join(issues) if issues else 'диапазон и наличие текста согласованы; содержательная сверка с первоисточником остаётся обязательной'
        lines.append(f"{row['id']} — {row['document_id']}, {row['locator']}: {row['literal']} → {row['target_id']} — {outcome}")
        if row.get('content_review'):
            lines.append('Записанная содержательная оценка: ' + row['content_review'])
    registered = {d['id']: [] for d in data['documents'] if d['status'] == 'рабочая копия'}
    for row in data['citations']:
        registered.setdefault(row['document_id'], []).append(norm(row['literal']))
    for did, literals in registered.items():
        try:
            text = read_text(project, docs[did])
            candidates = [m.group(0) for m in CITATION_RX.finditer(text)]
            unlisted = [c.strip() for c in candidates if norm(c.strip()) not in literals]
            if unlisted:
                lines.append(did + ': возможные незарегистрированные ссылки — ' + '; '.join(unlisted))
        except Exception as exc:
            lines.append(did + ': поиск незарегистрированных ссылок не выполнен — ' + str(exc))
    if not data['citations']:
        lines.append('Ссылки не зарегистрированы; проверка не выполнена.')
    return lines


def summary(data, current):
    lines = [f"Проект: {data['case']['label']}. Редакция реестра: {data['revision']}.", 'Реестр отражает зарегистрированный объём. Контрольная сумма подтверждает неизменность файла, а не достоверность его содержания.']
    for doc in data['documents']:
        lines.append(f"{doc['id']} — {doc['title']} | {doc['status']} | {current[doc['id']]} | {doc['path']}")
    for fact in data['facts']:
        lines.append(f"{fact['id']} — {fact['text']} | {fact['status']} | {source_text(fact.get('sources', []))}")
        if stale_reasons(fact, current):
            lines.append('Требует повторной проверки: ' + '; '.join(stale_reasons(fact, current)))
    for task in data['tasks']:
        actual = task.get('actual')
        lines.append(f"{task['id']} — {task['title']} | подготовка: {task['preparation']} | факт: " + (actual['status'] + ' ' + actual['date'] + '; ' + source_text(actual['sources']) if actual else 'не подтверждён'))
        if stale_reasons(task, current):
            lines.append('Статус требует сверки: ' + '; '.join(stale_reasons(task, current)))
        lines.append('Следующий шаг: ' + task.get('next_step', 'не указан'))
    lines.append('Открытых противоречий: ' + str(sum(r['status'] != 'разрешено' for r in data['contradictions'])))
    return lines
