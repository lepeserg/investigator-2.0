"""CLI for local case management. Every successful command exports a Word report."""
import argparse
import json
from case_store import init, load, apply, strict_json, freshness
from case_checks import summary, impact, requests_report, contradictions_report, deadlines_report, citations_report
from case_reports import export

MODES = ('summary', 'requests', 'contradictions', 'deadlines', 'citations', 'all')


def sections(project, data, mode, as_of, horizon=7):
    current = freshness(project, data)
    result = []
    if mode in ('summary', 'all'):
        result.append(('Состояние дела и следующий шаг', summary(data, current)))
    if mode in ('requests', 'all'):
        result.append(('Запросы и ответы', requests_report(data, current)))
    if mode in ('contradictions', 'all'):
        result.append(('Противоречия и результаты проверки', contradictions_report(data, current)))
    if mode in ('deadlines', 'all'):
        if not as_of:
            raise ValueError('Для сроков явно укажите --as-of: дату или время с часовым поясом')
        result.append(('Сроки', deadlines_report(data, current, as_of, horizon)))
    if mode in ('citations', 'all'):
        result.append(('Ссылки на тома и листы', citations_report(project, data, current)))
    result.append(('Источники и пределы проверки', ['Проверка выполнена только по зарегистрированным источникам. Содержательные оценки в реестре вносит помощник после чтения материалов.', 'Документы: ' + str(len(data['documents'])) + '; из них изменённых или непроверенных: ' + str(sum(v != 'совпадает' for v in current.values()))]))
    return result


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--project', required=True, help='Существующая папка проекта')
    sub = parser.add_subparsers(dest='command', required=True)
    p = sub.add_parser('init', help='Создать пустой реестр без сканирования корпуса')
    p.add_argument('--label', required=True)
    p = sub.add_parser('apply', help='Внести проверенные записи JSON с проверкой редакции')
    p.add_argument('--patch', required=True)
    p = sub.add_parser('report', help='Прочитать реестр и создать новый Word-отчёт')
    p.add_argument('--mode', choices=MODES, default='summary')
    p.add_argument('--as-of', help='YYYY-MM-DD либо ISO дата и время с часовым поясом')
    p.add_argument('--horizon', type=int, default=7)
    p = sub.add_parser('impact', help='Показать последствия предложенного изменения, без замен')
    p.add_argument('--fact', required=True)
    p.add_argument('--old', required=True)
    p.add_argument('--new', required=True)
    args = parser.parse_args(argv)
    try:
        if args.command == 'init':
            data = init(args.project, args.label)
        elif args.command == 'apply':
            data = apply(args.project, strict_json(args.patch))
        else:
            data = load(args.project)
        if args.command == 'impact':
            result = [('Последствия предложенного изменения', impact(args.project, data, freshness(args.project, data), args.fact, args.old, args.new))]
        else:
            mode = args.mode if args.command == 'report' else 'summary'
            result = sections(args.project, data, mode, getattr(args, 'as_of', None), getattr(args, 'horizon', 7))
        path = export(args.project, data, args.command, result)
        print(json.dumps({'revision': data['revision'], 'report': str(path), 'visual_review': 'требуется'}, ensure_ascii=False))
        return 0
    except (OSError, ValueError, KeyError, TypeError) as exc:
        print(json.dumps({'error': str(exc), 'note': 'При ошибке экспорта после apply перечитайте номер редакции: изменение реестра могло сохраниться.'}, ensure_ascii=False))
        return 1


if __name__ == '__main__':
    raise SystemExit(main())
