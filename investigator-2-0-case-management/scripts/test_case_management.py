"""Synthetic regression scenarios; no case corpus, network or Office sessions used."""
from copy import deepcopy
from datetime import date
import json
from pathlib import Path
import tempfile
import unittest
import zipfile

from case_store import init, apply, load, freshness, stale_reasons, digest, validate
from case_checks import calculate_deadline, deadlines_report, impact, requests_report, contradictions_report, citations_report, read_text
from case_reports import export


class CaseManagementTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.project = Path(self.temp.name)
        (self.project / 'source.txt').write_text('Выписка: сумма 100. Точную дату не помню.', encoding='utf-8')
        (self.project / 'draft.txt').write_text('Сумма 100. Ссылка т. 1 л.д. 5–6. Точную дату не помню.', encoding='utf-8')
        (self.project / 'inventory.txt').write_text('Выписка: том 1 листы 5–8', encoding='utf-8')
        init(self.project, 'Условное дело для испытаний')
        self.patch({'documents': [
            {'id': 'D1', 'title': 'Выписка', 'path': 'source.txt', 'status': 'скан', 'volume': 1, 'leaf_start': 5, 'leaf_end': 8, 'inventory_sources': [self.ref('I1')]},
            {'id': 'D2', 'title': 'Проект', 'path': 'draft.txt', 'status': 'рабочая копия', 'fact_ids': ['F1']},
            {'id': 'I1', 'title': 'Опись', 'path': 'inventory.txt', 'status': 'подписанный документ'}],
            'facts': [{'id': 'F1', 'text': 'Сумма 100', 'status': 'сведения документа', 'sources': [self.ref()]}]}, ack=['D1', 'D2', 'I1'])

    def ref(self, did='D1'):
        return {'document_id': did, 'locator': 'строка 1 условного материала'}

    def patch(self, rows, ack=None):
        return apply(self.project, {'expected_revision': load(self.project)['revision'], 'reason': 'Синтетическая проверка', 'upsert': rows, 'acknowledge_documents': ack or []})

    def deadline(self, **kwargs):
        row = {'id': 'L1', 'title': 'Условный срок', 'event_on': '2026-10-01', 'sources': [self.ref()],
               'basis': {'reviewed': True, 'rule': 'Синтетическое правило для теста', 'source': 'Условный тест без ссылки на закон', 'checked_on': '2026-10-10', 'applicability': 'Только арифметическое испытание'},
               'calculation': {'kind': 'days', 'amount': 5, 'include_start': False, 'roll': 'none'}}
        row.update(kwargs)
        return row

    def test_init_does_not_overwrite(self):
        before = digest(self.project / '_ведение_дела' / 'state.json')
        with self.assertRaises(ValueError): init(self.project, 'Другое дело')
        self.assertEqual(before, digest(self.project / '_ведение_дела' / 'state.json'))

    def test_old_revision_rejected(self):
        with self.assertRaises(ValueError): apply(self.project, {'expected_revision': 0, 'reason': 'устаревшая операция'})

    def test_atomic_invalid_update(self):
        before = load(self.project)
        with self.assertRaises(ValueError): self.patch({'tasks': [{'id': 'T1', 'title': 'Допрос', 'preparation': 'проект после проверки', 'actual': {'status': 'проведено', 'date': '2026-10-10'}}]})
        self.assertEqual(before, load(self.project))

    def test_unknown_source_rejected(self):
        with self.assertRaises(ValueError): self.patch({'facts': [{'id': 'F2', 'text': 'Факт', 'status': 'подтверждено', 'sources': [self.ref('MISSING')]}]})

    def test_missing_source_reported(self):
        (self.project / 'source.txt').unlink()
        self.assertTrue(freshness(self.project, load(self.project))['D1'].startswith('недоступен'))

    def test_changed_source_detected(self):
        (self.project / 'source.txt').write_text('Сумма 200', encoding='utf-8')
        self.assertEqual(freshness(self.project, load(self.project))['D1'], 'изменён')

    def test_recheck_source_does_not_recheck_assertion(self):
        (self.project / 'source.txt').write_text('Сумма 200', encoding='utf-8')
        data = self.patch({}, ack=['D1'])
        self.assertEqual(freshness(self.project, data)['D1'], 'совпадает')
        self.assertTrue(stale_reasons(data['facts'][0], freshness(self.project, data)))

    def test_recheck_fact_clears_stale(self):
        (self.project / 'source.txt').write_text('Сумма 200', encoding='utf-8')
        row = {'id': 'F1', 'text': 'Сумма 200', 'status': 'сведения документа', 'sources': [self.ref()]}
        data = self.patch({'facts': [row]}, ack=['D1'])
        self.assertFalse(stale_reasons(data['facts'][0], freshness(self.project, data)))

    def test_history_preserves_previous(self):
        self.patch({})
        previous = json.loads((self.project / '_ведение_дела/history/revision-000001.json').read_text(encoding='utf-8'))
        self.assertEqual(previous['revision'], 1)

    def test_closed_contradiction_requires_reason(self):
        row = {'id': 'C1', 'title': 'Суммы', 'status': 'разрешено', 'versions': [{'text': '100', 'sources': [self.ref()]}, {'text': '200', 'sources': [self.ref('D2')]}]}
        with self.assertRaises(ValueError): self.patch({'contradictions': [row]})

    def test_impact_read_only_and_linked(self):
        data = load(self.project)
        paths = [self.project / x for x in ('source.txt', 'draft.txt', '_ведение_дела/state.json')]
        before = [digest(p) for p in paths]
        report = '\n'.join(impact(self.project, data, freshness(self.project, data), 'F1', '100', '200'))
        self.assertIn('D2', report)
        self.assertEqual(before, [digest(p) for p in paths])

    def test_impact_incomplete_explicit(self):
        (self.project / 'source.txt').unlink()
        data = load(self.project)
        self.assertIn('Поиск неполон', '\n'.join(impact(self.project, data, freshness(self.project, data), 'F1', '100', '200')))

    def test_numeric_substring_not_match(self):
        (self.project / 'source.txt').write_text('Сумма 1000', encoding='utf-8')
        data = load(self.project)
        lines = impact(self.project, data, freshness(self.project, data), 'F1', '100', '200')
        self.assertFalse(any(x.startswith('D1 —') for x in lines))

    def test_request_partial_period_and_attachment(self):
        row = {'id': 'R1', 'title': 'Запрос выписки', 'document_id': 'D2', 'items': [{'id': '1', 'question': 'Операции за период', 'period': ['2026-01-01', '2026-03-31'], 'covered_period': ['2026-02-01', '2026-03-31'], 'assessment': 'частичный', 'assessment_note': 'Нет января', 'sources': [self.ref()]}], 'responses': [{'document_id': 'D1', 'received_on': '2026-10-09', 'attachments': [{'title': 'Приложение 1'}]}]}
        data = self.patch({'requests': [row]})
        text = '\n'.join(requests_report(data, freshness(self.project, data)))
        self.assertIn('не покрывает', text)
        self.assertIn('не представлено', text)
        self.assertIn('Направление не подтверждено', text)

    def test_request_unassessed_not_complete(self):
        row = {'id': 'R1', 'title': 'Запрос', 'document_id': 'D2', 'items': [{'id': '1', 'question': 'Операции'}], 'responses': [{'document_id': 'D1', 'received_on': '2026-10-09'}]}
        data = self.patch({'requests': [row]})
        self.assertIn('полнота не оценена', '\n'.join(requests_report(data, freshness(self.project, data))))

    def test_invalid_period_rejected(self):
        row = {'id': 'R1', 'title': 'Запрос', 'document_id': 'D2', 'items': [{'id': '1', 'question': 'Операции', 'period': ['2026-03-31', '2026-01-01']}]}
        with self.assertRaises(ValueError): self.patch({'requests': [row]})

    def test_days_explicit_start(self):
        row = self.deadline()
        self.assertEqual(calculate_deadline(row)[0], date(2026, 10, 6))
        row['calculation']['include_start'] = True
        self.assertEqual(calculate_deadline(row)[0], date(2026, 10, 5))

    def test_month_end_leap(self):
        row = self.deadline(event_on='2024-01-31', calculation={'kind': 'months', 'amount': 1, 'roll': 'none', 'month_overflow': 'last_day'})
        self.assertEqual(calculate_deadline(row)[0], date(2024, 2, 29))

    def test_missing_legal_basis_not_calculated(self):
        row = self.deadline(basis={})
        with self.assertRaises(ValueError): calculate_deadline(row)

    def test_no_calendar_guess(self):
        row = self.deadline()
        row['calculation']['roll'] = 'next_workday'
        with self.assertRaises(ValueError): calculate_deadline(row)

    def test_calendar_roll_and_boundary(self):
        row = self.deadline(event_on='2026-10-01', calculation={'kind': 'days', 'amount': 2, 'include_start': False, 'roll': 'next_workday', 'calendar': {'source': 'synthetic', 'checked_on': '2026-10-01', 'start': '2026-10-01', 'end': '2026-10-05', 'nonworking_dates': ['2026-10-03', '2026-10-04']}})
        self.assertEqual(calculate_deadline(row)[0], date(2026, 10, 5))
        row['calculation']['calendar']['end'] = '2026-10-04'
        with self.assertRaises(ValueError): calculate_deadline(row)

    def test_hours_require_timezone(self):
        row = self.deadline(event_on='2026-10-01T12:00:00', calculation={'kind': 'hours', 'amount': 48, 'roll': 'none'})
        with self.assertRaises(ValueError): calculate_deadline(row)

    def test_horizon_includes_all_overdue_and_seven_days(self):
        a = self.deadline()
        b = self.deadline(id='L2', event_on='2026-10-12')
        c = self.deadline(id='L3', event_on='2026-10-13')
        data = self.patch({'deadlines': [a, b, c]})
        text = '\n'.join(deadlines_report(data, freshness(self.project, data), '2026-10-10'))
        self.assertIn('L1 —', text)
        self.assertIn('L2 —', text)
        self.assertNotIn('L3 —', text)
        self.assertIn('ПРОСРОЧЕН', text)

    def test_unverified_deadline_visible(self):
        data = self.patch({'deadlines': [self.deadline(basis={})]})
        text = '\n'.join(deadlines_report(data, freshness(self.project, data), '2026-10-10'))
        self.assertIn('ТРЕБУЕТ ПРОВЕРКИ', text)
        self.assertIn('L1', text)

    def test_changed_deadline_not_claimed_overdue(self):
        data = self.patch({'deadlines': [self.deadline()]})
        (self.project / 'source.txt').write_text('Иная дата', encoding='utf-8')
        text = '\n'.join(deadlines_report(data, freshness(self.project, data), '2026-10-10'))
        self.assertIn('ТРЕБУЕТ ПРОВЕРКИ', text)
        self.assertNotIn('ПРОСРОЧЕН', text)

    def test_hourly_date_only_not_premature_overdue(self):
        row = self.deadline(event_on='2026-10-10T08:00:00+03:00', calculation={'kind': 'hours', 'amount': 2, 'roll': 'none'})
        data = self.patch({'deadlines': [row]})
        current = freshness(self.project, data)
        self.assertNotIn('ПРОСРОЧЕН', '\n'.join(deadlines_report(data, current, '2026-10-10')))
        self.assertIn('ПРОСРОЧЕН', '\n'.join(deadlines_report(data, current, '2026-10-10T12:00:00+03:00')))

    def test_completed_deadline_excluded(self):
        row = self.deadline(completed={'date': '2026-10-05', 'status': 'действие проведено', 'sources': [self.ref()]})
        data = self.patch({'deadlines': [row]})
        self.assertNotIn('L1 —', '\n'.join(deadlines_report(data, freshness(self.project, data), '2026-10-10')))

    def citation(self, **kwargs):
        row = {'id': 'Q1', 'document_id': 'D2', 'target_id': 'D1', 'locator': 'абзац 1', 'literal': 'т. 1 л.д. 5–6', 'volume': 1, 'leaf_start': 5, 'leaf_end': 6}
        row.update(kwargs)
        return row

    def test_correct_citation_no_invented_error(self):
        data = self.patch({'citations': [self.citation()]})
        text = '\n'.join(citations_report(self.project, data, freshness(self.project, data)))
        self.assertIn('диапазон и наличие текста согласованы', text)
        self.assertNotIn('не соответствует', text)

    def test_wrong_volume(self):
        data = self.patch({'citations': [self.citation(volume=2)]})
        self.assertIn('не соответствует', '\n'.join(citations_report(self.project, data, freshness(self.project, data))))

    def test_outside_leaf_range(self):
        data = self.patch({'citations': [self.citation(leaf_end=10)]})
        self.assertIn('не соответствует', '\n'.join(citations_report(self.project, data, freshness(self.project, data))))

    def test_literal_mismatch_detected(self):
        data = self.patch({'citations': [self.citation(literal='т. 2 л.д. 5–6')]})
        self.assertIn('не совпадает с числовыми', '\n'.join(citations_report(self.project, data, freshness(self.project, data))))

    def test_unregistered_citation_discovered(self):
        data = load(self.project)
        self.assertIn('незарегистрированные ссылки', '\n'.join(citations_report(self.project, data, freshness(self.project, data))))

    def test_inventory_recheck_does_not_revalidate_mapping(self):
        self.patch({'citations': [self.citation()]})
        (self.project / 'inventory.txt').write_text('Выписка теперь том 2 листы 10–12', encoding='utf-8')
        data = self.patch({}, ack=['I1'])
        self.assertIn('версия источника не подтверждена', '\n'.join(citations_report(self.project, data, freshness(self.project, data))))

    def test_cycle_of_deadlines_rejected(self):
        a = self.deadline(superseded_by='L2', supersession_sources=[self.ref()])
        b = self.deadline(id='L2', superseded_by='L1', supersession_sources=[self.ref()])
        with self.assertRaises(ValueError): self.patch({'deadlines': [a, b]})

    def test_corrupt_docx_search_is_incomplete(self):
        (self.project / 'bad.docx').write_bytes(b'not a docx')
        data = self.patch({'documents': [{'id': 'BAD', 'title': 'Повреждённый файл', 'path': 'bad.docx', 'status': 'рабочая копия'}]})
        self.assertIn('Поиск неполон', '\n'.join(impact(self.project, data, freshness(self.project, data), 'F1', '100', '200')))

    def test_track_changes_stop_reading(self):
        from docx import Document
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn
        path = self.project / 'tracked.docx'
        doc = Document()
        p = doc.add_paragraph('Начало ')
        ins = OxmlElement('w:ins')
        ins.set(qn('w:id'), '1')
        run, text = OxmlElement('w:r'), OxmlElement('w:t')
        text.text = 'новый текст'
        run.append(text); ins.append(run); p._p.append(ins)
        doc.save(path)
        with self.assertRaises(ValueError): read_text(self.project, {'path': str(path)})

    def test_uncertainty_preserved_in_impact_excerpt(self):
        data = load(self.project)
        text = '\n'.join(impact(self.project, data, freshness(self.project, data), 'F1', '100', '200'))
        self.assertIn('точную дату не помню', text)

    def test_report_separate_files(self):
        data = load(self.project)
        first = export(self.project, data, 'test', [('Результат', ['Проверка'])])
        second = export(self.project, data, 'test', [('Результат', ['Проверка'])])
        self.assertNotEqual(first, second)
        self.assertTrue(zipfile.is_zipfile(first))
        self.assertTrue(zipfile.is_zipfile(second))


if __name__ == '__main__':
    unittest.main()
