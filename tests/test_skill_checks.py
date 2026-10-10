"""Проверки check_skill.py, которые относятся к навыкам и справочникам (синтетический навык)."""
from pathlib import Path
import sys
import tempfile
import unittest

ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'investigator-2-0/scripts'))

class SkillCheckTests(unittest.TestCase):
    def setUp(self):
        (ROOT/'.test-tmp').mkdir(exist_ok=True)
        temp=tempfile.TemporaryDirectory(dir=ROOT/'.test-tmp')
        self.addCleanup(temp.cleanup);self.folder=Path(temp.name)

    def synthetic_skill(self,files):
        root=self.folder/'skill';(root/'references').mkdir(parents=True,exist_ok=True)
        (root/'SKILL.md').write_text('---\nname: skill\ndescription: "Синтетический навык."\n---\n# Навык\n\n### Шаг 6. Сверка права\n\nТекст.\n\n### Шаг 9. Подготовка\n\nТекст.\n',encoding='utf-8')
        for name,text in files.items():(root/'references'/name).write_text(text,encoding='utf-8')
        return root

    def test_check_skill_flags_empty_sections(self):
        import check_skill
        root=self.synthetic_skill({
            'a.md':'# A\n\n## Паспорт\n\n### 3. ФАТАЛЬНЫЕ ОШИБКИ ЖАНРА\n\n### 4. ПРАВИЛА\n\nТекст.\n\n## Карта правил\n\n---\n\n## Полный текст\n\nТекст.\n',
            'b.md':'# B\n\n## Раздел\n\n### Подраздел\n\nТекст.\n\n## Код\n\n```python\n# не заголовок\n```\n\n## Последний\n\nТекст.\n',
            'legacy-dispatcher.md':'# Архив\n\n## Пустой в архиве\n\n## Ещё\n\nТекст.\n'})
        empty=[p for p in check_skill.check(str(root))[0] if 'пустой раздел' in p]
        self.assertEqual(len(empty),2,empty)
        self.assertTrue(any(p.startswith('a.md:5:') and 'ФАТАЛЬНЫЕ ОШИБКИ' in p for p in empty))
        self.assertTrue(any(p.startswith('a.md:11:') and 'Карта правил' in p for p in empty))

    def test_check_skill_flags_dangling_step_references(self):
        import check_skill
        root=self.synthetic_skill({
            'a.md':'# A\n\nПаспорт читается полностью (Шаг 3 скилла).\n\nДонор — Шаг 9.0-bis; сверка — Шаг 6; правка — в Шаге 9.\n\nСпрашивай по Шагу 8.\n\n```\nШаг 5 в коде не проверяется\n```\n',
            'b.md':'# B\n\n### Шаг 1. Свой алгоритм\n\nТекст.\n\n**Шаг 2. Ещё.** Нарушение Шага 1 и Шага 2 — локальные ссылки.\n',
            'legacy-dispatcher.md':'# Архив\n\nШаг 3 и Шаг 9.0-bis живут здесь.\n'})
        steps=[p for p in check_skill.check(str(root))[0] if '«Шаг' in p]
        self.assertEqual(sorted(p.split('«Шаг ')[1].split('»')[0] for p in steps),['3','8','9.0-bis'],steps)
        self.assertTrue(all(p.startswith('a.md:') for p in steps),steps)

    def test_check_skill_flags_dangling_razdel_references(self):
        import check_skill
        root=self.synthetic_skill({
            'a.md':'# A\n\n## 2.12. Есть\n\nТекст.\n\n## 08.22. Есть\n\nТекст.\n',
            'b.md':'# B\n\nСм. (раздел 2.13). Также раздел 2.12 и разделы 08.22.\n\nРаздел 3.2 карточки и раздел 17.4 — однозначная часть, не проверяется.\n'})
        bad=[p for p in check_skill.check(str(root))[0] if 'несуществующий раздел' in p]
        self.assertEqual(bad,['ссылка на несуществующий раздел: §02.13'],bad)

    def test_check_skill_cards_checked_in_split_parts(self):
        import check_skill,tempfile
        with tempfile.TemporaryDirectory() as d:
            Path(d,'05-genres-lifecycle.md').write_text('# 05\n\n## Карта файла\n\nТекст.\n',encoding='utf-8')
            Path(d,'05-genres-lifecycle-vud.md').write_text('# Часть\n\n## 05.01. ВУД\n\nТекст без карточки.\n',encoding='utf-8')
            Path(d,'05b-other.md').write_text('## 05.99. Чужой файл\n\nТекст.\n',encoding='utf-8')
            self.assertIn('05-genres-lifecycle-vud.md',check_skill._card_files(d))
            self.assertNotIn('05b-other.md',check_skill._card_files(d))
            problems=check_skill._check_cards(d)
            self.assertTrue(any('05-genres-lifecycle-vud.md' in p and '§05.01' in p for p in problems),problems)

if __name__=='__main__':unittest.main()
