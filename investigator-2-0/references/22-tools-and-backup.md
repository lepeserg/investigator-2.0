# 22. Инструменты и рабочие копии

DOCX читает общий модуль `scripts/word_text.py`, сохраняя порядок абзацев и таблиц. При непринятых исправлениях сначала остановись и уточни редакцию; `--revisions current|original|marked` допускается после явного выбора. Поиск, пропустивший такие файлы, сообщает неполный результат (код 2).

Публичная поставка запускается через `python run.py <имя_скрипта.py> <аргументы>`. Перед незнакомой командой прочитай её `--help`. Внешние утилиты прежней рабочей станции в комплект не входят.

| Задача | Скрипты |
|---|---|
| Проверка комплекта и среды | `scripts/check_skill.py`, `scripts/check_env.py` |
| Извлечение Word и восстановление файла | `scripts/extract_docx.py`, `scripts/docx_recover.py` |
| Адресная правка и контроль Word | `scripts/docx_edit.py`, `scripts/docx_integrity.py`, `scripts/doc_meta.py` |
| Поиск и согласованность | `scripts/corpus_search.py`, `scripts/consistency_check.py` |
| Текст, формат и стиль | `scripts/style_lint.py`, `scripts/format_lint.py` |
| Генераторы проектов | `scripts/make_docx.py`, `scripts/make_povestka.py`, `scripts/make_prodlenie_krsp.py`, `scripts/make_tablica_nedostatkov.py` |
| Описи и тома | `scripts/make_opis_toma.py`, `scripts/opis_verify.py`, `scripts/check_tom.py` |
| Учётные формы | `scripts/make_statcards.py`, `scripts/fill_formfields.py`, `scripts/fix_statcard_font.py`, `scripts/post_release_gvp.py` |
| PDF и OCR | `scripts/pdf_ingest.py`, `scripts/pdf_no_text.py`, `scripts/ocr_smart.py` |
| Аудио и доступ к модели | `scripts/av_ingest.py`, `scripts/setup_hf_access.py` |
| Дополнительные проверки | `scripts/check_account.py`, `scripts/legal_rights.py` |

Сохраняй исходник, выбранную рабочую копию и краткий журнал существенных правок в папке проекта. Перед массовой правкой сохрани предыдущую редакцию; после неё проверь содержательную полноту. Не закрывай рабочий экземпляр Word пользователя через общий COM-сеанс. Файл, открытый в Word, не перезаписывай без решения конфликта.

Наличие команды не означает наличие внешнего инструмента, модели или актуального бланка. Проверяй зависимости по задаче. Контроль структуры не подтверждает юридическое содержание, а контрольный ключ счёта — его принадлежность.
