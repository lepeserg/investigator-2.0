# -*- coding: utf-8 -*-
"""Read DOCX paragraphs and tables in document order, labeling tracked changes.

    python extract_docx.py "file.docx" "folder" -o out.txt
    python extract_docx.py "file.docx" --revisions marked|current|original

Without a selected view, tracked changes stop reading until the user chooses.
Selecting a view does not accept changes or establish an approved document.
Images need visual inspection.
"""
import argparse
import glob
import os
import sys
from word_text import read_docx, VIEWS, RevisionChoiceRequired

os.environ.setdefault('PYTHONUTF8', '1')
try:
    sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    pass


def dump(path, out, revisions=None):
    """Return whether text was read; warn explicitly about tracked changes."""
    out.write('\n' + '=' * 80 + '\nФАЙЛ: ' + os.path.basename(path) + '\n' + '=' * 80 + '\n')
    try:
        text, metadata = read_docx(path, revisions)
    except RevisionChoiceRequired as exc:
        out.write('!! ТРЕБУЕТСЯ ВЫБОР РЕДАКЦИИ: %s\n' % exc)
        return False
    except Exception as exc:
        out.write('!! ошибка: %s\n' % exc)
        return False
    if metadata['revisions']:
        out.write('!! ИСПРАВЛЕНИЯ WORD: режим ' + revisions + '; ' +
                  ', '.join(key + '=' + str(value) for key, value in metadata['revisions'].items()) +
                  '. Это представление для чтения, не подтверждение окончательной редакции.\n')
    out.write(text)
    return True


def collect(args):
    """Preserve the previous helper contract: paths and optional -o -> pair."""
    files, out_path, i = [], None, 0
    while i < len(args):
        item = args[i]
        if item == '-o' and i + 1 < len(args):
            out_path = args[i + 1]
            i += 2
            continue
        if os.path.isdir(item):
            files += sorted(glob.glob(os.path.join(item, '**', '*.docx'), recursive=True))
        else:
            files.append(item)
        i += 1
    return files, out_path


def main(argv):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('paths', nargs='+')
    parser.add_argument('-o', dest='output')
    parser.add_argument('--revisions', choices=VIEWS, default=None,
                        help='только после выбора пользователя; без выбора исправления останавливают чтение')
    if argv and argv[0] in ('/?', 'help'):
        argv = ['--help']
    args = parser.parse_args(argv)
    files, _ = collect(args.paths)
    files = [path for path in files if not os.path.basename(path).startswith('~$')]
    bad = 0
    if args.output:
        with open(args.output, 'w', encoding='utf-8') as output:
            for path in files:
                if not dump(path, output, args.revisions):
                    bad += 1
        print('Записано %d файл(ов) в %s' % (len(files), args.output))
    else:
        for path in files:
            if not dump(path, sys.stdout, args.revisions):
                bad += 1
    if bad:
        print('!! не удалось открыть файлов: %d из %d' % (bad, len(files)), file=sys.stderr)
        return 1
    if not files:
        print('Ни одного .docx не найдено.', file=sys.stderr)
        return 2
    return 0


if __name__ == '__main__':
    sys.exit(main(sys.argv[1:]))
