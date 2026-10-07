"""Create project-local Codex entrypoints without copying private data or global skills."""
from pathlib import Path
import argparse
import json

ROOT = Path(__file__).resolve().parent
NAMES = ['investigator-2-0'] + ['investigator-2-0-' + suffix for suffix in
    ('analysis', 'correspondence', 'interrogations', 'actions', 'decisions', 'charging', 'stat-cards')]
MARKER = '<!-- investigator-public-entrypoint -->'


def connect(root=ROOT):
    """Write only our own wrappers; stop before modifying any unrelated entrypoint."""
    files = []
    for name in NAMES:
        source = root/name/'SKILL.md'
        content = source.read_text(encoding='utf-8')
        if not content.startswith('---\n'):
            raise ValueError('Missing skill metadata: ' + name)
        front = content.split('---', 2)[1].strip()
        target = root/'.agents'/'skills'/name/'SKILL.md'
        if not target.resolve().is_relative_to(root.resolve()):
            raise ValueError('Entrypoint path must stay inside this project')
        if target.exists() and MARKER not in target.read_text(encoding='utf-8'):
            raise FileExistsError('Existing skill is not managed by this installer: ' + str(target))
        text = '---\n' + front + '\n---\n\n' + MARKER + '\n\n'
        text += 'Открой и примени [инструкции этого направления](../../../' + name + '/SKILL.md). '
        text += 'Ресурсы читай относительно исходного каталога навыка, а не этой точки входа. '
        text += 'Личные настройки берутся из config.local.json в корне комплекта; материалы дела хранятся вне репозитория.\n'
        files.append((target, text))
    for target, text in files:
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text(text, encoding='utf-8')
    return len(files)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.parse_args()
    try:
        count = connect()
    except (OSError, ValueError) as exc:
        print('Skills were not connected: ' + str(exc))
        return 1
    print(f'Created {count} project-local entrypoints. Open this folder as a Codex project.')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
