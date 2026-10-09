"""Run a bundled tool in the isolated public-edition environment."""
from pathlib import Path
import argparse
import json
import subprocess
import sys
import bootstrap
import local_config
import system_tools

ROOT = Path(__file__).resolve().parent


def environment(config):
    runtime = ROOT/'runtime'
    temp = runtime/'temp'
    temp.mkdir(parents=True, exist_ok=True)
    env = system_tools.environment(config, ROOT)
    paths = [bootstrap.python_path().parent]
    env['PATH'] = system_tools.prepend_path((p for p in paths if p.exists()), env.get('PATH', ''))
    env.update(PYTHONIOENCODING='utf-8', PYTHONUTF8='1', PYTHONDONTWRITEBYTECODE='1',
               TEMP=str(temp), TMP=str(temp), HF_HOME=str(runtime/'models'),
               HF_HUB_DISABLE_TELEMETRY='1', PYANNOTE_METRICS_ENABLED='0',
               HF_HUB_DISABLE_SYMLINKS_WARNING='1')
    if config['tools']['tessdata_prefix']:
        env['TESSDATA_PREFIX'] = config['tools']['tessdata_prefix']
    # Explicit legacy profiles take precedence. Generated data stays outside Git.
    if not env.get('SK_CONSTANTS'):
        profile = runtime/'local-constants.json'
        profile.write_text(json.dumps(local_config.legacy_constants(config), ensure_ascii=False, indent=2), encoding='utf-8')
        env['SK_CONSTANTS'] = str(profile)
    return env


def select_script(name):
    scripts = (ROOT/'investigator-2-0'/'scripts').resolve()
    path = (scripts/name).resolve()
    if path.parent != scripts or path.suffix != '.py' or not path.is_file():
        raise ValueError('Select an existing Python script directly inside investigator-2-0/scripts')
    return path


MISSING_PACKAGES = ('Python-окружение не подготовлено или неполное. run.py по умолчанию ничего не устанавливает.\n'
                    'Запустите start.cmd или явно разрешите установку: py -3.12 run.py --install <скрипт> ...')


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--install', action='store_true',
                        help='Allow installing missing Python packages from PyPI (off by default)')
    parser.add_argument('--offline', action='store_true',
                        help='Accepted for compatibility; offline is now the default')
    parser.add_argument('tool', help='Script name, e.g. extract_docx.py')
    parser.add_argument('args', nargs=argparse.REMAINDER)
    args = parser.parse_args()
    if args.install and args.offline:
        parser.error('--install and --offline are mutually exclusive')
    try:
        script = select_script(args.tool)
        config = local_config.load()
        try:
            python = bootstrap.ensure('base', offline=not args.install)
        except RuntimeError as exc:
            if args.install:
                raise
            print(f'Cannot run tool: {exc}\n{MISSING_PACKAGES}', file=sys.stderr)
            return 2
        return subprocess.call([str(python), str(script), *args.args], env=environment(config))
    except (OSError, ValueError, RuntimeError, subprocess.CalledProcessError) as exc:
        print(f'Cannot run tool: {exc}', file=sys.stderr)
        return 1


if __name__ == '__main__':
    raise SystemExit(main())
