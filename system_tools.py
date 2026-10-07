"""Discover Windows tools without changing the machine or the user's profile."""
from pathlib import Path
import os
import shutil
import subprocess

PACKAGES = {
    'tesseract': 'tesseract-ocr.tesseract',
    'ffmpeg': 'Gyan.FFmpeg',
    'soffice': 'TheDocumentFoundation.LibreOffice',
}


def search_paths():
    paths = []
    for name in ('ProgramFiles', 'ProgramFiles(x86)', 'ProgramW6432', 'LOCALAPPDATA'):
        base = os.environ.get(name)
        if base:
            root = Path(base)
            paths.extend((root/'Tesseract-OCR', root/'LibreOffice'/'program'))
    local = os.environ.get('LOCALAPPDATA')
    if local:
        packages = Path(local)/'Microsoft'/'WinGet'/'Packages'
        paths.append(Path(local)/'Microsoft'/'WinGet'/'Links')
        if packages.exists():
            paths.extend(p.parent for p in packages.glob('Gyan.FFmpeg*/**/ffmpeg.exe'))
    return list(dict.fromkeys(p for p in paths if p.is_dir()))


def environment(config, root):
    env = dict(os.environ)
    paths = [Path(p) for p in config['tools']['extra_path']]
    paths.extend(search_paths() if os.name == 'nt' else [])
    env['PATH'] = os.pathsep.join(str(p) for p in paths if p.is_dir()) + os.pathsep + env.get('PATH', '')
    if config['tools']['tessdata_prefix']:
        env['TESSDATA_PREFIX'] = config['tools']['tessdata_prefix']
    elif not env.get('TESSDATA_PREFIX'):
        data = root/'runtime'/'tessdata'
        if all((data/f'{language}.traineddata').is_file() for language in ('rus', 'eng', 'osd')):
            env['TESSDATA_PREFIX'] = str(data)
    return env


def find(name, env):
    return shutil.which(name, path=env['PATH'])


def word_registered():
    if os.name != 'nt':
        return False
    import winreg
    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, 'Word.Application'):
            return True
    except OSError:
        return False


def probe(name, env):
    path = find(name, env)
    if not path:
        return False
    try:
        return subprocess.run([path, '--version'], env=env, capture_output=True, timeout=30).returncode == 0
    except (OSError, subprocess.TimeoutExpired):
        return False
