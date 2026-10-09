"""Install a selected Windows capability; --check is read-only and never downloads."""
from pathlib import Path
from datetime import datetime, timezone
import argparse
import hashlib
import json
import os
import shutil
import subprocess
import sys
import tempfile
import urllib.request

import bootstrap
import local_config
import system_tools

ROOT = Path(__file__).resolve().parent
PROFILES = ('base', 'documents', 'ocr', 'audio', 'all')
REVISION = '87416418657359cb625c412a48b6e1d6d41c29bd'
LANGUAGES = {
    'rus': 'e16e5e036cce1d9ec2b00063cf8b54472625b9e14d893a169e2b0dedeb4df225',
    'eng': '7d4322bd2a7749724879683fc3912cb542f19906c83bcc1a52132556427170b2',
    'osd': '9cf5d576fcc47564f11265841e5ca839001e7e6f38ff7f7aacf46d15a96b00ff',
}
# WinGet APPINSTALLER_CLI_ERROR_NO_APPLICABLE_INSTALLER (winget-cli doc/.../returnCodes.md).
NO_APPLICABLE_INSTALLER = 0x8A150010


def selected(profile):
    if profile not in PROFILES:
        raise ValueError('Unknown profile')
    return {'base', 'documents', 'ocr', 'audio'} if profile == 'all' else {'base', profile}


def missing_tools(capabilities, env):
    names = []
    if 'documents' in capabilities and not system_tools.word_registered():
        names.append('soffice')
    if 'ocr' in capabilities:
        names.append('tesseract')
    if 'audio' in capabilities:
        names.append('ffmpeg')
    return [name for name in names if not system_tools.probe(name, env)]


def languages_available(env):
    tess = system_tools.find('tesseract', env)
    if not tess:
        return False
    result = subprocess.run([tess, '--list-langs'], env=env, capture_output=True,
                            text=True, encoding='utf-8', errors='replace', timeout=30)
    return result.returncode == 0 and set(LANGUAGES).issubset(set(result.stdout.splitlines()))


def install_package(name):
    winget = shutil.which('winget')
    if not winget:
        raise RuntimeError('WinGet missing. Install/update App Installer: https://aka.ms/getwinget ; then retry.')
    package = system_tools.PACKAGES[name]
    # Keep license/source prompts visible; never bypass installer hash validation.
    command = [winget, 'install', '--id', package, '--exact', '--source', 'winget']
    # Work PCs often deny admin rights: try a per-user install first. Only when WinGet
    # reports that no installer fits user scope (machine-wide NSIS/MSI) retry without it;
    # any other failure (declined agreement/UAC, network, hash) stops setup as before.
    result = subprocess.run(command + ['--scope', 'user'])
    if result.returncode == 0:
        return
    code = result.returncode & 0xFFFFFFFF
    if code != NO_APPLICABLE_INSTALLER:
        raise subprocess.CalledProcessError(result.returncode, result.args)
    print(f'{package} has no per-user installer (--scope user, WinGet code 0x{code:08X}). '
          f'Retrying the default install: Windows may ask for administrator rights.', flush=True)
    subprocess.run(command, check=True)


def download_language(language, folder, opener=urllib.request.urlopen):
    expected = LANGUAGES[language]
    target = folder/f'{language}.traineddata'
    if target.is_file() and hashlib.sha256(target.read_bytes()).hexdigest() == expected:
        return
    folder.mkdir(parents=True, exist_ok=True)
    url = f'https://raw.githubusercontent.com/tesseract-ocr/tessdata_fast/{REVISION}/{language}.traineddata'
    staging = None
    try:
        digest = hashlib.sha256()
        total = 0
        with tempfile.NamedTemporaryFile(dir=folder, suffix='.download', delete=False) as output:
            staging = Path(output.name)
            with opener(url, timeout=60) as response:
                while chunk := response.read(1024 * 1024):
                    total += len(chunk)
                    if total > 32 * 1024 * 1024:
                        raise RuntimeError('Language download exceeds the expected size limit')
                    digest.update(chunk)
                    output.write(chunk)
        if digest.hexdigest() != expected:
            raise RuntimeError(f'Checksum mismatch for {language}; download rejected')
        staging.replace(target)
    finally:
        if staging and staging.exists():
            staging.unlink()


def check(capabilities, config, root=ROOT):
    env = system_tools.environment(config, root)
    issues = [f'Missing or unusable program: {system_tools.PACKAGES[n]}' for n in missing_tools(capabilities, env)]
    for profile in sorted(capabilities & set(bootstrap.PROFILES)):
        try:
            bootstrap.ensure(profile, offline=True, root=root)
        except (OSError, RuntimeError, ValueError, subprocess.CalledProcessError) as exc:
            issues.append(f'Python profile {profile}: {exc}')
    if 'ocr' in capabilities and not languages_available(env):
        issues.append('OCR language data missing: rus, eng, osd')
    return issues


def prepare(profile, config, root=ROOT, confirm=input):
    capabilities = selected(profile)
    env = system_tools.environment(config, root)
    missing = missing_tools(capabilities, env)
    print('Python profiles: ' + ', '.join(sorted(capabilities & set(bootstrap.PROFILES))))
    print('System packages to install: ' + (', '.join(system_tools.PACKAGES[n] for n in missing) or 'none'))
    if 'ocr' in capabilities:
        print('OCR prepares rus/eng/osd data when needed (about 19 MB).')
    if 'audio' in capabilities:
        print('Audio packages may require several GB; speech models are downloaded separately on first use.')
    if confirm('Prepare the selected profile? [y/N] ').strip().lower() != 'y':
        print('Cancelled. No packages or models installed.')
        return 2
    if missing and not shutil.which('winget'):
        raise RuntimeError('Install/update App Installer first: https://aka.ms/getwinget')
    for name in missing:
        install_package(name)
        env = system_tools.environment(config, root)
        if not system_tools.probe(name, env):
            raise RuntimeError(f'{name} still unavailable. Restart the terminal or set tools.extra_path; then retry.')
    for item in sorted(capabilities & set(bootstrap.PROFILES)):
        bootstrap.ensure(item, root=root)
    if 'ocr' in capabilities and not languages_available(env):
        # Preserve explicit language locations. Do not modify existing system/user data.
        if config['tools']['tessdata_prefix'] or os.environ.get('TESSDATA_PREFIX'):
            raise RuntimeError('Explicit TESSDATA_PREFIX lacks rus/eng/osd. Fix it or clear it to use project data.')
        for language in LANGUAGES:
            print(f'Preparing Tesseract language: {language}', flush=True)
            download_language(language, root/'runtime'/'tessdata')
    issues = check(capabilities, config, root)
    if issues:
        raise RuntimeError('\n'.join(issues))
    subprocess.run([str(bootstrap.python_path(root)), str(root/'connect_skills.py')], check=True)
    report = root/'runtime'/'setup-report.json'
    report.parent.mkdir(parents=True, exist_ok=True)
    report.write_text(json.dumps({'profile': profile, 'checks': 'passed',
                                 'checked_at_utc': datetime.now(timezone.utc).isoformat(),
                                 'models_tested': False, 'word_com_execution_tested': False}, indent=2), encoding='utf-8')
    print('Selected dependencies and project skill links prepared.')
    print('This does not validate speech models, Word automation or final document layout.')
    return 0


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--profile', choices=PROFILES, default='base')
    parser.add_argument('--check', action='store_true')
    args = parser.parse_args()
    try:
        if os.name != 'nt':
            raise RuntimeError('This installer supports Windows only. Use bootstrap.py on other platforms.')
        config = local_config.load()
        if args.check:
            issues = check(selected(args.profile), config)
            print('\n'.join(issues) if issues else 'Selected dependencies available; models and document layout not tested.')
            return 1 if issues else 0
        return prepare(args.profile, config)
    except (OSError, ValueError, RuntimeError, subprocess.SubprocessError, EOFError) as exc:
        print(f'Setup incomplete: {exc}', file=sys.stderr)
        return 1


if __name__ == '__main__':
    raise SystemExit(main())
