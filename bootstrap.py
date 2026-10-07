"""Create an isolated Python environment; install explicitly selected profiles."""
from pathlib import Path
import argparse
import hashlib
import json
import os
import re
import subprocess
import sys
import venv

ROOT = Path(__file__).resolve().parent
PROFILES = ('base', 'ocr', 'audio')


def python_path(root=ROOT):
    return root / '.venv' / ('Scripts/python.exe' if os.name == 'nt' else 'bin/python')


def fingerprint(profiles, root=ROOT):
    digest = hashlib.sha256()
    digest.update(f'{sys.version_info[:2]}:{sys.platform}'.encode())
    for profile in sorted(profiles):
        digest.update(profile.encode())
        digest.update((root/'requirements'/f'{profile}.txt').read_bytes())
    return digest.hexdigest()


def run_checked(command, **kwargs):
    subprocess.run(command, check=True, **kwargs)


def ensure(profile='base', offline=False, root=ROOT):
    """Install cumulatively; do not mark a failed installation as ready."""
    if profile not in PROFILES:
        raise ValueError('Unknown dependency profile')
    if sys.version_info[:2] != (3, 12):
        raise RuntimeError('Use Python 3.12 for this release candidate.')
    executable = python_path(root)
    stamp = root/'.venv'/'investigator-install.json'
    previous = {}
    if stamp.exists():
        try:
            previous = json.loads(stamp.read_text(encoding='utf-8'))
            if not isinstance(previous, dict) or not isinstance(previous.get('profiles', []), list):
                previous = {}
        except (OSError, ValueError):
            pass
    profiles = set(previous.get('profiles', [])) & set(PROFILES)
    profiles.update(('base', profile))
    expected = fingerprint(profiles, root)
    if executable.exists() and previous.get('fingerprint') == expected:
        # Check metadata rather than importing heavy audio libraries on every run.
        probe = 'from importlib.metadata import version\n'
        probe += '\n'.join(f'assert version({name!r}) == {version!r}' for name, version in required_versions(profiles, root).items())
        result = subprocess.run([str(executable), '-c', probe], capture_output=True)
        if result.returncode == 0:
            print('Environment is already prepared.', flush=True)
            return executable
    if offline:
        raise RuntimeError('Environment is not prepared. Run bootstrap.py with network access first.')
    if not executable.exists():
        print('Creating .venv ...', flush=True)
        venv.EnvBuilder(with_pip=True).create(root/'.venv')
    # A stale stamp must not survive a partially successful pip run.
    if stamp.exists():
        stamp.unlink()
    command = [str(executable), '-m', 'pip', '--disable-pip-version-check', 'install',
               '--index-url', 'https://pypi.org/simple']
    for item in sorted(profiles):
        command.extend(('-r', str(root/'requirements'/f'{item}.txt')))
    print('Installing profiles: ' + ', '.join(sorted(profiles)), flush=True)
    run_checked(command)
    run_checked([str(executable), '-m', 'pip', 'check'])
    run_checked([str(executable), '-c', 'import docx, pymupdf, pypdf, PIL, yaml'])
    freeze = subprocess.check_output([str(executable), '-m', 'pip', 'freeze'], text=True)
    (root/'.venv'/'installed-versions.txt').write_text(freeze, encoding='utf-8')
    stamp.write_text(json.dumps({'profiles': sorted(profiles), 'fingerprint': expected}, indent=2), encoding='utf-8')
    print('Python packages installed. System tools and models require separate checks.', flush=True)
    return executable


def required_versions(profiles, root=ROOT):
    versions = {}
    for profile in sorted(profiles):
        for raw in (root/'requirements'/f'{profile}.txt').read_text(encoding='utf-8').splitlines():
            line = raw.strip()
            if not line or line.startswith('#'):
                continue
            spec, _, marker = line.partition(';')
            if marker:
                if marker.strip() != 'sys_platform == "win32"':
                    raise ValueError('Unsupported dependency marker')
                if sys.platform != 'win32':
                    continue
            match = re.fullmatch(r'([A-Za-z0-9_.-]+)==([A-Za-z0-9.+_-]+)', spec.strip())
            if not match:
                raise ValueError('Only pinned dependency versions are supported')
            versions[match[1]] = match[2]
    return versions


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--profile', choices=PROFILES, default='base')
    parser.add_argument('--offline', action='store_true', help='Never download or install packages')
    args = parser.parse_args()
    try:
        ensure(args.profile, args.offline)
    except (OSError, ValueError, RuntimeError, subprocess.CalledProcessError) as exc:
        print(f'Installation incomplete: {exc}', file=sys.stderr)
        return 1
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
