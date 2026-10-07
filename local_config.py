"""Read the optional local profile without printing personal values."""
from pathlib import Path
import argparse
from datetime import date
import json

ROOT = Path(__file__).resolve().parent


def load(root=ROOT):
    example = json.loads((root/'config.example.json').read_text(encoding='utf-8'))
    path = root/'config.local.json'
    if not path.exists():
        return example
    value = json.loads(path.read_text(encoding='utf-8-sig'))
    validate(value, example, root)
    return value


def validate(value, example, root=ROOT):
    def walk(actual, template, prefix=''):
        if not isinstance(actual, dict) or set(actual) != set(template):
            raise ValueError(f'{prefix or "profile"}: keys must match config.example.json')
        for key, default in template.items():
            field = prefix + key
            item = actual[key]
            if isinstance(default, dict):
                walk(item, default, field + '.')
            elif key == 'schema_version':
                if type(item) is not int or item != 1:
                    raise ValueError('Unsupported schema_version')
            elif key == 'requested_reply_days':
                if item is not None and (type(item) is not int or item < 1):
                    raise ValueError(f'{field}: expected null or positive integer')
            elif isinstance(default, list):
                if not isinstance(item, list) or any(not isinstance(x, str) for x in item):
                    raise ValueError(f'{field}: expected list of strings')
            elif not isinstance(item, str):
                raise ValueError(f'{field}: expected string')
    walk(value, example)
    if value['signer']['valid_from']:
        try:
            date.fromisoformat(value['signer']['valid_from'])
        except ValueError:
            raise ValueError('signer.valid_from: expected ISO date YYYY-MM-DD') from None
    folder = value['projects_root']
    if folder:
        path = Path(folder)
        if not path.is_absolute():
            raise ValueError('projects_root must be absolute')
        if path.resolve().is_relative_to(root.resolve()):
            raise ValueError('projects_root must be outside the public repository')
    paths = [value['templates_root'], value['correspondence']['letterhead_file'],
             value['tools']['tessdata_prefix'], *value['tools']['extra_path']]
    if any(item and not Path(item).is_absolute() for item in paths):
        raise ValueError('Configured paths must be absolute or empty')


def legacy_constants(config, root=ROOT):
    """Map explicit profile values to the existing generator schema, without inference."""
    data = json.loads((root/'investigator-2-0'/'константы.json').read_text(encoding='utf-8'))
    signer = config['signer']
    correspondence = config['correspondence']
    person = {}
    for source, target in [('full_name', 'фио'), ('short_name', 'инициалы_фам'),
                           ('position', 'должность'), ('rank', 'звание'), ('valid_from', 'действует_с')]:
        if signer[source]:
            person[target] = signer[source]
    if person:
        data['следователь'] = person
    if signer['department']:
        data['орган_полное'] = signer['department']
    for source, target in [('reply_postal_address', 'адрес'), ('reply_email', 'email'),
                           ('phone', 'исполнитель_телефон')]:
        if correspondence[source]:
            data[target] = correspondence[source]
    data['профиль_запросов'] = {
        'срок_по_умолчанию_суток': correspondence['requested_reply_days'],
        'email_ответа': correspondence['reply_email'] or None,
    }
    return data


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--check', action='store_true')
    parser.parse_args()
    try:
        value = load()
        print('Profile structure: OK')
        print('Local file: ' + ('present' if (ROOT/'config.local.json').exists() else 'absent; empty defaults'))
        print('Projects folder: ' + ('specified' if value['projects_root'] else 'not specified'))
        print('Profile values do not prove authority or facts of a case.')
    except (ValueError, OSError):
        print('Invalid local profile. Compare fields and types with config.example.json.')
        return 1
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
