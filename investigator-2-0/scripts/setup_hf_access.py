"""Настройка доступа к модели разделения говорящих в Диспетчере учётных данных Windows."""
import argparse
import getpass
import sys
import re

TARGET = 'Investigator2/HuggingFace'


def normalize_token(value):
    value = value.strip()
    parts = re.findall(r'hf_[A-Za-z0-9]+?(?=hf_|$)', value)
    if len(parts) > 1 and ''.join(parts) == value and len(set(parts)) == 1:
        print('Обнаружена повторная вставка одного токена; оставлена одна копия.')
        return parts[0]
    return value


def saved_token():
    import win32cred
    try:
        blob = win32cred.CredRead(TARGET, win32cred.CRED_TYPE_GENERIC)['CredentialBlob']
    except Exception:
        return None
    return blob.decode('utf-16-le') if isinstance(blob, bytes) else blob


def verify_access(token):
    """Check authentication and model access without printing credentials or profile."""
    from huggingface_hub import HfApi, hf_hub_download
    from pathlib import Path
    try:
        HfApi().whoami(token=token)
    except Exception as exc:
        status = getattr(getattr(exc, 'response', None), 'status_code', None)
        if status in (401, 403):
            print('Hugging Face отклонил токен. Создайте новый Read-токен и скопируйте полное значение hf_...')
        else:
            print('Не удалось проверить токен: проверьте соединение с Hugging Face и повторите запуск.')
        return False
    try:
        cache = Path(__file__).resolve().parents[2] / 'runtime' / 'models' / 'hub'
        hf_hub_download('pyannote/speaker-diarization-community-1', 'config.yaml',
                        token=token, cache_dir=str(cache))
    except Exception as exc:
        status = getattr(getattr(exc, 'response', None), 'status_code', None)
        if status in (401, 403):
            print('Токен действителен, но модель недоступна. Примите её условия под той же учётной записью и проверьте право чтения токена.')
        else:
            print('Токен действителен, но загрузка конфигурации модели не удалась. Повторите при доступном соединении.')
        return False
    print('Hugging Face принял токен; доступ к модели подтверждён.')
    return True


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--check', action='store_true', help='Проверить наличие без вывода токена')
    parser.add_argument('--verify', action='store_true', help='Проверить токен и доступ к модели через интернет')
    args = parser.parse_args()
    if args.verify:
        token = saved_token()
        if not token:
            print('Доступ пока не настроен')
            return 1
        return 0 if verify_access(token) else 3
    if args.check:
        print('Доступ сохранён' if saved_token() else 'Доступ пока не настроен')
        return 0 if saved_token() else 1
    print('Сначала примите условия модели pyannote/speaker-diarization-community-1 на Hugging Face.')
    print('Вставьте токен с правом чтения. Символы при вводе не отображаются. Нажмите Enter.')
    token = normalize_token(getpass.getpass('Hugging Face token: '))
    if not token.startswith('hf_') or len(token) < 10:
        print('Токен не сохранён: неверный формат.')
        return 2
    print('Проверяю токен и доступ к модели через интернет...')
    if not verify_access(token):
        print('Новый токен не сохранён. Исправьте указанную причину и повторите запуск.')
        return 3
    import win32cred
    win32cred.CredWrite({'Type':win32cred.CRED_TYPE_GENERIC,'TargetName':TARGET,
                        'UserName':'HuggingFace','CredentialBlob':token,
                        'Persist':win32cred.CRED_PERSIST_LOCAL_MACHINE},0)
    print('Токен сохранён в Диспетчере учётных данных текущего пользователя Windows.')
    print('Теперь сообщите в чате: доступ настроен. Сам токен в чат не отправляйте.')
    return 0


if __name__ == '__main__':
    sys.exit(main())
