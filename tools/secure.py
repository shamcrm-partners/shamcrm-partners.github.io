#!/usr/bin/env python3
"""
Закрытая часть сайта партнёров shamCRM: обучение, экзамен, шаблон сертификата.

Всё это лежит в secure.json в зашифрованном виде (AES-256-GCM, ключ из кода доступа
через PBKDF2-SHA256). Без кода прочитать содержимое нельзя — ни на сайте, ни на GitHub.

Команды:
  Зашифровать из исходника:
    python3 tools/secure.py encrypt content.json secure.json КОД
  Расшифровать (например, чтобы поправить темы или добавить видео):
    python3 tools/secure.py decrypt secure.json content.json КОД
  Сменить код (старый перестанет работать):
    python3 tools/secure.py rekey secure.json СТАРЫЙ_КОД НОВЫЙ_КОД

Нужен пакет cryptography:  pip install cryptography
Код можно вводить с дефисами и в любом регистре: SHAM-7K4Q-9XPD = sham7k4q9xpd.
"""
import base64, json, os, re, sys
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from cryptography.hazmat.primitives import hashes

ITER = 250_000


def norm(code):
    return re.sub(r'[^0-9A-Za-zА-Яа-яЁё]', '', code).upper()


def key_from(code, salt, iters=ITER):
    kdf = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=iters)
    return kdf.derive(norm(code).encode('utf-8'))


def encrypt(obj, code):
    salt, iv = os.urandom(16), os.urandom(12)
    data = AESGCM(key_from(code, salt)).encrypt(iv, json.dumps(obj, ensure_ascii=False).encode('utf-8'), None)
    b = lambda x: base64.b64encode(x).decode()
    return {'v': 1, 'kdf': 'PBKDF2-SHA256', 'iter': ITER, 'salt': b(salt), 'iv': b(iv), 'data': b(data)}


def decrypt(box, code):
    d = lambda k: base64.b64decode(box[k])
    plain = AESGCM(key_from(code, d('salt'), box['iter'])).decrypt(d('iv'), d('data'), None)
    return json.loads(plain.decode('utf-8'))


if __name__ == '__main__':
    cmd = sys.argv[1] if len(sys.argv) > 1 else ''
    if cmd == 'encrypt':
        src, dst, code = sys.argv[2:5]
        box = encrypt(json.load(open(src, encoding='utf-8')), code)
        json.dump(box, open(dst, 'w'), separators=(',', ':'))
        print('Зашифровано:', dst)
    elif cmd == 'decrypt':
        src, dst, code = sys.argv[2:5]
        json.dump(decrypt(json.load(open(src)), code), open(dst, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
        print('Расшифровано:', dst)
    elif cmd == 'rekey':
        path, old, new = sys.argv[2:5]
        box = encrypt(decrypt(json.load(open(path)), old), new)
        json.dump(box, open(path, 'w'), separators=(',', ':'))
        print('Код сменён:', path)
    else:
        print(__doc__)
