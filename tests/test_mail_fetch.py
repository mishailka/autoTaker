"""Тестовый скрипт для вывода писем за последнюю неделю."""

import datetime
import email
import imaplib
import sys
from email.header import decode_header

from rt_watcher import IMAP_HOST, IMAP_PASSWORD, IMAP_PORT, IMAP_USERNAME


def _decode_header(value: str) -> str:
    decoded = []
    for part, enc in decode_header(value or ""):
        if isinstance(part, bytes):
            try:
                decoded.append(part.decode(enc or "utf-8", errors="replace"))
            except LookupError:
                decoded.append(part.decode("utf-8", errors="replace"))
        else:
            decoded.append(part)
    return "".join(decoded)


def _print_message_summary(msg):
    subject = _decode_header(msg.get("Subject", ""))
    sender = _decode_header(msg.get("From", ""))
    date = msg.get("Date", "")
    print(f"От: {sender}\nТема: {subject}\nДата: {date}\n{'-'*40}")


def main():
    since_date = (datetime.datetime.now() - datetime.timedelta(days=7)).strftime("%d-%b-%Y")
    try:
        with imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT) as imap:
            imap.login(IMAP_USERNAME, IMAP_PASSWORD)
            imap.select("INBOX")
            status, data = imap.search(None, f'(SINCE "{since_date}")')
            if status != "OK":
                print("Не удалось выполнить поиск писем", file=sys.stderr)
                return 1

            message_ids = data[0].split()
            print(f"Найдено писем за неделю: {len(message_ids)}\n{'='*40}")
            for msg_id in message_ids:
                status, msg_data = imap.fetch(msg_id, "(RFC822)")
                if status != "OK" or not msg_data:
                    continue
                try:
                    msg = email.message_from_bytes(msg_data[0][1])
                except Exception:
                    continue
                _print_message_summary(msg)
    except Exception as exc:  # pragma: no cover - диагностический скрипт
        print(f"Ошибка подключения или получения писем: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
