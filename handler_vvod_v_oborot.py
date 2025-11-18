# handler_vvod_v_oborot.py
from handler_common import process_ticket_link


def is_vvod_v_oborot(subject: str) -> bool:
    """
    Определяем, что это заявка вида 'Ввод в оборот'.
    Можно усложнить логику, если нужно.
    """
    s = subject.lower()
    return "ввод в оборот" in s


def handle_vvod_v_oborot(driver, ticket: dict, main_window_handle: str):
    """
    Специализированный обработчик заявок 'Ввод в оборот'.

    Здесь в будущем можно:
      - распарсить параметры из темы/описания,
      - сходить в какие-то внешние системы,
      - провести проверки и т.д.

    Пока просто:
      - всегда 'Взять',
      - перевести 'В работе',
      - 'Изменить заявку'.
    """
    subject = ticket["subject"]
    ticket_link = ticket["link"]

    print("  → Спец-обработка 'Ввод в оборот'")
    print(f"    тема: {subject}")

    # Для таких заявок считаем, что всегда нужно 'Взять'
    process_ticket_link(driver, ticket_link, main_window_handle, need_take=True)
