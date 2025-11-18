# handler_vvod_v_oborot.py
from handler_common import process_ticket_link
from settings import AUTO_TAKE_ENABLED, AUTO_STATUS_UPDATE_ENABLED


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

    should_take = AUTO_TAKE_ENABLED  # спец-заявки берём при включённом авто-взятии
    should_set_status = AUTO_STATUS_UPDATE_ENABLED

    if not should_take and not should_set_status:
        print("  → Для 'Ввод в оборот' авто-действия отключены, пропускаем.")
        return

    process_ticket_link(
        driver,
        ticket_link,
        main_window_handle,
        should_take=should_take,
        should_set_status=should_set_status,
    )
