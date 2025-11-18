# handler_common.py
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

from settings import AUTO_TAKE_ENABLED, AUTO_STATUS_UPDATE_ENABLED

# Ключевые слова для "обычных" заявок, при которых надо сначала нажать "Взять"
KEYWORDS = [
    "ввод кодов",
    "что то ещё",
    # добавляй свои
]


def subject_matches_keywords(subject: str) -> bool:
    """
    Проверяет, содержит ли тема одно из ключевых слов (без учёта регистра).
    """
    s = subject.lower()
    for kw in KEYWORDS:
        if kw.lower() in s:
            return True
    return False


def process_ticket_link(
    driver,
    ticket_link: str,
    main_window_handle: str,
    should_take: bool,
    should_set_status: bool,
):
    """
    Низкоуровневая обработка тикета:
      - открыть в новой вкладке
      - (опционально) нажать 'Взять'
      - нажать 'В работе'
      - нажать 'Изменить заявку'
    Вкладка тикета остаётся открытой, возвращаемся на основную.
    """
    if not should_take and not should_set_status:
        print("  → Автоматические действия отключены, тикет пропускаем.")
        return

    driver.execute_script("window.open(arguments[0]);", ticket_link)
    new_window_handle = driver.window_handles[-1]
    driver.switch_to.window(new_window_handle)

    wait = WebDriverWait(driver, 15)

    try:
        # --- шаг 1: при необходимости "Взять" ---
        if should_take:
            # Меню "Действия" (если нужно раскрыть)
            try:
                actions_menu = wait.until(
                    EC.element_to_be_clickable((By.ID, "page-actions"))
                )
                actions_menu.click()
            except TimeoutException:
                pass

            try:
                take_action = wait.until(
                    EC.element_to_be_clickable((By.ID, "page-actions-take"))
                )
                button_text = take_action.text.strip().lower()
                if "взять" not in button_text:
                    print(
                        "  → Пропускаем 'Взять': кнопка уже не предлагает взять (заявка, возможно, взята вручную)."
                    )
                else:
                    take_action.click()
                    # Ждём, пока страница обновится и снова появится "Действия"
                    wait.until(EC.presence_of_element_located((By.ID, "page-actions")))
            except TimeoutException:
                print(f"Не удалось нажать 'Взять' для {ticket_link}")
        else:
            print("  → Авто-взятие отключено или не требуется.")

        # --- шаг 2: 'В работе' ---
        if should_set_status:
            try:
                actions_menu = wait.until(
                    EC.element_to_be_clickable((By.ID, "page-actions"))
                )
                actions_menu.click()
            except TimeoutException:
                pass

            work_action = wait.until(
                EC.element_to_be_clickable((By.ID, "page-actions-inprogress"))
            )
            work_action.click()

            # --- шаг 3: 'Изменить заявку' ---
            submit_button = wait.until(
                EC.element_to_be_clickable(
                    (
                        By.XPATH,
                        "//input[@type='submit' and @name='SubmitTicket' and @value='Изменить заявку']",
                    )
                )
            )
            submit_button.click()

            if should_take:
                print(
                    f"  → ВЗЯЛИ (где нужно) и перевели в 'В работе' + сохранили: {ticket_link}"
                )
            else:
                print(f"  → Перевели в 'В работе' + сохранили: {ticket_link}")
        else:
            print("  → Автоустановка статуса отключена.")

    except TimeoutException:
        print(f"Не удалось полностью обработать тикет: {ticket_link}")
    finally:
        # Возврат на основную вкладку, вкладку тикета не закрываем
        driver.switch_to.window(main_window_handle)


def handle_common_ticket(driver, ticket: dict, main_window_handle: str):
    """
    Общий обработчик для любых заявок, которые не требуют спец-логики.
    Сам решает, надо ли сначала 'Взять' (по KEYWORDS).
    """
    subject = ticket["subject"]
    ticket_link = ticket["link"]

    need_take = subject_matches_keywords(subject)
    if need_take:
        if AUTO_TAKE_ENABLED:
            print(
                "  → Тема содержит ключевое слово (общий обработчик): будем сначала 'Взять' тикет."
            )
        else:
            print(
                "  → Тема содержит ключевое слово, но авто-взятие отключено в настройках."
            )

    should_take = need_take and AUTO_TAKE_ENABLED
    should_set_status = AUTO_STATUS_UPDATE_ENABLED

    if not should_take and not should_set_status:
        print("  → Для этого тикета автоматические действия отключены, пропускаем.")
        return

    process_ticket_link(
        driver,
        ticket_link,
        main_window_handle,
        should_take=should_take,
        should_set_status=should_set_status,
    )
