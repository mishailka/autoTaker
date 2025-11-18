# rt_watcher.py
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    StaleElementReferenceException,
)

from handler_common import handle_common_ticket
from handler_vvod_v_oborot import is_vvod_v_oborot, handle_vvod_v_oborot

# ===== НАСТРОЙКИ =====
RT_URL = "https://rt.original-group.ru/"
USERNAME = "m.siluyanov"
PASSWORD = "Mukunda2004!"
POLL_INTERVAL = 5  # интервал проверки (сек)


def login(driver):
    """
    Логин в RT по форме, которую ты присылал.
    """
    driver.get(RT_URL)
    wait = WebDriverWait(driver, 20)

    # Ждём форму логина
    wait.until(EC.presence_of_element_located((By.ID, "login")))

    user_input = driver.find_element(By.NAME, "user")
    pass_input = driver.find_element(By.NAME, "pass")

    user_input.clear()
    user_input.send_keys(USERNAME)
    pass_input.clear()
    pass_input.send_keys(PASSWORD)

    submit_button = driver.find_element(
        By.XPATH, "//form[@id='login']//input[@type='submit' and @value='Войти в систему']"
    )
    submit_button.click()

    # Ждём основной экран
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    print("Успешный логин")


def find_unassigned_block(driver):
    """
    Находит div.titlebox с заголовком '10 последних неназначенных заявок'.
    """
    wait = WebDriverWait(driver, 10)
    link = wait.until(
        EC.presence_of_element_located(
            (
                By.XPATH,
                "//div[@class='titlebox-title']//a[contains(., '10 последних неназначенных заявок')]",
            )
        )
    )
    titlebox = link.find_element(By.XPATH, "./ancestor::div[@class='titlebox']")
    return titlebox


def get_tickets_from_block(titlebox):
    """
    Собирает тикеты из блока '10 последних неназначенных заявок'.

    Возвращает список словарей:
      {
        'id': '141867',
        'link': 'https://rt.original-group.ru/Ticket/Display.html?id=141867',
        'status': 'новый' / 'в работе' / ...,
        'subject': 'тема заявки'
      }
    """
    tickets = []

    try:
        content_div = titlebox.find_element(By.CSS_SELECTOR, "div.titlebox-content")
    except NoSuchElementException:
        return tickets

    try:
        table = content_div.find_element(By.CSS_SELECTOR, "table.ticket-list")
    except NoSuchElementException:
        return tickets

    bodies = table.find_elements(By.CSS_SELECTOR, "tbody.list-item")
    for body in bodies:
        try:
            row = body.find_element(By.CSS_SELECTOR, "tr")
            cells = row.find_elements(By.CSS_SELECTOR, "td.collection-as-table")
            if len(cells) < 5:
                continue

            number_cell = cells[0]
            subject_cell = cells[1]
            status_cell = cells[4]

            link_el = number_cell.find_element(By.TAG_NAME, "a")
            ticket_link = link_el.get_attribute("href")
            ticket_id = link_el.text.strip()

            subject_text = subject_cell.text.strip()
            status_text = status_cell.text.strip()

            tickets.append(
                {
                    "id": ticket_id,
                    "link": ticket_link,
                    "status": status_text,
                    "subject": subject_text,
                }
            )
        except (NoSuchElementException, StaleElementReferenceException):
            continue

    return tickets


def main():
    driver = webdriver.Chrome()
    driver.maximize_window()

    try:
        login(driver)

        main_window_handle = driver.current_window_handle
        processed_ids = set()  # заявки, которые уже обрабатывали в этом запуске

        while True:
            try:
                driver.refresh()
                time.sleep(1)

                titlebox = find_unassigned_block(driver)
                tickets = get_tickets_from_block(titlebox)

                for ticket in tickets:
                    ticket_id = ticket["id"]
                    status = ticket["status"].strip().lower()
                    subject = ticket["subject"]

                    # Не трогаем уже "в работе"
                    if status == "в работе":
                        continue

                    # Обрабатываем только новые (чтобы не спамить)
                    if ticket_id in processed_ids:
                        continue

                    print(
                        f"\nНовая заявка #{ticket_id} | статус: '{ticket['status']}' | тема: '{subject}'"
                    )

                    # Выбор обработчика по типу
                    if is_vvod_v_oborot(subject):
                        handle_vvod_v_oborot(driver, ticket, main_window_handle)
                    else:
                        handle_common_ticket(driver, ticket, main_window_handle)

                    processed_ids.add(ticket_id)

                # если новых не появилось — молчим
            except Exception as e:
                print(f"Ошибка в основном цикле: {e}")

            time.sleep(POLL_INTERVAL)

    finally:
        # Вкладки не закрываем, чтобы можно было смотреть обработанные заявки.
        # Если нужно закрывать браузер при завершении скрипта – раскомментируй:
        # driver.quit()
        pass


if __name__ == "__main__":
    main()
