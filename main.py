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

# ===== НАСТРОЙКИ =====
RT_URL = "https://rt.original-group.ru/"
USERNAME = "m.siluyanov"
PASSWORD = "Mukunda2004!"
POLL_INTERVAL = 5  # интервал проверки (сек)

# Ключевые слова для темы заявки (без учёта регистра)
KEYWORDS = [
    "ввод кодов",
    # добавляй свои слова сюда
]


def subject_matches_keywords(subject: str) -> bool:
    """
    Проверка: тема содержит какое-либо ключевое слово (без учёта регистра).
    """
    s = subject.lower()
    for kw in KEYWORDS:
        if kw.lower() in s:
            return True
    return False


def login(driver):
    """
    Логин в RT по форме, которую ты прислал.
    """
    driver.get(RT_URL)

    wait = WebDriverWait(driver, 20)

    # Ждём форму логина
    wait.until(EC.presence_of_element_located((By.ID, "login")))

    # Поле "Имя пользователя:" -> <input name="user" id="user">
    user_input = driver.find_element(By.NAME, "user")
    # Поле "Пароль:" -> <input type="password" name="pass">
    pass_input = driver.find_element(By.NAME, "pass")

    user_input.clear()
    user_input.send_keys(USERNAME)
    pass_input.clear()
    pass_input.send_keys(PASSWORD)

    # Кнопка "Войти в систему"
    submit_button = driver.find_element(
        By.XPATH, "//form[@id='login']//input[@type='submit' and @value='Войти в систему']"
    )
    submit_button.click()

    # Ждём загрузки основной страницы после логина
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    print("Успешный логин")


def find_unassigned_block(driver):
    """
    Находит <div class='titlebox'>, в заголовке которого ссылка
    '10 последних неназначенных заявок'.
    Возвращает элемент div.titlebox.
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

    Ожидаем структуру:
      <div class="titlebox-content">
        <table class="ticket-list">
          <tbody class="list-item" data-record-id="...">
            <tr>
              <td>Номер (ссылкой)</td>
              <td>Тема</td>
              <td>Приоритет</td>
              <td>Очередь</td>
              <td>Статус</td>
            </tr>
          </tbody>
        </table>

    Возвращает список словарей:
      {
        'id': '141867',
        'link': 'https://rt.original-group.ru/Ticket/Display.html?id=141867',
        'status': 'в работе' / 'новый' / ...,
        'subject': 'текст темы'
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


def set_ticket_in_work(driver, ticket_link, main_window_handle, need_take: bool):
    """
    Открывает тикет в НОВОЙ вкладке.

    Если need_take == True:
      - нажимает "Взять" (id="page-actions-take"),
    затем:
      - нажимает "В работе" (id="page-actions-inprogress"),
      - на форме изменения жмёт "Изменить заявку".

    Вкладку тикета не закрывает, возвращается на основную.
    """
    # Открываем новую вкладку c тикетом
    driver.execute_script("window.open(arguments[0]);", ticket_link)
    new_window_handle = driver.window_handles[-1]
    driver.switch_to.window(new_window_handle)

    wait = WebDriverWait(driver, 15)

    try:
        # --- шаг 1: при необходимости "Взять" ---
        if need_take:
            # раскрываем меню "Действия" (если есть)
            try:
                actions_menu = wait.until(
                    EC.element_to_be_clickable((By.ID, "page-actions"))
                )
                actions_menu.click()
            except TimeoutException:
                pass  # бывает, что меню уже открыто или выглядит иначе

            try:
                take_action = wait.until(
                    EC.element_to_be_clickable((By.ID, "page-actions-take"))
                )
                take_action.click()
                # ждём перезагрузку/обновление страницы после "Взять"
                wait.until(
                    EC.presence_of_element_located((By.ID, "page-actions"))
                )
            except TimeoutException:
                print(f"Не удалось 'Взять' тикет: {ticket_link}")

        # --- шаг 2: "В работе" ---
        # снова раскрываем меню "Действия" (после возможной перезагрузки)
        try:
            actions_menu = wait.until(
                EC.element_to_be_clickable((By.ID, "page-actions"))
            )
            actions_menu.click()
        except TimeoutException:
            pass

        # <a id="page-actions-inprogress" ...>В работе</a>
        work_action = wait.until(
            EC.element_to_be_clickable((By.ID, "page-actions-inprogress"))
        )
        work_action.click()

        # --- шаг 3: жмём "Изменить заявку" ---
        submit_button = wait.until(
            EC.element_to_be_clickable(
                (
                    By.XPATH,
                    "//input[@type='submit' and @name='SubmitTicket' and @value='Изменить заявку']",
                )
            )
        )
        submit_button.click()

        time.sleep(1)
        if need_take:
            print(f"Тикет {ticket_link} ВЗЯТ и переведён в 'В работе' + сохранён.")
        else:
            print(f"Тикет {ticket_link} переведён в 'В работе' + сохранён.")
    except TimeoutException:
        print(f"Не удалось полностью обработать тикет: {ticket_link}")
    finally:
        # Возвращаемся на главную вкладку, вкладку тикета не закрываем
        driver.switch_to.window(main_window_handle)


def main():
    driver = webdriver.Chrome()
    driver.maximize_window()

    try:
        login(driver)

        main_window_handle = driver.current_window_handle
        processed_ids = set()  # тикеты, которые уже обработали в этом запуске

        while True:
            try:
                driver.refresh()
                time.sleep(1)  # даём странице перерисоваться

                titlebox = find_unassigned_block(driver)
                tickets = get_tickets_from_block(titlebox)

                # Ищем только новые тикеты (которых ещё нет в processed_ids)
                for ticket in tickets:
                    ticket_id = ticket["id"]
                    ticket_link = ticket["link"]
                    status = ticket["status"].strip().lower()
                    subject = ticket["subject"]

                    # Берём только заявки НЕ в статусе "в работе"
                    if status == "в работе":
                        continue

                    # Если мы уже этот тикет обрабатывали — пропускаем без вывода
                    if ticket_id in processed_ids:
                        continue

                    # Сюда попадаем ТОЛЬКО когда видим новую заявку впервые
                    need_take = subject_matches_keywords(subject)

                    print(
                        f"Новая заявка #{ticket_id} | статус: '{ticket['status']}' | тема: '{subject}'"
                    )
                    if need_take:
                        print("  → Тема содержит ключевое слово, будем сначала 'Взять' тикет.")

                    set_ticket_in_work(
                        driver,
                        ticket_link,
                        main_window_handle,
                        need_take=need_take,
                    )
                    processed_ids.add(ticket_id)

                # если новых не появилось — просто молчим
            except Exception as e:
                print(f"Ошибка в основном цикле: {e}")

            time.sleep(POLL_INTERVAL)

    finally:
        # Чтобы вкладки не закрывались автоматически, driver.quit() не вызываем.
        # Если нужно авто-закрытие браузера при завершении скрипта — раскомментируй:
        # driver.quit()
        pass


if __name__ == "__main__":
    main()
