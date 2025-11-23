"""Монитор новых заявок RT с дополнительными режимами оповещения."""

import sys
import time
import imaplib
import threading
from queue import SimpleQueue, Empty

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    StaleElementReferenceException,
)

try:
    import winsound
except ImportError:  # pragma: no cover - зависит от ОС
    winsound = None

from handler_common import handle_common_ticket
from handler_vvod_v_oborot import is_vvod_v_oborot, handle_vvod_v_oborot
from settings import (
    NEW_TICKET_FOCUS_SECONDS,
    NEW_MAIL_FOCUS_LINK,
    OPEN_ONLY_MODE,
    SOUND_ALERT_ON_NEW_TICKET,
    TEST_TICKET_LINK,
)

# ===== НАСТРОЙКИ =====
RT_URL = "https://rt.original-group.ru/"
USERNAME = "m.siluyanov"
PASSWORD = "Mukunda2004!"
POLL_INTERVAL = 5  # интервал проверки (сек)

IMAP_HOST = "owa.original-group.ru"
IMAP_PORT = 993
IMAP_USERNAME = "m.siluyanov"
IMAP_PASSWORD = "Mukunda2004!"
IMAP_MAILBOX = "INBOX"
IMAP_IDLE_TIMEOUT = 60


def play_sound_alert():
    """Простой сигнал при появлении новой заявки."""

    if not SOUND_ALERT_ON_NEW_TICKET:
        return

    if winsound is not None:
        try:
            winsound.MessageBeep(winsound.MB_ICONEXCLAMATION)
            return
        except RuntimeError:
            try:
                winsound.MessageBeep()
                return
            except RuntimeError:
                pass

    sys.stdout.write("\a")
    sys.stdout.flush()


def focus_new_ticket_tab(driver):
    """Пробует вывести вкладку RT на передний план."""

    try:
        driver.execute_script("window.focus();")
    except Exception:
        pass

    if NEW_TICKET_FOCUS_SECONDS > 0:
        time.sleep(NEW_TICKET_FOCUS_SECONDS)


def open_ticket_in_new_tab(driver, ticket_link: str) -> str:
    """Открывает ссылку тикета во вкладке, озвучивает событие и возвращает handle."""

    driver.execute_script("window.open(arguments[0]);", ticket_link)
    new_window_handle = driver.window_handles[-1]
    driver.switch_to.window(new_window_handle)
    play_sound_alert()
    focus_new_ticket_tab(driver)
    return new_window_handle


def simulate_new_ticket(driver, link: str = None):
    """Тестовая имитация новой заявки: открытие заданной ссылки + звук."""

    link_to_open = link or TEST_TICKET_LINK
    if not link_to_open:
        print(
            "TEST_TICKET_LINK не задан: укажи ссылку в settings.py или передай её аргументом."
        )
        return None

    print(f"Тестовая заявка: открываем {link_to_open}")
    handle = open_ticket_in_new_tab(driver, link_to_open)
    return handle


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


class ImapIdleWatcher:
    """Подписка на новые входящие письма через IMAP IDLE."""

    def __init__(
        self,
        host: str,
        port: int,
        username: str,
        password: str,
        mailbox: str = "INBOX",
        idle_timeout: int = 60,
        on_new_message=None,
    ) -> None:
        self.host = host
        self.port = port
        self.username = username
        self.password = password
        self.mailbox = mailbox
        self.idle_timeout = idle_timeout
        self.on_new_message = on_new_message
        self._stop_event = threading.Event()
        self._thread = None

    def start(self) -> None:
        if self._thread and self._thread.is_alive():
            return
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()

    def stop(self) -> None:
        self._stop_event.set()

    def _get_message_count(self, imap: imaplib.IMAP4_SSL) -> int:
        try:
            typ, data = imap.select(self.mailbox, readonly=True)
            if typ == "OK" and data and data[0]:
                return int(data[0])
        except Exception:
            return 0
        return 0

    def _parse_idle_responses(self, responses, last_count: int) -> int:
        new_count = last_count
        for response in responses or []:
            if not response:
                continue

            if isinstance(response, tuple) and len(response) >= 2:
                counter, marker = response[0], response[1]
                marker_bytes = marker if isinstance(marker, bytes) else str(marker).encode()
                if marker_bytes.upper() == b"EXISTS":
                    try:
                        exists_count = int(counter)
                    except (TypeError, ValueError):
                        continue
                    if exists_count > new_count:
                        new_count = exists_count

        return new_count

    def _notify(self) -> None:
        if callable(self.on_new_message):
            try:
                self.on_new_message()
            except Exception:
                pass

    def _run(self) -> None:
        while not self._stop_event.is_set():
            try:
                with imaplib.IMAP4_SSL(self.host, self.port) as imap:
                    imap.login(self.username, self.password)
                    last_exists = self._get_message_count(imap)

                    while not self._stop_event.is_set():
                        try:
                            if hasattr(imap, "idle"):
                                typ, _ = imap.idle()
                                if typ != "OK":
                                    break
                                try:
                                    responses = imap.idle_check(timeout=self.idle_timeout)
                                    updated_count = self._parse_idle_responses(
                                        responses, last_exists
                                    )
                                    if updated_count > last_exists:
                                        self._notify()
                                        last_exists = updated_count
                                finally:
                                    imap.idle_done()
                            else:
                                time.sleep(self.idle_timeout)
                                updated_count = self._get_message_count(imap)
                                if updated_count > last_exists:
                                    self._notify()
                                    last_exists = updated_count
                        except imaplib.IMAP4.abort:
                            break
                        except Exception:
                            time.sleep(1)
            except Exception:
                time.sleep(5)


def handle_new_mail_event(driver, main_window_handle: str):
    print("Получено новое письмо или ответ: воспроизводим оповещение")
    if NEW_MAIL_FOCUS_LINK:
        try:
            open_ticket_in_new_tab(driver, NEW_MAIL_FOCUS_LINK)
        finally:
            driver.switch_to.window(main_window_handle)
    else:
        play_sound_alert()
        focus_new_ticket_tab(driver)


def main():
    driver = webdriver.Chrome()

    try:
        login(driver)

        main_window_handle = driver.current_window_handle
        processed_ids = set()  # заявки, которые уже обрабатывали в этом запуске

        mail_events: SimpleQueue = SimpleQueue()
        imap_watcher = ImapIdleWatcher(
            host=IMAP_HOST,
            port=IMAP_PORT,
            username=IMAP_USERNAME,
            password=IMAP_PASSWORD,
            mailbox=IMAP_MAILBOX,
            idle_timeout=IMAP_IDLE_TIMEOUT,
            on_new_message=lambda: mail_events.put(None),
        )
        imap_watcher.start()

        while True:
            try:
                while True:
                    try:
                        mail_events.get_nowait()
                    except Empty:
                        break
                    else:
                        handle_new_mail_event(driver, main_window_handle)

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

                    if OPEN_ONLY_MODE:
                        print("  → Открываем без автоматических действий (режим наблюдения)")
                        try:
                            open_ticket_in_new_tab(driver, ticket["link"])
                            print("    вкладка остаётся открытой для ручной работы")
                        finally:
                            # Возвращаемся на основное окно, чтобы продолжать мониторить
                            driver.switch_to.window(main_window_handle)
                    else:
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
