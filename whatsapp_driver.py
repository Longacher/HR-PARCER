import os
import pickle
import time
import re
import threading
import logging
from datetime import datetime, timedelta

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import WebDriverException, TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys


from constants import ACCOUNTS_DIR, MAX_BROWSER_WAIT, DRIVER_WAIT_TIMEOUT, DOWNLOAD_TIMEOUT, SELECTOR_CONSTANTS
from utils import get_unique_filename, wait_for_new_file

logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)



class WhatsAppSession:
    def __init__(self, account: str):
        self.account = account
        self.profile_path = os.path.join(ACCOUNTS_DIR, account)
        os.makedirs(self.profile_path, exist_ok=True)
        self.download_dir = os.path.join(self.profile_path, "downloads")
        os.makedirs(self.download_dir, exist_ok=True)
        self.cookies_file = os.path.join(self.profile_path, "cookies.pkl")
        self.driver = None
        self.lock = threading.RLock()

    def create_driver(self) -> webdriver.Chrome:
        options = Options()
        options.add_argument(f"user-data-dir={self.profile_path}")
        options.add_argument("--disable-blink-features=AutomationControlled")
        # Настройка загрузок
        prefs = {
            "download.default_directory": self.download_dir,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        }
        options.add_experimental_option("prefs", prefs)
        # Для headless-режима можно раскомментировать:
        # options.add_argument("--headless")
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
        driver.maximize_window()
        return driver

    def get_driver(self, create_if_missing: bool = False) -> webdriver.Chrome:
        with self.lock:
            if self.driver:
                try:
                    current_url = self.driver.current_url
                except WebDriverException as e:
                    logger.warning("Driver не доступен для аккаунта %s: %s", self.account, str(e))
                    if create_if_missing:
                        self.driver = None
                    else:
                        raise Exception("Браузер закрыт. Используйте /open для запуска браузера.")
                if "web.whatsapp.com" not in current_url:
                    raise Exception("Браузер не находится на странице WhatsApp Web. Используйте /open для открытия страницы.")
                return self.driver
            else:
                if create_if_missing:
                    self.driver = self.create_driver()
                    return self.driver
                else:
                    raise Exception("Браузер не запущен. Используйте /open для его запуска.")

    def open_browser_and_login(self) -> dict:
        driver = self.get_driver(create_if_missing=True)
        driver.get("https://web.whatsapp.com/")
        start_time = time.time()

        while True:
            # Проверка успешного входа
            chat_elements = driver.find_elements(By.CSS_SELECTOR, "div[aria-label='Список чатов']")
            if chat_elements:
                with open(self.cookies_file, "wb") as f:
                    pickle.dump(driver.get_cookies(), f)
                return {"message": f"Аккаунт '{self.account}': вход выполнен"}

            # Получение QR-кода с canvas
            qr_canvas = driver.find_elements(By.CSS_SELECTOR, "canvas[aria-label='Scan this QR code to link a device!']")
            if qr_canvas:
                try:
                    # Делаем скриншот именно canvas элемента
                    qr_png = qr_canvas[0].screenshot_as_png
                    
                    # Дополнительная проверка что скриншот не пустой
                    if len(qr_png) > 1000:  # Минимальный размер валидного PNG
                        return {"qr_code": qr_png}
                    else:
                        logger.warning("Получен слишком маленький QR-код (%d байт)", len(qr_png))
                except Exception as e:
                    logger.error("Ошибка при создании скриншота QR: %s", str(e))

            if time.time() - start_time > MAX_BROWSER_WAIT:
                raise Exception(f"Timeout waiting for QR code or login")
            time.sleep(3)

    def create_new_chat(self, phone_number: str) -> dict:
        """
        Создает новый чат по номеру телефона
        Формат номера: 79123456789 (без + и других символов)
        """
        driver = self.get_driver(create_if_missing=False)
        wait = WebDriverWait(driver, DRIVER_WAIT_TIMEOUT)
        
        try:
            # Открываем страницу нового чата
            new_chat_button = wait.until(
                EC.element_to_be_clickable((By.XPATH, SELECTOR_CONSTANTS["new_chat_button_xpath"]))
            )
            new_chat_button.click()
            time.sleep(1)
            
            # Вводим номер телефона
            search_input = wait.until(
                EC.presence_of_element_located((By.XPATH, SELECTOR_CONSTANTS["search_input_xpath"]))
            )
            search_input.send_keys(phone_number)
            time.sleep(2)  # Ждем появления результатов
            
            # Выбираем контакт из результатов
            contact = wait.until(
                EC.element_to_be_clickable((By.XPATH, f"//span[@title='{phone_number}']/ancestor::div[@role='button']"))
            )
            contact.click()
            time.sleep(1)
            
            return {"status": "success", "message": f"Чат с {phone_number} создан"}
            
        except Exception as e:
            logger.exception(f"Ошибка при создании чата: {e}")
            return {"status": "error", "message": f"Не удалось создать чат: {str(e)}"}

    def send_message(self, phone_number: str, message: str) -> dict:
        """
        Отправляет сообщение на номер телефона, создавая новый чат при необходимости
        Номер должен быть в формате 79123456789 (без +)
        """
        driver = self.get_driver(create_if_missing=False)
        wait = WebDriverWait(driver, DRIVER_WAIT_TIMEOUT)
        
        try:
            # Сначала создаем новый чат
            self._create_new_chat(phone_number)
            
            # Отправляем сообщение
            input_field = wait.until(
                EC.visibility_of_element_located(
                    (By.XPATH, SELECTOR_CONSTANTS["message_input_xpath"])
                )
            )
            input_field.click()
            input_field.send_keys(message)
            
            send_button = wait.until(
                EC.element_to_be_clickable(
                    (By.XPATH, SELECTOR_CONSTANTS["send_button_xpath"])
                )
            )
            time.sleep(1)
            send_button.click()
            time.sleep(1)
            
            return {"status": "success", "message": "Сообщение отправлено"}
            
        except Exception as e:
            logger.exception(f"Ошибка при отправке сообщения: {e}")
            return {"status": "error", "message": f"Ошибка при отправке сообщения: {str(e)}"}
        finally:
            try:
                self.close_chat()
            except Exception as e:
                logger.warning(f"Ошибка при закрытии чата: {e}")

    def _open_existing_chat(self, phone_number: str):
        """Пытается открыть существующий чат по номеру телефона"""
        driver = self.get_driver()
        wait = WebDriverWait(driver, 10)
        
        # Ищем в списке чатов
        search_input = wait.until(
            EC.presence_of_element_located((By.XPATH, "//div[@contenteditable='true' and @data-tab='3']"))
        )
        search_input.clear()
        search_input.send_keys(phone_number)
        time.sleep(2)
        
        # Проверяем, появился ли чат
        chat_items = driver.find_elements(By.XPATH, "//div[@role='listitem']")
        for item in chat_items:
            try:
                title = item.find_element(By.XPATH, ".//span[@dir='auto' and @title]").get_attribute("title")
                if phone_number in title:
                    item.click()
                    return
            except:
                continue
        
        raise Exception("Чат не найден")

    def _create_new_chat(self, phone_number: str):
        """Создает новый чат по номеру телефона"""
        driver = self.get_driver()
        wait = WebDriverWait(driver, DRIVER_WAIT_TIMEOUT)
        
        try:
            # 1. Клик по кнопке нового чата
            new_chat_btn = wait.until(
                EC.element_to_be_clickable((By.XPATH, SELECTOR_CONSTANTS["new_chat_button_xpath"])))
            new_chat_btn.click()
            logger.info("Clicked new chat button")
            time.sleep(2)  # Важно дать время для открытия интерфейса нового чата

            # 2. Находим поле ввода для номера
            search_input = wait.until(
                EC.presence_of_element_located((By.XPATH, SELECTOR_CONSTANTS["search_input_xpath"])))
            
            # 3. Очищаем поле (если нужно) и вводим номер
            search_input.click()  # Активируем поле ввода
            actions = ActionChains(driver)
            actions.send_keys(phone_number)
            actions.perform()
            logger.info(f"Entered phone number: {phone_number}")
            time.sleep(2)  # Ждем появления результатов поиска
            actions.send_keys(Keys.RETURN)
            actions.perform()

            # 4. Клик по найденному контакту
            # contact = wait.until(
            #     EC.element_to_be_clickable((By.XPATH, 
            #         f"//span[contains(@title, '{phone_number}')]/ancestor::div[@role='button']")))
            # contact.click()
            # logger.info("Clicked on contact")
            # time.sleep(1)  # Ждем открытия чата

        except Exception as e:
            logger.error(f"Ошибка при создании чата: {e}")
            # Делаем скриншот для отладки
            driver.save_screenshot("create_chat_error.png")
            raise Exception(f"Не удалось создать чат: {str(e)}")

    def send_file(self, chat_name: str, file_path: str) -> dict:
        """
        Отправляет файл в указанный чат.
        """
        driver = self.get_driver(create_if_missing=False)
        wait = WebDriverWait(driver, DRIVER_WAIT_TIMEOUT)
        try:
            container = driver.find_element(By.XPATH, SELECTOR_CONSTANTS["chat_list_xpath"])
            chat_items = container.find_elements(By.XPATH, SELECTOR_CONSTANTS["chat_item_xpath"])
            target_chat = None
            for item in chat_items:
                try:
                    title = item.find_element(By.XPATH, SELECTOR_CONSTANTS["chat_title_xpath"]).get_attribute("title")
                    if title and title.lower() == chat_name.lower():
                        target_chat = item
                        break
                except NoSuchElementException:
                    continue
            if not target_chat:
                raise Exception("Чат не найден")
            target_chat.click()
            time.sleep(2)
            attach_button = wait.until(EC.element_to_be_clickable((By.XPATH, SELECTOR_CONSTANTS["attach_button_xpath"])))
            attach_button.click()
            time.sleep(1)
            file_input = driver.find_element(By.XPATH, SELECTOR_CONSTANTS["file_input_xpath"])
            abs_file_path = os.path.abspath(file_path)
            file_input.send_keys(abs_file_path)
            time.sleep(2)
            send_button = wait.until(EC.element_to_be_clickable((By.XPATH, SELECTOR_CONSTANTS["file_send_button_xpath"])))
            send_button.click()
            time.sleep(2)
            return {"message": "Файл успешно отправлен"}
        except Exception as e:
            logger.exception("Ошибка при отправке файла: %s", e)
            raise Exception(f"Ошибка при отправке файла: {e}")
        finally:
            try:
                self.close_chat()
            except Exception as e:
                logger.warning("Ошибка закрытия чата: %s", e)


    def get_new_messages_unread(self) -> dict:
        """
        Возвращает новые сообщения из чатов с индикаторами непрочитанных сообщений.
        Гарантируется, что ни в каком шаге не будет бесконечного ожидания.
        """
        import os
        from datetime import datetime, timedelta
        import re
        from selenium.webdriver.common.action_chains import ActionChains
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.support import expected_conditions as EC
        from selenium.common.exceptions import TimeoutException

        driver = self.get_driver(create_if_missing=False)
        wait = WebDriverWait(driver, DRIVER_WAIT_TIMEOUT)
        new_messages = {}
        download_dir = self.download_dir
        os.makedirs(download_dir, exist_ok=True)

        # Глобальный таймаут на выполнение всего метода (например, 3 минуты)
        global_timeout = 180  # секунд
        method_start_time = time.time()

        try:
            chat_list_container = wait.until(
                EC.presence_of_element_located((By.XPATH, SELECTOR_CONSTANTS["chat_list_xpath"]))
            )
        except Exception as e:
            # Если не удаётся получить список чатов, завершаем выполнение метода
            return new_messages

        chat_items = chat_list_container.find_elements(By.XPATH, SELECTOR_CONSTANTS["chat_item_xpath"])
        for chat_item in chat_items:
            # Проверка глобального таймаута
            if time.time() - method_start_time > global_timeout:
                break

            try:
                unread_badges = chat_item.find_elements(By.XPATH, SELECTOR_CONSTANTS["unread_badge_xpath"])
                if not unread_badges:
                    continue
            except Exception:
                continue

            try:
                chat_title_elem = chat_item.find_element(By.XPATH, SELECTOR_CONSTANTS["chat_title_xpath"])
                chat_title = chat_title_elem.get_attribute("title") or "Unknown Chat"
            except Exception:
                chat_title = "Unknown Chat"

            if chat_title not in new_messages:
                new_messages[chat_title] = []

            # Попытка открыть чат (ограничим число повторов, если не удаётся)
            open_attempts = 0
            max_open_attempts = 3
            while open_attempts < max_open_attempts:
                try:
                    chat_item.click()
                    time.sleep(1)
                    break
                except Exception:
                    open_attempts += 1
                    time.sleep(0.5)
            else:
                # Если чат так и не открылся, переходим к следующему
                continue

            # Поиск якоря "непрочит" с ограниченным числом попыток
            anchor = None
            anchor_attempts = 0
            max_anchor_attempts = 3
            while anchor_attempts < max_anchor_attempts and anchor is None:
                try:
                    anchor = wait.until(
                        EC.presence_of_element_located((By.XPATH, "//span[contains(text(), 'непрочит')]"))
                    )
                except Exception:
                    anchor_attempts += 1
                    time.sleep(0.5)
            if anchor:
                try:
                    message_rows = anchor.find_elements(By.XPATH, "following::div[contains(@class, 'message-in')]")
                except Exception:
                    message_rows = []
            else:
                # Если якоря нет, пытаемся получить все сообщения
                try:
                    message_rows = wait.until(
                        EC.presence_of_all_elements_located((By.XPATH, "//div[contains(@class, 'message-in')]"))
                    )
                except Exception:
                    message_rows = []

            # Если не найдено ни одного сообщения, закрываем чат и переходим дальше
            if not message_rows:
                try:
                    self.close_chat()
                except Exception:
                    pass
                continue

            for row in message_rows:
                # Проверка глобального таймаута
                if time.time() - method_start_time > global_timeout:
                    break

                # Определяем отправителя
                sender = "Unknown"
                try:
                    meta_elem = row.find_element(By.XPATH, ".//div[@data-pre-plain-text]")
                    meta = meta_elem.get_attribute("data-pre-plain-text")
                    meta_parts = meta.split("]")
                    if len(meta_parts) > 1:
                        meta_text = meta_parts[1].strip().rstrip(":")
                        if "," in meta_text:
                            sender_candidate = meta_text.split(",", 1)[1].strip()
                        else:
                            sender_candidate = meta_text
                        if sender_candidate:
                            sender = sender_candidate
                except Exception:
                    try:
                        sender_elem = row.find_element(By.XPATH, ".//span[@aria-label]")
                        sender_val = sender_elem.get_attribute("aria-label")
                        if sender_val:
                            sender = sender_val.rstrip(":").strip()
                    except Exception:
                        sender = "Unknown"

                now = datetime.now() + timedelta(seconds=1)
                time_str = now.strftime("%H:%M:%S")
                date_str = now.strftime("%Y-%m-%d")

                # 1. Обработка голосового сообщения
                audio_buttons = row.find_elements(By.XPATH, ".//button[@aria-label='Воспроизвести голосовое сообщение']")
                if audio_buttons:
                    try:
                        existing_files = set(os.listdir(download_dir))
                        audio_button = audio_buttons[0]
                        actions = ActionChains(driver)
                        actions.move_to_element(audio_button).perform()
                        time.sleep(0.5)
                        context_menu_button = WebDriverWait(driver, 10).until(
                            EC.visibility_of_element_located(
                                (By.XPATH, "//div[@data-js-context-icon='true' and @aria-label='Контекстное меню']")
                            )
                        )
                        context_menu_button.click()
                        time.sleep(0.5)
                        download_option = WebDriverWait(driver, 10).until(
                            EC.visibility_of_element_located((By.XPATH, "//div[@aria-label='Скачать']"))
                        )
                        download_option.click()

                        new_file = wait_for_new_file(download_dir, existing_files, timeout_sec=DOWNLOAD_TIMEOUT)
                        if not new_file:
                            raise Exception("Download timed out for audio message")
                        unique_name = get_unique_filename("audio", "audio.ogg", ".ogg")
                        os.rename(os.path.join(download_dir, new_file),
                                os.path.join(download_dir, unique_name))
                        msg = {
                            "type": "audio",
                            "sender": sender,
                            "time": time_str,
                            "date": date_str,
                            "file_name": unique_name,
                            "file_path": os.path.join(download_dir, unique_name)
                        }
                        new_messages[chat_title].append(msg)
                    except Exception as audio_err:
                        new_messages[chat_title].append({
                            "type": "audio",
                            "sender": sender,
                            "time": time_str,
                            "date": date_str,
                            "message": f"[Ошибка скачивания голосового сообщения: {audio_err}]"
                        })
                    continue

                # 2. Обработка файлового сообщения
                file_download_buttons = row.find_elements(By.XPATH, ".//div[@role='button'][contains(@title, 'Скачать')]")
                if file_download_buttons:
                    try:
                        download_button = file_download_buttons[0]
                        title_attr = download_button.get_attribute("title")
                        match = re.search(r'Скачать\s+"(.+)"', title_attr)
                        original_file_name = match.group(1) if match else "unknown_file"

                        try:
                            file_type_elem = row.find_element(By.XPATH, ".//span[@data-meta-key='type']")
                            file_type = file_type_elem.get_attribute("title") or file_type_elem.text
                        except Exception:
                            file_type = "unknown"
                        try:
                            size_elem = row.find_element(By.XPATH, ".//span[contains(text(), 'КБ') or contains(text(), 'МБ') or contains(text(), 'ГБ')]")
                            file_size = size_elem.text
                        except Exception:
                            file_size = "unknown"

                        download_button.click()
                        file_path = os.path.join(download_dir, original_file_name)
                        timeout_limit = time.time() + DOWNLOAD_TIMEOUT
                        while time.time() < timeout_limit:
                            if os.path.exists(file_path):
                                if not any(fname.startswith(original_file_name) and fname.endswith(".crdownload")
                                        for fname in os.listdir(download_dir)):
                                    break
                            time.sleep(0.5)
                        if not os.path.exists(file_path):
                            raise Exception("Download timed out for file: " + original_file_name)
                        unique_name = get_unique_filename("file", original_file_name, os.path.splitext(original_file_name)[1] or "")
                        os.rename(file_path, os.path.join(download_dir, unique_name))
                        file_path = os.path.join(download_dir, unique_name)
                        new_messages[chat_title].append({
                            "type": "file",
                            "sender": sender,
                            "time": time_str,
                            "date": date_str,
                            "file_name": unique_name,
                            "file_type": file_type,
                            "file_size": file_size,
                            "file_path": file_path
                        })
                    except Exception as e:
                        new_messages[chat_title].append({
                            "type": "file",
                            "sender": sender,
                            "time": time_str,
                            "date": date_str,
                            "message": f"[Ошибка скачивания файла: {e}]"
                        })
                    continue

                # 3. Обработка сообщения с изображением
                image_elements = row.find_elements(By.XPATH, ".//img[contains(@src, 'blob:')]")
                if image_elements:
                    try:
                        existing_files = set(os.listdir(download_dir))
                        img_element = image_elements[0]
                        actions = ActionChains(driver)
                        actions.move_to_element(img_element).perform()
                        time.sleep(0.5)
                        context_menu_button = WebDriverWait(driver, 10).until(
                            EC.visibility_of_element_located((By.XPATH, "//div[@data-js-context-icon='true' and @aria-label='Контекстное меню']"))
                        )
                        context_menu_button.click()
                        time.sleep(0.5)
                        download_option = WebDriverWait(driver, 10).until(
                            EC.visibility_of_element_located((By.XPATH, "//div[@aria-label='Скачать']"))
                        )
                        download_option.click()

                        new_file = wait_for_new_file(download_dir, existing_files, timeout_sec=DOWNLOAD_TIMEOUT)
                        if not new_file:
                            raise Exception("Download timed out for image message")
                        unique_name = get_unique_filename("image", "image.png", ".png")
                        os.rename(os.path.join(download_dir, new_file),
                                os.path.join(download_dir, unique_name))
                        file_path = os.path.join(download_dir, unique_name)
                        new_messages[chat_title].append({
                            "type": "image",
                            "sender": sender,
                            "time": time_str,
                            "date": date_str,
                            "file_name": unique_name,
                            "file_path": file_path
                        })
                    except Exception as e:
                        new_messages[chat_title].append({
                            "type": "image",
                            "sender": sender,
                            "time": time_str,
                            "date": date_str,
                            "message": f"[Ошибка скачивания изображения: {e}]"
                        })
                    continue

                # 4. Обработка текстового сообщения
                try:
                    text_elem = row.find_element(By.XPATH, ".//span[contains(@class, 'selectable-text')]")
                    text = text_elem.text.strip()
                except Exception:
                    text = ""
                if text:
                    new_messages[chat_title].append({
                        "type": "text",
                        "sender": sender,
                        "time": time_str,
                        "date": date_str,
                        "message": text
                    })

            try:
                self.close_chat()
            except Exception:
                pass

        return new_messages

    def close_driver(self) -> dict:
        """
        Закрывает браузер для текущей сессии.
        """
        with self.lock:
            if self.driver:
                try:
                    self.driver.quit()
                    self.driver = None
                    return {"message": f"Браузер для аккаунта '{self.account}' закрыт."}
                except Exception as e:
                    logger.exception("Ошибка при закрытии драйвера: %s", e)
                    raise Exception(f"Ошибка при закрытии драйвера: {e}")
            else:
                raise Exception(f"Драйвер для аккаунта '{self.account}' не найден.")

    def close_chat(self) -> None:
        """
        Закрывает текущий открытый чат.
        """
        driver = self.get_driver(create_if_missing=False)
        wait = WebDriverWait(driver, 10)
        try:
            menu_button = wait.until(
                EC.element_to_be_clickable((By.XPATH, SELECTOR_CONSTANTS["menu_button_xpath"]))
            )
            menu_button.click()
            time.sleep(0.1)
            close_option = wait.until(
                EC.element_to_be_clickable((By.XPATH, SELECTOR_CONSTANTS["close_chat_xpath"]))
            )
            close_option.click()
            time.sleep(0.1)
        except Exception as e:
            logger.exception("Ошибка при закрытии чата: %s", e)
            # Не поднимаем исключение, чтобы не прерывать основную логику

# Класс для управления сессиями по аккаунтам
class WhatsAppManager:
    _sessions = {}
    _lock = threading.RLock()

    @classmethod
    def get_session(cls, account: str) -> WhatsAppSession:
        with cls._lock:
            if account not in cls._sessions:
                cls._sessions[account] = WhatsAppSession(account)
            return cls._sessions[account]

    @classmethod
    def close_session(cls, account: str) -> dict:
        with cls._lock:
            if account in cls._sessions:
                result = cls._sessions[account].close_driver()
                del cls._sessions[account]
                return result
            else:
                raise Exception(f"Сессия для аккаунта '{account}' не найдена.")

