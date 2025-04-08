import os

# Базовые директории
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
ACCOUNTS_DIR = os.path.join(CURRENT_DIR, "accounts")
os.makedirs(ACCOUNTS_DIR, exist_ok=True)
DB_FILE = os.path.join(CURRENT_DIR, "messages.db")

# Таймауты и интервалы ожидания (в секундах)
MAX_BROWSER_WAIT = 60         # Общее время ожидания входа в аккаунт
DRIVER_WAIT_TIMEOUT = 60      # Таймаут для WebDriverWait (используется в других операциях)
DOWNLOAD_TIMEOUT = 20         # Таймаут для загрузки файлов

# Селекторы (XPath, CSS, классы) для элементов WhatsApp Web
SELECTOR_CONSTANTS = {
    # Для авторизации и QR‑кода
    "chat_list_class": "x1iyjqo2",  # (может меняться, корректируйте при необходимости)
    "qr_code_xpath": '/html/body/div[1]/div/div/div[2]/div[2]/div[1]/div/div/div[2]/div[2]/div[1]/canvas',

    # Для списка чатов и элементов чата
    "chat_list_xpath": "//div[@aria-label='Список чатов']",
    "chat_item_xpath": ".//div[@role='listitem']",
    "chat_title_xpath": ".//span[@dir='auto' and @title]",

    # Для ввода и отправки текстового сообщения
    "message_input_xpath": "//div[@contenteditable='true' and @aria-label='Введите сообщение' and @data-tab='10']",
    "send_button_xpath": "//div[contains(@class, 'x123j3cw')]//button[@data-tab='11' and @aria-label='Отправить']",

    # Для отправки файлов
    "attach_button_xpath": "//span[@data-icon='plus']",
    "file_input_xpath": "//input[@accept='*']",
    "file_send_button_xpath": "//span[@data-icon='send']",

    # Для поиска непрочитанных сообщений
    "unread_badge_xpath": ".//span[contains(@aria-label, 'непрочит')]",
    "unread_anchor_xpath": "//span[contains(text(), 'непрочит')]",

    # Для обработки входящих сообщений
    "message_in_xpath": "//div[contains(@class, 'message-in')]",
    "meta_xpath": ".//div[@data-pre-plain-text]",
    "sender_xpath": ".//span[@aria-label]",
    "audio_button_xpath": ".//button[@aria-label='Воспроизвести голосовое сообщение']",
    "file_download_button_xpath": ".//div[@role='button'][contains(@title, 'Скачать')]",
    "file_type_xpath": ".//span[@data-meta-key='type']",
    "file_size_xpath": ".//span[contains(text(), 'КБ') or contains(text(), 'МБ') or contains(text(), 'ГБ')]",
    "image_xpath": ".//img[contains(@src, 'blob:')]",
    "text_xpath": ".//span[contains(@class, 'selectable-text')]",

    # Для работы с контекстным меню и скачиванием
    "context_menu_xpath": "//div[@data-js-context-icon='true' and @aria-label='Контекстное меню']",
    "download_option_xpath": "//div[@aria-label='Скачать']",

    # Для закрытия чата
    "menu_button_xpath": "/html/body/div[1]/div/div/div[3]/div/div[4]/div/header/div[3]/div/div[3]/div/button/span",
    "close_chat_xpath": "//div[@aria-label='Закрыть чат']",

    "new_chat_button_xpath": "//button[@aria-label='Новый чат']",
    "search_input_xpath": "//p[@class='selectable-text copyable-text x15bjb6t x1n2onr6']",
    "contact_item_xpath": "//div[@role='listitem']",
}
