import asyncio
import logging
from fastapi import FastAPI, HTTPException, Response
from whatsapp_driver import WhatsAppManager
from constants import ACCOUNTS_DIR

app = FastAPI()
logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO)

@app.get("/open")
async def api_open(account: str = "default"):
    """
    Эндпоинт для входа в аккаунт. Если вход не выполнен, возвращается QR‑код.
    """
    try:
        session = WhatsAppManager.get_session(account)
        result = await asyncio.to_thread(session.open_browser_and_login)
        # Если возвращён QR‑код, отправляем его как изображение
        if "qr_code" in result:
            return Response(content=result["qr_code"], media_type="image/png")
        return result
    except Exception as e:
        logger.exception("Ошибка в /open: %s", e)
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/accounts")
async def api_accounts():
    """
    Возвращает список аккаунтов (папок) в директории ACCOUNTS_DIR.
    """
    try:
        import os
        accounts = [name for name in os.listdir(ACCOUNTS_DIR)
                    if os.path.isdir(os.path.join(ACCOUNTS_DIR, name))]
        return {"accounts": accounts}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/chats")
async def api_get_chats(account: str = "default"):
    """
    Возвращает список чатов для выбранного аккаунта.
    """
    try:
        session = WhatsAppManager.get_session(account)
        driver = session.get_driver(create_if_missing=False)
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.support import expected_conditions as EC
        from selenium.webdriver.common.by import By
        wait = WebDriverWait(driver, 60)
        container = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@aria-label='Список чатов']")))
        chat_items = container.find_elements(By.XPATH, ".//div[@role='listitem']")
        chats_list = []
        for chat in chat_items:
            try:
                title = chat.find_element(By.XPATH, ".//span[@dir='auto' and @title]").get_attribute("title")
                if title:
                    chats_list.append({"chat": title})
            except Exception:
                continue
        return {"chats": chats_list}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Ошибка при получении чатов: {e}")

@app.post("/send_message")
async def api_send_message(chat: str, message: str, account: str = "default"):
    """
    Отправляет текстовое сообщение в указанный чат.
    """
    try:
        session = WhatsAppManager.get_session(account)
        result = await asyncio.to_thread(session.send_message, chat, message)
        return result
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/send_file")
async def api_send_file(chat: str, file_path: str, account: str = "default"):
    """
    Отправляет файл в указанный чат.
    """
    try:
        session = WhatsAppManager.get_session(account)
        result = await asyncio.to_thread(session.send_file, chat, file_path)
        return result
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/close")
async def api_close(account: str = "default"):
    """
    Закрывает браузер для выбранного аккаунта.
    """
    try:
        result = await asyncio.to_thread(WhatsAppManager.close_session, account)
        return result
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/get_new_messages_unread")
async def api_get_new_messages_unread(account: str = "default"):
    """
    Возвращает новые непрочитанные сообщения для выбранного аккаунта.
    """
    try:
        session = WhatsAppManager.get_session(account)
        new_msgs = await asyncio.to_thread(session.get_new_messages_unread)
        return {"new_messages": new_msgs}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Ошибка при получении сообщений: {e}")
