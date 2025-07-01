import asyncio
import logging
from telegram import Bot
from django.conf import settings
from django.utils import timezone

logger = logging.getLogger(__name__)

async def send_telegram_message_async(message, chat_ids=None):
    """
    Асинхронно отправляет сообщение через Telegram-бота.
    
    Args:
        message (str): Текст сообщения
        chat_ids (list): Список ID чатов для отправки. Если None, используются ID из настроек.
    
    Returns:
        bool: True, если сообщение успешно отправлено хотя бы одному получателю, иначе False
    """
    if not settings.TELEGRAM_BOT_TOKEN:
        logger.error("TELEGRAM_BOT_TOKEN не настроен")
        return False
    
    if not chat_ids:
        chat_ids = settings.TELEGRAM_CHAT_IDS
        if not chat_ids or chat_ids == ['']:
            logger.error("Не указаны ID чатов для отправки сообщений")
            return False
    
    bot = Bot(token=settings.TELEGRAM_BOT_TOKEN)
    success = False
    
    for chat_id in chat_ids:
        if not chat_id:
            continue
        
        try:
            await bot.send_message(chat_id=chat_id, text=message, parse_mode='HTML')
            success = True
            logger.info(f"Сообщение отправлено в чат {chat_id}")
        except Exception as e:
            logger.error(f"Ошибка при отправке сообщения в чат {chat_id}: {str(e)}")
    
    return success

def send_telegram_message(message, chat_ids=None):
    """
    Синхронная обертка для отправки сообщения через Telegram-бота.
    
    Args:
        message (str): Текст сообщения
        chat_ids (list): Список ID чатов для отправки. Если None, используются ID из настроек.
    
    Returns:
        bool: True, если сообщение успешно отправлено хотя бы одному получателю, иначе False
    """
    try:
        return asyncio.run(send_telegram_message_async(message, chat_ids))
    except Exception as e:
        logger.error(f"Ошибка при отправке сообщения в Telegram: {str(e)}")
        return False 