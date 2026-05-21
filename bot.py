import asyncio
import logging
import io
from aiogram import Bot, Dispatcher, F
from aiogram.types import Message
from aiogram.filters import CommandStart
from aiogram.client.session.aiohttp import AiohttpSession
from config import BOT_TOKEN, EXCEL_PATH
import os
import llm_service
import voice_handler
import excel_parser

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

MAX_MSG_LEN = 4096

async def send_long_message(message: Message, text: str):
    """Відправляє довге повідомлення частинами якщо перевищує ліміт Telegram"""
    if not text:
        return
        
    if len(text) <= MAX_MSG_LEN:
        await message.answer(text, disable_web_page_preview=True)
    else:
        # Split by paragraphs if possible to avoid cutting in the middle of a line
        paragraphs = text.split('\n\n')
        current_chunk = ""
        for p in paragraphs:
            if len(current_chunk) + len(p) + 2 <= MAX_MSG_LEN:
                current_chunk += p + '\n\n'
            else:
                if current_chunk:
                    await message.answer(current_chunk.strip(), disable_web_page_preview=True)
                
                # If a single paragraph is too long, split it by characters
                if len(p) > MAX_MSG_LEN:
                    for i in range(0, len(p), MAX_MSG_LEN):
                        await message.answer(p[i:i+MAX_MSG_LEN], disable_web_page_preview=True)
                    current_chunk = ""
                else:
                    current_chunk = p + '\n\n'
        
        if current_chunk:
            await message.answer(current_chunk.strip(), disable_web_page_preview=True)

# Збільшений timeout щоб Telegram не обривав з'єднання поки LLM думає
session = AiohttpSession(timeout=300)
bot = Bot(token=BOT_TOKEN, session=session)
dp = Dispatcher()

ADMIN_IDS = [340517348, 8482582995]

@dp.message(CommandStart())
async def start_cmd(message: Message):
    await message.answer("Привіт! Я бот для формування красивих підбірок турів. Надішли мені текст або голосове повідомлення з деталями туру і цінами, і я все красиво оформлю.")

@dp.message(F.document)
async def handle_document(message: Message):
    if message.from_user.id not in ADMIN_IDS:
        return
        
    doc = message.document
    if not doc.file_name.lower().endswith(('.xlsx', '.xls')):
        await message.answer("❌ Будь ласка, надішліть файл бази готелів у форматі .xlsx")
        return
        
    msg = await message.answer("⏳ Завантажую та перевіряю нову базу готелів...")
    try:
        file_info = await bot.get_file(doc.file_id)
        uploaded_path = EXCEL_PATH.replace("tours.xlsx", "tours_uploaded.xlsx")
        os.makedirs(os.path.dirname(uploaded_path), exist_ok=True)
        await bot.download_file(file_info.file_path, uploaded_path)
        
        # Invalidate cache
        excel_parser._db_cache["data"] = None
        
        # Test loading (this will also warm up the cache instantly!)
        db = excel_parser.get_hotel_db()
        if db:
            await msg.edit_text(f"✅ Базу успішно оновлено та проіндексовано! Знайдено {len(db)} напрямків.")
        else:
            await msg.edit_text("⚠️ Файл завантажено, але він здається порожнім або має невірний формат.")
            
    except Exception as e:
        logger.error(f"Error updating DB: {e}")
        await msg.edit_text(f"❌ Помилка під час оновлення файлу: {str(e)}")

@dp.message(F.text)
async def handle_text(message: Message):
    msg = await message.answer("✨ Формую підбірку...")
    try:
        result = await llm_service.format_tour_message(message.text)
    except Exception as e:
        logger.error(f"format_tour_message error: {e}")
        await msg.edit_text("❌ Внутрішня помилка під час генерації. Спробуй ще раз.")
        return
    try:
        if len(result) <= MAX_MSG_LEN:
            await msg.edit_text(result, disable_web_page_preview=True)
        else:
            await msg.delete()
            await send_long_message(message, result)
    except Exception as e:
        logger.error(f"edit_text error: {e}")
        await send_long_message(message, result)

@dp.message(F.voice)
async def handle_voice(message: Message):
    msg = await message.answer("🎙 Розпізнаю голосове...")
    file_id = message.voice.file_id
    file_info = await bot.get_file(file_id)
    
    buf = io.BytesIO()
    await bot.download_file(file_info.file_path, buf)
    file_bytes = buf.getvalue()
    
    # 1. Get RAW transcription first
    raw_text = await voice_handler.transcribe_voice(file_bytes)
    if not raw_text or raw_text.startswith("❌"):
        await msg.edit_text("🤷 Не вдалося розпізнати текст.")
        return
    
    # 2. Show the RAW text to the user (as requested: "показує текст який менеджер сказав")
    await msg.edit_text(f"🗣 «{raw_text}»")
    
    # 3. Show the progress message
    status_msg = await message.answer("✨ Формую підбірку...")
    
    try:
        # 4. Quick destination detection from raw text to get hotel list for cleanup
        db = excel_parser.get_hotel_db()
        dest_hotels = None
        if db:
            dest = llm_service._pick_destination_by_keywords(raw_text, list(db.keys()))
            if dest:
                dest_hotels = db.get(dest, [])
        
        # 5. Cleanup the text with hotel DB context for better correction
        cleaned_text = await voice_handler.cleanup_transcribed_text(raw_text, destination_hotels=dest_hotels)
        logger.info(f"Voice cleaned: {cleaned_text[:200]}...")
        
        # 6. Use the cleaned text for formatting
        result = await llm_service.format_tour_message(user_text=cleaned_text, raw_voice_text=raw_text)
        
        # 6. Send the final result
        if len(result) <= MAX_MSG_LEN:
            await status_msg.edit_text(result, disable_web_page_preview=True)
        else:
            await status_msg.delete()
            await send_long_message(message, result)
            
    except Exception as e:
        logger.error(f"format_tour_message voice error: {e}")
        await status_msg.edit_text("❌ Внутрішня помилка під час генерації. Спробуй ще раз.")

async def main():
    if not BOT_TOKEN:
        logger.error("No BOT_TOKEN in .env")
        return
    logger.info("Bot started!")
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
