import asyncio
import httpx
import logging
import json
import re
import random
import itertools
from openai import AsyncOpenAI
from config import OPENROUTER_API_KEY, GROQ_API_KEY, GROQ_API_KEYS

logger = logging.getLogger(__name__)

# --- Voice Specific Settings ---
# Models for voice transcription cleanup
VOICE_CLEANUP_MODELS = ["openai/gpt-4o-mini", "google/gemini-2.5-flash"]

# Whisper prompt for transcription — the more context words, the better Whisper recognizes them
WHISPER_PROMPT = (
    "Авіатур на Майорку, Тенеріфе, Крит, Корфу, Родос, Кіпр, Ібіцу, Коста-Браву, Фуертевентуру, Лансароте, Гран-Канарію. "
    "Виліт з Берліна, Дюссельдорфа, Варшави, Києва. Двоє дорослих, троє дітей. "
    "Перший готель, другий готель, третій готель, четвертий готель, п'ятий готель, шостий готель, "
    "сьомий готель, восьмий готель, дев'ятий готель, десятий готель, одинадцятий готель, дванадцятий готель. "
    "BLUESEA, Globales, AzuLine, HSM, BJ Playamar, Iberostar, Rixos, Mitsis, Grecotel, H10, Riu, Barcelo, "
    "Sol, Melia, THB, Hipotels, Zafiro, Viva, Occidental, Allegro, Palladium, JS, Mar Hotels, "
    "Canvas, Kalypso, Knossos, Porto Platanias, Coriva Beach, Magda Hotel, Aquila, Daios Cove, Royal Hideaway, "
    "Out Of The Blue, Mythos Suites, Minos Palace, Domes, Nana, Elounda, Panoramica, GF Noelia, "
    "Villa Andromeda, Can Simoneta, Formentor. "
    "Сніданки, напівпансіон, повний пансіон, все включено, ультра все включено, без харчування. "
    "Ціна 100 євро, 200 євро, 500 євро, 1000 євро, 1500 євро, 2000 євро, 3000 євро за номер. "
    "Пріоріті, багаж 20 кілограм, індивідуальний трансфер, шатл-бас, екскурсійна програма."
)

# Prompt for cleaning up voice transcription
VOICE_CLEANUP_PROMPT = """Ти — коректор туристичних текстів із голосового розпізнавання. 
Твоє ЄДИНЕ завдання: виправити помилки та структурувати текст, НЕ ДОДАЮЧИ нічого від себе.

КРИТИЧНО ВАЖЛИВІ ПРАВИЛА:
1. ЗБЕРЕЖИ ВСЕ: Кожне слово, кожну цифру, кожну ціну, кожен готель. НЕ ВИДАЛЯЙ нічого!
2. СТРУКТУРА: Пронумеруй кожен готель окремим рядком:
   1 готель - [назва] зі [харчуванням] [ціна] євро за номер
   2 готель - [назва] зі [харчуванням] [ціна] євро за номер
3. КІЛЬКІСТЬ: Якщо ти чуєш N окремих назв готелів — має бути N рядків. Якщо згадано "перший", "другий"... "восьмий" — має бути 8 готелів.
4. РОЗДІЛЯЙ ГОТЕЛІ: Якщо кілька назв йдуть поспіль без нумерації — кожна назва це ОКРЕМИЙ готель. Наприклад: "Canvas 700 Porto 800 Domes 1000" = 3 окремих готелі.
5. ВИПРАВ транслітерацію: "Blau C", "BluSea", "блюсі/блю сі" → "BLUESEA", "Marzas" → "Marthas", "Ses Cades" → "Ses Cases", "S Bolero" → "Es Bolero", "глобаліс" → "Globales", "іберостар" → "Iberostar", "азулін" → "AzuLine", "ріксос" → "Rixos", "мітсіс" → "Mitsis", "грекотель" → "Grecotel", "акуалія/аквіла" → "Aquila", "ноелія" → "GF Noelia".
6. ХАРЧУВАННЯ: Не видаляй тип харчування (сніданки, все включено тощо).
7. НЕ ВИГАДУЙ назви готелів, яких не було в тексті!
8. Все крім готелів (дати, рейси, ціна авіа, послуги) — залиш одним абзацом на початку.
"""

# --- LLM Client for voice cleanup ---
client = AsyncOpenAI(
    api_key=OPENROUTER_API_KEY,
    base_url="https://openrouter.ai/api/v1",
)

# --- Key Rotation for Groq ---
def _create_key_rotator():
    keys = GROQ_API_KEYS.copy()
    if not keys and GROQ_API_KEY:
        keys = [GROQ_API_KEY]
    random.shuffle(keys)
    return itertools.cycle(keys)

_groq_key_rotator = _create_key_rotator()

# Gemini transcription prompt — gives Gemini context about what it's hearing
GEMINI_TRANSCRIBE_PROMPT = """Транскрибуй це голосове повідомлення ДОСЛІВНО. Це туристичний менеджер, який диктує деталі туру.

ПРАВИЛА:
1. Запиши ВСЕ, що сказано — кожне слово, кожну цифру, кожну назву.
2. Менеджер говорить українською/російською, але назви готелів — англійською. Збережи їх латиницею.
3. СТРУКТУРА ТЕКСТУ: Менеджер зазвичай спочатку називає напрямок, дати, рейси, ціну авіа, потім перераховує готелі по порядку (перший готель, другий готель... або 1, 2, 3...).
4. Кожен готель — окремий рядок з назвою, типом харчування та ціною.
5. НЕ ДОДАВАЙ нічого від себе, НЕ ВИПРАВЛЯЙ назви готелів, просто запиши що чуєш.
6. Якщо чутно нерозбірливо — запиши як чуєш, не пропускай.

СЛОВНИК ЧАСТИХ НАЗВ ТА БРЕНДІВ (використовуй для кращого розпізнавання на слух):
Майорка, Тенеріфе, Крит, Корфу, Родос, Кіпр, Ібіца, Коста-Брава, Фуертевентура, Лансароте, Гран-Канарія.
BLUESEA, Globales, AzuLine, HSM, BJ Playamar, Iberostar, Rixos, Mitsis, Grecotel, H10, Riu, Barcelo, Sol, Melia, THB, Hipotels, Zafiro, Viva, Occidental, Allegro, Palladium, JS, Mar Hotels, Jumeirah, Can Simoneta, Castell Son Claret, The Lodge.
"""

async def _transcribe_with_gemini(file_bytes: bytes) -> str:
    """Try transcribing voice using Gemini via OpenRouter (multimodal audio)."""
    import base64
    
    audio_b64 = base64.b64encode(file_bytes).decode('utf-8')
    
    try:
        resp = await client.chat.completions.create(
            model="google/gemini-2.5-flash",
            messages=[
                {"role": "system", "content": GEMINI_TRANSCRIBE_PROMPT},
                {"role": "user", "content": [
                    {
                        "type": "input_audio",
                        "input_audio": {
                            "data": audio_b64,
                            "format": "ogg"
                        }
                    }
                ]},
            ],
            temperature=0,
            timeout=30,
            max_tokens=3000,
        )
        text = resp.choices[0].message.content.strip()
        if text and len(text) > 10:
            logger.info(f"Gemini transcription success: {text[:100]}...")
            return text
    except Exception as e:
        logger.warning(f"Gemini transcription failed: {e}")
    return None

def apply_phonetic_fixes(text: str) -> str:
    if not text:
        return text
    fixes = {
        "Blau C": "BLUESEA", "blau c": "BLUESEA", "Blau c": "BLUESEA", "BlauC": "BLUESEA",
        "BluSea": "BLUESEA", "blusea": "BLUESEA", "Blusea": "BLUESEA",
        "Marzas": "Marthas", "marzas": "Marthas",
        "Ses Cades": "Ses Cases", "ses cades": "Ses Cases",
        " S Bolero": " Es Bolero", " s bolero": " Es Bolero",
        "блюсія": "BLUESEA", "блю сі": "BLUESEA", "Блюсія": "BLUESEA", "Блю сі": "BLUESEA",
        "глобаліс": "Globales", "Глобаліс": "Globales",
        "плеймар": "Playamar", "Плеймар": "Playamar",
        "азулін": "AzuLine", "Азулін": "AzuLine",
        "мітсіс": "Mitsis", "Мітсіс": "Mitsis",
        "ріксос": "Rixos", "Ріксос": "Rixos",
        "грекотель": "Grecotel", "Грекотель": "Grecotel",
        "акуаліа": "Aquila", "Акуаліа": "Aquila",
        "аквіла": "Aquila", "Аквіла": "Aquila",
    }
    for bad, good in fixes.items():
        text = text.replace(bad, good)
    return text

async def _transcribe_with_whisper(file_bytes: bytes) -> str:
    """Fallback: Transcribe using Groq Whisper."""
    active_keys = GROQ_API_KEYS if GROQ_API_KEYS else ([GROQ_API_KEY] if GROQ_API_KEY else [])
    if not active_keys:
        return None
    
    for _ in range(len(active_keys)):
        key = next(_groq_key_rotator)
        url_groq = "https://api.groq.com/openai/v1/audio/transcriptions"
        headers_groq = {"Authorization": f"Bearer {key}"}
        
        files = {"file": ("voice.ogg", file_bytes, "audio/ogg")}
        data = {
            "model": "whisper-large-v3-turbo",
            "prompt": WHISPER_PROMPT,
            "response_format": "json",
            "temperature": "0",
        }
        
        async with httpx.AsyncClient() as c:
            try:
                resp = await c.post(url_groq, headers=headers_groq, files=files, data=data, timeout=25)
                if resp.status_code == 200:
                    text = resp.json().get("text", "")
                    if text:
                        return apply_phonetic_fixes(text)
                logger.warning(f"Groq key {key[:10]}... returned status {resp.status_code}.")
            except Exception as e:
                logger.warning(f"Groq key {key[:10]}... failed: {e}")
                continue
    return None

async def transcribe_voice(file_bytes: bytes) -> str:
    """Transcribes voice using Gemini (primary) with Whisper fallback."""
    
    # 1. Try Gemini first — much better at understanding context and mixed languages
    gemini_text = await _transcribe_with_gemini(file_bytes)
    if gemini_text:
        return apply_phonetic_fixes(gemini_text)
    
    # 2. Fallback to Whisper on Groq
    logger.info("Gemini failed, falling back to Whisper...")
    whisper_text = await _transcribe_with_whisper(file_bytes)
    if whisper_text:
        return whisper_text
    
    # 3. Final fallback to OpenRouter Whisper
    if OPENROUTER_API_KEY:
        try:
            url_or = "https://openrouter.ai/api/v1/audio/transcriptions"
            headers_or = {"Authorization": f"Bearer {OPENROUTER_API_KEY}"}
            files = {"file": ("voice.ogg", file_bytes, "audio/ogg")}
            data_or = {
                "model": "openai/whisper-large-v3",
                "prompt": WHISPER_PROMPT
            }
            
            async with httpx.AsyncClient() as c:
                resp = await c.post(url_or, headers=headers_or, files=files, data=data_or, timeout=30)
                if resp.status_code == 200:
                    text = resp.json().get("text", "")
                    if text:
                        return apply_phonetic_fixes(text)
        except Exception as e:
            logger.error(f"OpenRouter Whisper fallback failed: {e}")
    
    return "❌ Помилка розпізнавання (всі сервіси недоступні)."

async def cleanup_transcribed_text(raw_text: str, destination_hotels: list = None) -> str:
    """Cleans up the transcription using LLM to fix errors and hallucinations."""
    if not raw_text:
        return raw_text
    
    logger.info(f"Voice transcription raw: {raw_text}")
    
    # Build context with hotel names if available
    user_content = raw_text
    if destination_hotels:
        hotel_names = "\n".join([h['hotel'] for h in destination_hotels])
        user_content = f"ТЕКСТ З ГОЛОСОВОГО:\n{raw_text}\n\nДОВІДНИК ГОТЕЛІВ (використай для виправлення назв):\n{hotel_names}"
    
    for model in VOICE_CLEANUP_MODELS:
        try:
            resp = await client.chat.completions.create(
                model=model,
                messages=[
                    {"role": "system", "content": VOICE_CLEANUP_PROMPT},
                    {"role": "user", "content": user_content},
                ],
                temperature=0,
                timeout=30,
                max_tokens=2000,
            )
            cleaned = resp.choices[0].message.content.strip()
            # Remove potential markdown code blocks
            cleaned = re.sub(r'```[a-z]*\n?', '', cleaned).strip('`').strip()
            
            if cleaned and len(cleaned) > 10:
                logger.info(f"Voice transcription after cleanup ({model}): {cleaned}")
                return cleaned
        except Exception as e:
            logger.error(f"Voice cleanup error with {model}: {e}")
            
    return raw_text

async def process_voice_message(file_bytes: bytes) -> str:
    """Full pipeline: transcription -> cleanup."""
    text = await transcribe_voice(file_bytes)
    if not text or text.startswith("❌"):
        return text
    
    cleaned_text = await cleanup_transcribed_text(text)
    return cleaned_text
