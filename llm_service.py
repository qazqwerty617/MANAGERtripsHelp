import asyncio
import logging
import json
import re
import difflib
from openai import AsyncOpenAI
from config import OPENROUTER_API_KEY, GROQ_API_KEY, GROQ_API_KEYS
from excel_parser import get_hotel_db, get_tourist_tax_db, get_tax_per_person_per_night, get_tax_info
from voice_handler import cleanup_transcribed_text

logger = logging.getLogger(__name__)

def _safe_int(val, default=0):
    try:
        if val is None: return default
        s = re.sub(r'[^\d]', '', str(val))
        return int(s) if s else default
    except: return default

client = AsyncOpenAI(
    api_key=OPENROUTER_API_KEY,
    base_url="https://openrouter.ai/api/v1",
)

_NOISE_TOKENS = {
    "hotel", "hotels", "apartments", "apartment", "apartamentos", "apartamento",
    "resort", "spa", "villas", "villa", "the", "by", "and", "suites", "suite",
    "hostal", "pension", "hostel",
}

_DESTINATION_ALIASES = {
    "майорк": "майорка",
    "mallorca": "майорка",
    "палма": "майорка",
    "коста-брава": "коста-брава",
    "costa brava": "коста-брава",
    "коста-дель-соль": "коста-дель-соль",
    "тенериф": "тенеріфе",
    "tenerife": "тенеріфе",
    "гран-канар": "гран-карарія",
    "фуертевентур": "фуертевентура",
    "fuerteventura": "фуертевентура",
    "лансарот": "лансароте",
    "lanzarote": "лансароте",
    "ібіц": "ібіца",
    "ibiza": "ібіца",
    "крит": "крит",
    "crete": "крит",
    "корфу": "корфу",
    "corfu": "корфу",
    "родос": "родос",
    "rhodes": "родос",
    "кіпр": "кіпр",
    "cyprus": "кіпр",
    "мальор": "майорка",
    "мальорк": "майорка",
}

# Brands for strict matching (Global for all functions)
BRANDS = {"bluesea", "hipotels", "globales", "iberostar", "rixos", "mitsis", "grecotel", "sol", "melia", "hsm", "azuline", "bj", "bg", "thb", "bahia", "palladium", "h10", "riu", "barcelo", "occidental", "allegro", "viva", "zafiro", "js", "bjs", "mar"}

_DESTINATION_PROMPT = """Ти — туристичний асистент. Тобі надіслали текст-чернетку від менеджера з описом туру.
Твоє завдання: визначити напрямок (країну/острів/регіон) з тексту і вибрати один найбільш підходящий варіант із наданого списку доступних напрямків.
Поверни ТІЛЬКИ назву напрямку зі списку. Якщо жоден не підходить, поверни "Unknown".
Без жодного іншого тексту.
"""

_EXTRACT_PROMPT = """Ти — спеціалізований AI-асистент для вилучення назв готелів. 
Твоє завдання: знайти у тексті менеджера ВСІ згадані готелі і зіставити їх з наданим списком з бази.

ПРАВИЛА:
1. ПОРЯДОК ТА КІЛЬКІСТЬ: Повертай готелі СУВОРО в тому порядку, в якому вони йдуть у тексті. Це КРИТИЧНО для голосових повідомлень.
2. КІЛЬКІСТЬ: Якщо менеджер назвав 8 готелів — ПОВЕРНИ 8. Якщо назвав 5 — поверни 5. НІКОЛИ не пропускай жодного готелю!
3. Твоя головна мета — знайти ТОЧНУ відповідність у "СПИСКУ ГОТЕЛІВ НАПРЯМКУ". Назви повинні збігатися на 95%+.
4. СУВОРО: ОЧИЩАЙ назви готелів від цін, типів харчування (напр. 'все включено', 'сніданки') та іншого сміття. У назві має залишитись ТІЛЬКИ назва готелю.
5. СУВОРО ЗАБОРОНЕНО ВИГАДУВАТИ! Якщо готелю з тексту НЕМАЄ в наданому списку (немає точного або 95%+ збігу) — НЕ ПІДСТАВЛЯЙ ІНШИЙ ГОТЕЛЬ! У такому випадку поверни оригінальну ОЧИЩЕНУ назву з тексту менеджера, додавши префікс [NOT_FOUND].
6. Якщо вказано 8 готелів — поверни 8. Не намагайся додати зайві готелі з бази, яких немає в тексті.
7. ФОРМАТ: Тільки JSON {"hotels": ["Name 1", "Name 2", "[NOT_FOUND] Name 3"]}. Жодного іншого тексту.
"""


_EXTRACT_PRICES_PROMPT = """Ти — фінансовий аналітик туристичних турів. 
Твоє завдання: витягти числові дані для розрахунку.

ПРАВИЛА:
1. adults: кількість дорослих.
2. children: кількість дітей.
3. infants: кількість немовлят.
4. nights: кількість ночей.
5. check_in_month: номер місяця (1-12).
6. check_in_day: число місяця.
7. flight_price: ЗАГАЛЬНА ціна авіа за всіх (число).
8. flight_is_per_person: true, якщо менеджер чітко сказав "за особу/на людину", інакше false.
9. hotel_prices: СЛОВНИК, де ключ - це назва готелю, а значення - загальна ціна за номер (тільки число). КРИТИЧНО: Витягни ціни для ВСІХ готелів у тексті! Якщо їх 15, має бути 15 цін. Не зупиняйся на півдорозі.
10. hotel_stars: список зірковості.
11. baggage_info: опис багажу/пріоріті, напр: "багаж 20кг", "пріоріті", "додаткова ручна поклажа 10кг", або null якщо не згадано.
12. extras: список додаткових послуг. Кожна послуга - це об'єкт:
    {"type": "transfer" | "excursion" | "priority" | "baggage" | "shuttle", "original_name": "назва з тексту", "price": <число в євро>, "is_per_person": <true/false>}
    ПРАВИЛА РОЗРАХУНКУ:
    - Якщо менеджер сказав "за особу" або "на людину" -> is_per_person: true.
    - В іншому випадку (навіть якщо це пріоріті) -> is_per_person: false.
    - Якщо вказано кількість та ціну за одиницю (наприклад, "2 багажі по 120 євро кожен"), ОБОВ'ЯЗКОВО перемнож їх і запиши ЗАГАЛЬНУ суму (240).
    - Більше ЖОДНИХ математичних дій не роби. Запиши ціну.
    ОБОВ'ЯЗКОВО витягни ВСІ послуги (трансфери, пріоріті, екскурсії, багаж), про які згадав менеджер.
    Якщо нічого не згадано — порожній список [].

ФОРМАТ: Тільки JSON.
КРИТИЧНО: Не роби жодних математичних розрахунків. Просто витягни сирі цифри з тексту.
"""

_FORMAT_PROMPT = """Ти — професійний тревел-дизайнер. Твоє завдання: написати вступну частину повідомлення та блок рекомендацій.

БЛОК 1: ВСТУП (ПРИКЛАД):
Авіатур на Майорку 🇪🇸
Із Берліна 🇩🇪
🌤️ 15-25 червня, 10 ночей
Туди 22:10-00:55
Назад 15:35-18:20
[РЯДОК ПРО БАГАЖ — адаптуй до тексту менеджера]:
   - Якщо згадано пріоріті: 🧳 ручна поклажа до 10 кг та розміром 20х40х30 см + 1 додаткова пріоріті
   - Якщо згадано багаж: 🧳 ручна поклажа до 10 кг + 1 багаж [вага]кг
   - Якщо нічого: 🧳 ручна поклажа до 10 кг та розміром 20х40х30 см
[ДОДАТКОВІ ПОСЛУГИ — пиши ТІЛЬКИ ті, що згадані менеджером, вибирай точну назву без слешів]:
   - 🚖 Індивідуальний трансфер в 2 сторони (або до готелю, залежно від тексту)
   - 🚍 Шатл-бас в 2 сторони (або до готелю)
   - ⭐️ Екскурсійна програма
   СУВОРО: Якщо в оригінальному тексті НЕМАЄ слова 'екскурсія', 'програма', 'трансфер' або 'шатл' — ти ПОВИНЕН видалити цей рядок з повідомлення! Ніколи не пиши послугу за замовчуванням.

ВАЖЛИВО ПРО ФОРМАТУВАННЯ ВСТУПУ:
НЕ РОБИ жодних порожніх рядків між блоком багажу та блоком додаткових послуг! Вони мають йти суцільним списком. Також не роби порожніх рядків в кінці вступу.

БЛОК 2: РЕКОМЕНДАЦІЇ (ОБОВ'ЯЗКОВО):
- Оберіть ТІЛЬКИ 2-3 найкращих готелі з наданого списку. НЕ БІЛЬШЕ.
- Пишіть опис без заголовків з зірочками. Назва готелю має бути органічно вписана в текст.
- Пишіть емоційно, від першої особи.
- ПРИКЛАД ОПИСУ:
Хочу звернути вашу увагу на Globales Verdemar Apartamentos. Це чудовий варіант для тих, хто шукає комфортний відпочинок з видом на море. Тут ви знайдете просторі апартаменти...

Якщо ви шукаєте щось більш активне, BLUESEA Costa Verde стане чудовим вибором. Цей готель пропонує безліч розваг...

(порожній рядок між рекомендаціями)

ПРАВИЛА:
1. Поверни ТІЛЬКИ Вступ та Рекомендації.
2. Використовуй роздільник "===END_INTRO===" між Вступом та Рекомендаціями.
3. СУВОРО ЗАБОРОНЕНО: Не пиши нумерований список готелів (1, 2, 3...) та типи харчування у вступі. Тільки заголовок та деталі перельоту.
4. НЕ ПИШИ жодних цифр вартості, сум або цін (наприклад, "💰 загальна вартість туру...")! Рядок з цінами буде додано автоматично після твого тексту.
5. НЕ ПИШИ фразу "Ціна актуальна..." — я додам її сам.
"""

def calculate_tour_prices(hotel_prices: list, flight_per_person: float,
                          other_per_person: float, total_people: int,
                          has_children: bool, tourist_tax_per_person: float = 0) -> list:
    results = []
    for hotel_total in hotel_prices:
        hotel_per_person = hotel_total / total_people if total_people > 0 else hotel_total
        cost = hotel_per_person + flight_per_person + other_per_person + tourist_tax_per_person
        
        # Markup logic from April 28th
        if cost < 350:
            final_per_person = cost + 150
        else:
            final_per_person = cost * 1.43
        
        # Original rounding from April 28th
        final_per_person = round(final_per_person) + 5
        
        if has_children:
            results.append(round(final_per_person * total_people))
        else:
            results.append(round(final_per_person))
    return results

async def extract_prices_from_text(user_text: str, fast_models: list) -> dict:
    raw = await _call_llm_with_retry(
        messages=[
            {"role": "system", "content": _EXTRACT_PRICES_PROMPT},
            {"role": "user", "content": user_text},
        ],
        models=fast_models,
        timeout=20,
        response_format={"type": "json_object"}
    )
    if raw:
        try:
            return json.loads(raw)
        except: pass
    return {}

def _dedupe_keep_order(items: list[str]) -> list[str]:
    seen = set()
    result = []
    for item in items:
        key = item.strip().lower()
        if not key or key in seen:
            continue
        seen.add(key)
        result.append(item)
    return result

_MEAL_EXTRACT_PROMPT = """Ти — спеціаліст із туристичного харчування.
Твоє завдання: витягти тип харчування для КОЖНОГО готелю з тексту менеджера.

ПРАВИЛА:
1. Поверни список типів харчування у тому ж порядку, в якому готелі згадуються в тексті.
2. Якщо менеджер вказав тип харчування один раз для всіх готелів (наприклад, "харчування сніданки" на початку), то поверни "сніданки" для КОЖНОГО готелю у списку.
3. Використовуй ТІЛЬКИ ці назви:
   - "сніданки"
   - "сніданки + вечері"
   - "повний пансіон"
   - "все включено"
   - "ультра все включено"
   - "без харчування"
4. КІЛЬКІСТЬ: Якщо в тексті 15 готелів, у масиві "meals" має бути РІВНО 15 елементів. Знайди тип харчування для КОЖНОГО готелю, не зупиняйся завчасно!
5. ФОРМАТ: Тільки JSON {"meals": ["тип 1", "тип 2"]}.
"""

async def _call_llm_with_retry(messages, models, temperature=0, timeout=30, max_tokens=None, response_format=None):
    """Calls LLM with fallback models and retry logic."""
    for model in models:
        for attempt in range(2): # Try each model up to 2 times
            try:
                params = {
                    "model": model,
                    "messages": messages,
                    "temperature": temperature,
                    "timeout": timeout,
                }
                if max_tokens:
                    params["max_tokens"] = max_tokens
                if response_format:
                    params["response_format"] = response_format
                    
                resp = await client.chat.completions.create(**params)
                content = resp.choices[0].message.content.strip()
                # Basic cleanup of markdown
                content = re.sub(r'```[a-z]*\n?', '', content).strip('`').strip()
                if content:
                    return content
            except Exception as e:
                err_str = str(e).lower()
                if "429" in err_str or "rate limit" in err_str:
                    logger.warning(f"Rate limited on {model}, attempt {attempt+1}. Switching/Retrying...")
                    if attempt == 0:
                        await asyncio.sleep(1) # Small pause before retry
                        continue
                logger.warning(f"LLM call failed for {model}: {e}")
                break # Try next model
    return None

async def _extract_meals(user_text: str, fast_models: list) -> list:
    """Extract meal types for each hotel if mentioned."""
    raw = await _call_llm_with_retry(
        messages=[
            {"role": "system", "content": _MEAL_EXTRACT_PROMPT},
            {"role": "user", "content": user_text},
        ],
        models=fast_models,
        timeout=15,
        response_format={"type": "json_object"}
    )
    if raw:
        try:
            return json.loads(raw).get("meals", [])
        except: pass
    return []

def fuzzy_match_hotel(hotel_name: str, db: list) -> tuple[dict, float]:
    def normalize_name(name: str) -> str:
        # Remove stars from name for better matching
        cleaned = re.sub(r'[1-5]\s*(?:\*|★)', '', name.lower())
        
        # Simple Transliteration for Ukrainian/Russian names to Latin
        trans_map = {
            'а': 'a', 'б': 'b', 'в': 'v', 'г': 'g', 'ґ': 'g', 'д': 'd', 'е': 'e', 'ё': 'yo', 'ж': 'zh',
            'з': 'z', 'и': 'i', 'й': 'y', 'к': 'k', 'л': 'l', 'м': 'm', 'н': 'n', 'о': 'o',
            'п': 'p', 'р': 'r', 'с': 's', 'т': 't', 'у': 'u', 'ф': 'f', 'х': 'h', 'ц': 'ts',
            'ч': 'ch', 'ш': 'sh', 'щ': 'shch', 'ъ': '', 'ы': 'y', 'ь': '', 'э': 'e', 'ю': 'yu',
            'я': 'ya', 'і': 'i', 'ї': 'yi', 'є': 'ye'
        }
        
        # Replace common transcription errors and synonyms
        replacements = {
            "blucia": "bluesea", "blusia": "bluesea", "bluesee": "bluesea", "блю сі": "bluesea", "блюсі": "bluesea",
            "бі джей": "bj", "би джей": "bj", "біджей": "bj", "плеймар": "playamar", "playmar": "playamar",
            "blaucel": "bluesea", "багамас": "bahamas",
            "іберостар": "iberostar", "ріксос": "rixos", "мітсіс": "mitsis",
            "глікотель": "grecotel", "грекотель": "grecotel", "соль": "sol", "мелія": "melia",
            "хсм": "hsm", "каста": "costa", "калла": "cala", "calla": "cala", "міллер": "millor",
            "miller": "millor", "медіадіа": "mediodia", "mediadia": "mediodia", "глобаліс": "globales",
            "globalis": "globales", "ізабель": "isabel", "азулін": "azuline", "гранд": "gran", "grand": "gran"
        }
        
        for old, new in replacements.items():
            cleaned = cleaned.replace(old, new)
            
        # Try transliterating Cyrillic tokens
        tokens = cleaned.split()
        normalized_tokens = []
        for token in tokens:
            if any(ord(c) > 127 for c in token): # Has Cyrillic
                trans_token = "".join(trans_map.get(c, c) for c in token)
                normalized_tokens.append(trans_token)
            else:
                normalized_tokens.append(token)
        
        cleaned = " ".join(normalized_tokens)
            
        # Remove common separators and noise
        cleaned = re.sub(r'[^a-z0-9\s]', ' ', cleaned)
        cleaned = re.sub(r'\s+', ' ', cleaned).strip()
        final_tokens = [t for t in cleaned.split() if t not in _NOISE_TOKENS]
        return " ".join(final_tokens)

    best_match = None
    max_score = 0.0
    # Strip [NOT_FOUND] if present
    query_name = hotel_name.replace("[NOT_FOUND]", "").strip()
    query_strict = re.sub(r'\s*[1-5]\s*(?:\*|★)', '', query_name.lower()).strip()
    query = normalize_name(query_name)
    if not query:
        query = query_name.lower()
    
    query_words = set(re.findall(r'\w+', query))
    query_brands = query_words & BRANDS
    
    for h in db:
        db_name_orig = h['hotel']
        
        # 0. Exact strict match (ignoring case and stars)
        db_strict = re.sub(r'\s*[1-5]\s*(?:\*|★)', '', db_name_orig.lower()).strip()
        if query_strict and query_strict == db_strict:
            return h, 2.0
            
        db_name = normalize_name(db_name_orig)
        if not db_name:
            db_name = db_name_orig.lower()
            
        # 1. Exact match (after normalization)
        if query == db_name:
            return h, 1.5 # Increased bonus for exact match

        # 2. SequenceMatcher score
        ratio = difflib.SequenceMatcher(None, query, db_name).ratio()
        
        # 3. Word overlap bonus
        db_words = set(re.findall(r'\w+', db_name))
        db_brands = db_words & BRANDS
        if not query_words: continue
        
        overlap_words = query_words & db_words
        overlap = len(overlap_words)
        overlap_ratio = overlap / len(query_words) if query_words else 0

        # Weighted score: overlap is more important for identifying the right hotel
        score = ratio * 0.3 + overlap_ratio * 0.7
        
        # BRAND PENALTY/BONUS - Strict but Fair
        if query_brands and db_brands: # Both exist
            if query_brands != db_brands:
                score -= 1.0 # Brand conflict (e.g. Riu vs Iberostar)
            else:
                score += 0.4 # Brands matched
        # If manager forgot the brand, no penalty anymore to prevent filtering out correct hotels
        
        # UNIQUE WORD BONUS (e.g. "Playamar", "Java", "Isabel")
        # Words that are NOT brands and NOT common noise
        unique_query_words = query_words - BRANDS
        unique_db_words = db_words - BRANDS
        unique_overlap = len(unique_query_words & unique_db_words)
        if unique_query_words:
            unique_ratio = unique_overlap / len(unique_query_words)
            score += unique_ratio * 0.7 # Increased bonus
            
            # Additional penalty if query has unique words that are NOT in DB name
            # (e.g. Query="Blue Sea Cala Millor", DB="Cala Millor Garden")
            extra_words = unique_query_words - unique_db_words
            if extra_words:
                score -= len(extra_words) * 0.2 # Lowered from 0.7 to allow extra words like "Lifestyle" or "Beach"

        # Penalty for large length difference
        len_diff = abs(len(query) - len(db_name))
        if len_diff > 10:
            score -= 0.1

        if score > max_score:
            max_score = score
            best_match = h
            
    if best_match and max_score >= 0.90: # Strict 95%+ similarity required
        return best_match, max_score
        
    return {"hotel": hotel_name, "link": "Посилання відсутнє ⚠️"}, 0.0

def _build_hotel_candidates(user_text: str, relevant_hotels: list, limit: int = 250) -> list:
    if len(relevant_hotels) <= limit:
        return relevant_hotels
    
    text_norm = re.sub(r'[^a-z0-9а-яіїєґ\s]', ' ', user_text.lower())
    text_norm = re.sub(r'\s+', ' ', text_norm).strip()
    text_words = set(re.findall(r'\w+', text_norm))
    
    # Filter out noise from text words for scoring
    text_words_clean = text_words - _NOISE_TOKENS
    
    scored = []
    for hotel in relevant_hotels:
        name = hotel.get("hotel", "")
        name_norm = re.sub(r'[^a-z0-9а-яіїєґ\s]', ' ', name.lower())
        name_norm = re.sub(r'\s+', ' ', name_norm).strip()
        hotel_words = set(re.findall(r'\w+', name_norm))
        hotel_words_clean = hotel_words - _NOISE_TOKENS
        
        if not hotel_words_clean:
            scored.append((0, hotel))
            continue
            
        overlap = len(hotel_words_clean & text_words_clean)
        
        # Brands match
        brand_overlap = len(hotel_words & BRANDS & text_words)
        
        # Sequence ratio for fuzzy parts
        ratio = difflib.SequenceMatcher(None, text_norm, name_norm).ratio()
        
        # Weighted score: overlap is most important, then brands
        score = overlap * 5 + brand_overlap * 3 + ratio * 2
        scored.append((score, hotel))
        
    scored.sort(key=lambda x: x[0], reverse=True)
    return [h for _, h in scored[:limit]]

def _extract_allowed_stars(hotel_name: str) -> str:
    """Extract stars from hotel name in DB. Support patterns like '5*', '5★', '5 *', or just ' 5 ' at the end."""
    # Clean name from warnings for star extraction
    clean_name = hotel_name.replace('⚠️', '').strip()
    
    # Pattern 1: Digit followed by * or ★ (e.g., 5*, 5 ★)
    m = re.search(r'([1-5])\s*(?:\*|★)', clean_name)
    if m:
        return f"{m.group(1)}★"
    
    # Pattern 2: Just a digit at the very end or after a space (e.g., "Hotel Name 5")
    m = re.search(r'\s([1-5])(?:\s|$)', clean_name)
    if m:
        return f"{m.group(1)}★"
        
    return ""

def _inject_links(text: str, hotel_link_map: dict) -> str:
    lines = text.split('\n')
    i = 0
    while i < len(lines):
        if '🥑' in lines[i] and i > 0:
            hotel_line = lines[i - 1].lower()
            line_words = set(re.findall(r'\w+', hotel_line))
            best_name, best_score = None, 0.0
            for h_name, link in hotel_link_map.items():
                h_words = set(re.findall(r'\w+', h_name.lower()))
                if not h_words: continue
                overlap = len(h_words & line_words) / len(h_words)
                if overlap > best_score:
                    best_score = overlap
                    best_name = h_name
            if best_name and best_score >= 0.4:
                link = hotel_link_map[best_name]
                next_i = i + 1
                if next_i < len(lines):
                    nxt = lines[next_i]
                    if 'http' in nxt or 'Посилання відсутнє' in nxt:
                        if 'http' in nxt and link and 'http' in link and link.split()[0] not in nxt:
                            lines[next_i] = link
                        i += 1
                        continue
                lines.insert(i + 1, link)
                i += 1
        i += 1
    return '\n'.join(lines)

def _count_listed_hotels(text: str) -> int:
    # Count both "N)" patterns and "🥑" icons
    # Filter out duplicated blocks by counting unique 🥑 icons that are not too close
    meal_icons = text.count('🥑')
    numbered_matches = len(re.findall(r'^\s*\d+[\)\.]\s+', text, re.MULTILINE))
    
    # If the LLM put descriptions inside the list, we might have more text but we care about the blocks
    return max(numbered_matches, meal_icons)

def _build_price_line(price_label: str, computed_prices: list) -> str:
    if not computed_prices:
        return "💰 не вказано"
    prices_str = ", ".join([f"{i+1}){p}€" for i, p in enumerate(computed_prices)])
    return f"{price_label} - {prices_str}"

def _inject_prices(text: str, price_label: str, computed_prices: list) -> str:
    enforced_price_line = _build_price_line(price_label, computed_prices)
    lines = text.split("\n")
    
    # 1. First, remove ANY line that starts with 💰 to prevent duplicates
    lines = [line for line in lines if not line.strip().startswith("💰")]
    
    # 2. Find the anchor to insert the correct price line.
    # We want it AFTER "онлайн підтримка". If not found, then AFTER "путівник".
    anchor_idx = None
    
    # Try to find "онлайн підтримка" first
    for i, line in enumerate(lines):
        if "онлайн підтримка" in line.lower():
            anchor_idx = i
            break
            
    # If not found, try "путівник"
    if anchor_idx is None:
        for i, line in enumerate(lines):
            if "путівник" in line.lower():
                anchor_idx = i
                break
            
    if anchor_idx is not None:
        # Check if the next line is already a price line or similar (should be removed by step 1)
        lines.insert(anchor_idx + 1, enforced_price_line)
    else:
        # If no footer anchors found, the LLM might have failed to generate the footer.
        # We find the last hotel/meal line and insert after it.
        last_meal_idx = -1
        for i, line in enumerate(lines):
            if "🥑" in line:
                last_meal_idx = i
        
        if last_meal_idx != -1:
            # Insert after the last hotel block (usually hotel name + meal + link)
            # Find the end of that block
            target_idx = last_meal_idx + 1
            if target_idx < len(lines) and ("http" in lines[target_idx] or "Посилання" in lines[target_idx]):
                target_idx += 1
            
            # Add footer elements if missing
            footer = [
                "",
                "✔️ путівник + тур страхування",
                "🤓 онлайн підтримка 24/7",
                enforced_price_line,
                "",
                "❗️Ціна актуальна на момент розрахунку подорожі"
            ]
            for j, f_line in enumerate(footer):
                lines.insert(target_idx + j, f_line)
        else:
            # Total fallback
            lines.append("")
            lines.append(enforced_price_line)
        
    return "\n".join(lines)

def _append_missing_hotels(text: str, matched_hotels: list, hotel_prices: list) -> str:
    listed = _count_listed_hotels(text)
    total = len(matched_hotels)
    if listed >= total:
        return text
    lines = text.split("\n")
    insert_idx = len(lines)
    for i, line in enumerate(lines):
        stripped = line.strip()
        if stripped.startswith("✔️") or stripped.startswith("💰"):
            insert_idx = i
            break
    missing_block = []
    for idx in range(listed, total):
        h = matched_hotels[idx]
        stars_str = _extract_allowed_stars(h['hotel'])
        header = f"{idx + 1}) {h['hotel']}{(' ' + stars_str) if stars_str else ''}"
        meal = "🥑 не вказано"
        link_line = h['link']
        missing_block += ["", header, meal, link_line]
    for i, block_line in enumerate(missing_block):
        lines.insert(insert_idx + i, block_line)
    return "\n".join(lines)

def _pick_destination_by_keywords(user_text: str, destinations: list) -> str | None:
    text = user_text.lower()
    for key, normalized_dest in _DESTINATION_ALIASES.items():
        if key in text:
            for d in destinations:
                if normalized_dest in d.lower():
                    return d
    return None

def _fallback_hotel_extraction(user_text: str, candidate_hotels: list) -> list:
    """Non-LLM fallback: finds hotels by simple string matching/overlap when LLM fails."""
    if not candidate_hotels:
        return []
    
    # Pre-normalize the user text for better matching
    def normalize_for_fallback(t: str) -> str:
        t = t.lower()
        # Remove ordinals and common prefixes
        t = re.sub(r'\b(перший|другий|третій|четвертий|п’ятий|шостий|сьомий|восьмий|дев’ятий|десятий|одинадцятий|дванадцятий)\b', '', t)
        t = re.sub(r'\b(готель|отель|номер|варіант)\b', '', t)
        
        replacements = {
            "blucia": "bluesea", "blusia": "bluesea", "bluesee": "bluesea", "блю сі": "bluesea", "блюсі": "bluesea",
            "бі джей": "bj", "бі джи": "bj", "би джей": "bj", "би джи": "bj", "біджей": "bj", "плеймар": "playamar",
            "playmar": "playamar", "blaucel": "bluesea", "багамас": "bahamas", "casta": "costa", "calla": "cala",
            "mediadia": "mediodia", "globalis": "globales", "ізабель": "isabel", "азулін": "azuline",
            "каста": "costa", "калла": "cala", "міллер": "millor", "медіадіа": "mediodia", "глобаліс": "globales",
            "гранд": "gran", "grand": "gran",
            "bg": "bj", "bg ": "bj ", " bg": " bj" # Common transcription swap
        }
        for old, new in replacements.items():
            t = t.replace(old, new)
        return t

    text_norm = normalize_for_fallback(user_text)
    text_words = set(re.findall(r'\w+', text_norm))
    found_hotels = []
    
    # Sort hotels by length descending to match longer names first (e.g. "Hotel Brand Name" before "Hotel Brand")
    sorted_candidates = sorted(candidate_hotels, key=lambda x: len(x['hotel']), reverse=True)
    
    # 1. Check if the full normalized name is in the text
    # Prioritize exact matches first
    exact_matches = []
    for h in sorted_candidates:
        name = h['hotel']
        # Clean stars and trailing digits from DB name
        name_clean = re.sub(r'\s*[1-5]\s*(?:\*|★)', ' ', name.lower())
        name_clean = re.sub(r'\s+[1-5]\s*$', ' ', name_clean)
        
        name_clean = re.sub(r'[^a-z0-9а-яіїєґ\s]', ' ', name_clean)
        name_norm = normalize_for_fallback(name_clean)
        
        if len(name_norm) > 5 and name_norm in text_norm:
            exact_matches.append(name)
            
    # 2. Check word overlap for the rest
    fuzzy_matches = []
    for h in sorted_candidates:
        name = h['hotel']
        if name in exact_matches: continue
        
        # Clean stars and trailing digits from DB name
        name_clean = re.sub(r'\s*[1-5]\s*(?:\*|★)', ' ', name.lower())
        name_clean = re.sub(r'\s+[1-5]\s*$', ' ', name_clean)
        
        name_clean = re.sub(r'[^a-z0-9а-яіїєґ\s]', ' ', name_clean)
        name_norm = normalize_for_fallback(name_clean)
        name_words = [w for w in re.findall(r'\w+', name_norm) if w not in _NOISE_TOKENS]
        
        if not name_words: continue
            
        matches = 0
        for nw in name_words:
            if nw in text_words:
                matches += 1
            else:
                for tw in text_words:
                    if len(tw) > 3 and nw.startswith(tw[:3]) and difflib.SequenceMatcher(None, nw, tw).ratio() > 0.8:
                        matches += 1
                        break
        
        if matches / len(name_words) >= 0.95: 
            fuzzy_matches.append(name)
            
    # Sort by appearance in text
    all_found = exact_matches + fuzzy_matches
    all_found.sort(key=lambda name: text_norm.find(normalize_for_fallback(re.sub(r'[^a-z0-9а-яіїєґ\s]', ' ', name.lower()))))
    
    return _dedupe_keep_order(all_found)

def _count_potential_hotels(text: str) -> int:
    """Estimates how many hotels are mentioned based on numbering patterns."""
    text = text.lower()
    # Count patterns like "1 готель", "2 вариант", "3)", "4.", etc.
    # We avoid matching dates like 15.06 by requiring space after dot/parenthesis
    patterns = [
        r'(?:^|\n|\s)\d+\s*[)\.]\s+', # 1) or 1. at start or after space
        r'(?:^|\n|\s)\d+\s+(?:готель|отель|варіант|вариант)', # 1 готель
        r'(?:перший|другий|третій|четвертий|п[\'’]ятий|шостий|сьомий|восьмий|дев[\'’]ятий|десятий|одинадцятий|дванадцятий|тринадцятий|чотирнадцятий|п[\'’]ятнадцятий|шістнадцятий|наступний)\s+(?:готель|отель|варіант|вариант)'
    ]
    all_matches = set()
    for p in patterns:
        for m in re.finditer(p, text):
            # Only count if not inside a date-like pattern (e.g. 15.06)
            # Check if there's a number immediately after the space
            start = m.start()
            match_str = m.group()
            # If the pattern is like "15. ", check if next char is a digit
            if '.' in match_str:
                after_dot = text[m.end():m.end()+1]
                if after_dot.isdigit():
                    continue
            all_matches.add(start)
    
    count = len(all_matches)
    return count if count > 0 else 1

def _sort_hotels_by_appearance(hotels: list[str], text: str) -> list[str]:
    """Sorts hotel names based on their first appearance in the text."""
    text_lower = text.lower()
    
    # 1. First, find all occurrences of ordinal markers (1, 2, 3... or first, second...)
    ordinals = []
    ordinal_patterns = [
        (r'\b(\d+)\s*[)\.]\s+', 1),
        (r'\b(\d+)\s+(?:готель|отель|варіант|вариант)', 1),
        (r'\b(перший)\b', 1), (r'\b(другий)\b', 2), (r'\b(третій)\b', 3),
        (r'\b(четвертий)\b', 4), (r'\b(п[\'’]ятий)\b', 5), (r'\b(шостий)\b', 6),
        (r'\b(сьомий)\b', 7), (r'\b(восьмий)\b', 8), (r'\b(дев[\'’]ятий)\b', 9), (r'\b(десятий)\b', 10),
        (r'\b(одинадцятий)\b', 11), (r'\b(дванадцятий)\b', 12), (r'\b(тринадцятий)\b', 13), 
        (r'\b(чотирнадцятий)\b', 14), (r'\b(п[\'’]ятнадцятий)\b', 15), (r'\b(шістнадцятий)\b', 16),
        (r'\b(наступний)\b', 20) # Treat 'наступний' as a general marker
    ]
    
    found_ordinals = []
    for pattern, weight in ordinal_patterns:
        for m in re.finditer(pattern, text_lower):
            try:
                val = int(m.group(1)) if m.group(1).isdigit() else weight
                # If it's a word, weight is already the value
                if not m.group(1).isdigit():
                    val = weight
                found_ordinals.append((m.start(), val))
            except: pass
    
    found_ordinals.sort() # Sort by position in text

    # 2. For each hotel, find its position in text
    hotel_positions = []
    for h_name in hotels:
        clean_name = h_name.replace("[NOT_FOUND]", "").strip().lower()
        clean_name = re.sub(r'\s*[1-5]\s*(?:\*|★)', ' ', clean_name)
        clean_name = re.sub(r'[^a-z0-9а-яіїєґ\s]', ' ', clean_name)
        clean_name = re.sub(r'\s+', ' ', clean_name).strip()
        
        if not clean_name: 
            hotel_positions.append((999999, h_name))
            continue
            
        # Find all positions of this hotel
        pos = text_lower.find(clean_name)
        if pos == -1:
            # Try unique words
            words = [w for w in clean_name.split() if w not in _NOISE_TOKENS and len(w) > 3]
            if words:
                pos = min([text_lower.find(w) for w in words if text_lower.find(w) != -1] or [999999])
            else:
                pos = 999999
        
        # 3. Associate hotel with the nearest preceding ordinal
        best_ordinal = 999
        for o_pos, o_val in found_ordinals:
            # If ordinal is close before the hotel name (within 100 chars)
            if o_pos < pos and (pos - o_pos) < 150:
                best_ordinal = o_val
                break
        
        # If no ordinal found, use the raw position but deprioritize
        final_rank = best_ordinal * 1000000 + pos
        hotel_positions.append((final_rank, h_name))

    # Sort by rank (ordinal first, then position)
    hotel_positions.sort()
    
    seen = set()
    final = []
    for _, h in hotel_positions:
        if h.lower() not in seen:
            final.append(h)
            seen.add(h.lower())
    return final

async def format_tour_message(user_text: str, do_cleanup: bool = False, raw_voice_text: str = None) -> str:
    db = get_hotel_db()
    destinations = list(db.keys())
    
    # Text to use for hotel name extraction (raw is better for fuzzy matching)
    hotel_search_text = raw_voice_text if raw_voice_text else user_text
    
    potential_count = _count_potential_hotels(hotel_search_text)
    logger.info(f"Potential hotels count detected: {potential_count}")
    
    # Pre-clean hotel search text from common ordinal words that might confuse extraction
    hotel_search_text_cleaned = re.sub(r'\b(перший|другий|третій|четвертий|п’ятий|шостий|сьомий|восьмий|дев’ятий|десятий)\s+готель\b', 'готель', hotel_search_text.lower())
    
    # 0. Cleanup in parallel with initial detection if requested
    cleanup_task = None
    if do_cleanup:
        cleanup_task = asyncio.create_task(cleanup_transcribed_text(user_text))

    selected_dest = _pick_destination_by_keywords(hotel_search_text_cleaned, destinations)
    
    fast_models = ["openai/gpt-4o-mini", "google/gemini-2.5-flash"]
    smart_models = ["openai/gpt-4o-mini", "google/gemini-2.5-flash"]
    
    start_time = asyncio.get_event_loop().time()

    async def _detect_destination(text):
        if selected_dest: return selected_dest
        dest_content = f"ТЕКСТ:\n{text}\n\nДОСТУПНІ НАПРЯМКИ:\n{', '.join(destinations)}"
        
        raw = await _call_llm_with_retry(
            messages=[{"role": "system", "content": _DESTINATION_PROMPT}, {"role": "user", "content": dest_content}],
            models=fast_models,
            timeout=15
        )
        if raw:
            model_name = re.sub(r'[^\w\s-]', '', raw.lower()).strip()
            for d in destinations:
                if d.lower() in model_name or model_name in d.lower(): return d
        return destinations[0] if destinations else "Unknown"

    # ЗАВЖДИ беремо сирий текст для витягування цін та дат, бо очищений текст втрачає дані!
    base_text_for_data = raw_voice_text if raw_voice_text else user_text
    
    # Pass hint about potential count to price extractor
    price_content = base_text_for_data
    if potential_count > 1:
        price_content = f"(ВАЖЛИВО: Я очікую {potential_count} готелів)\n{base_text_for_data}"

    price_task = asyncio.create_task(extract_prices_from_text(price_content, fast_models))
    dest_task = asyncio.create_task(_detect_destination(hotel_search_text))
    meal_task = asyncio.create_task(_extract_meals(user_text, fast_models))
    
    async def _extract_hotels_broadly(text):
        raw = await _call_llm_with_retry(
            messages=[{"role": "system", "content": _EXTRACT_PROMPT}, {"role": "user", "content": f"ТЕКСТ МЕНЕДЖЕРА:\n{text}\n\nЗнайди всі готелі."}],
            models=fast_models,
            timeout=20,
            response_format={"type": "json_object"}
        )
        if raw:
            try:
                return json.loads(raw).get("hotels", [])
            except: pass
        return []

    broad_hotel_task = asyncio.create_task(_extract_hotels_broadly(hotel_search_text))
    
    # Wait for first-round tasks
    if cleanup_task:
        user_text = await cleanup_task
        
    selected_dest = await dest_task
    
    # ОЧИЩЕННЯ ВІД "20стр" (видаляємо цифри і слово "стр" з кінця)
    clean_dest_name = "Unknown"
    if selected_dest:
        clean_dest_name = re.sub(r'\s*\d+\s*стр.*', '', selected_dest, flags=re.IGNORECASE).strip().title()
    
    price_data = await price_task
    broad_hotels = await broad_hotel_task
    extracted_meals = await meal_task
    
    # NEW: If price_data has more hotels than we extracted, we need to be careful
    hotel_prices_raw = price_data.get("hotel_prices", {}) if price_data else {}
    if isinstance(hotel_prices_raw, dict):
        expected_count = len(hotel_prices_raw)
    elif isinstance(hotel_prices_raw, list):
        expected_count = len(hotel_prices_raw)
    else:
        expected_count = 0
    
    # If price extractor found fewer hotels than our heuristic, use the higher number as expected
    if potential_count > expected_count:
        logger.warning(f"Price extractor found {expected_count} but heuristic found {potential_count}. Using {potential_count} as target.")
        expected_count = potential_count
    
    logger.info(f"Step 1 parallel done in {asyncio.get_event_loop().time() - start_time:.2f}s. Dest: {clean_dest_name}")

    relevant_hotels = db.get(selected_dest, [])
    candidate_hotels = _build_hotel_candidates(hotel_search_text, relevant_hotels)
    
    # -----------------------------------------
    # STRICT DIRECT MATCHING PHASE
    # -----------------------------------------
    text_clean_for_search = re.sub(r'[^a-z0-9\s]', ' ', hotel_search_text_cleaned.lower())
    text_clean_for_search = re.sub(r'\s+', ' ', text_clean_for_search).strip()
    
    direct_matched_hotels = []
    for h in relevant_hotels:
        h_name = h['hotel'].lower()
        h_name = re.sub(r'\s*[1-5]\s*(?:\*|★)', ' ', h_name)
        h_name = re.sub(r'\s+[1-5]\s*$', ' ', h_name)
        
        h_clean = re.sub(r'[^a-z0-9\s]', ' ', h_name)
        h_clean = re.sub(r'\s+', ' ', h_clean).strip()
        h_words = [w for w in h_clean.split() if w not in _NOISE_TOKENS and len(w) > 2]
        
        if h_words:
            if h_clean in text_clean_for_search:
                direct_matched_hotels.append(h['hotel'])
            else:
                unique_db_words = set(h_words) - BRANDS
                text_words = set(text_clean_for_search.split())
                if unique_db_words and all(word in text_words for word in unique_db_words):
                    db_brands = set(h_words) & BRANDS
                    text_brands = text_words & BRANDS
                    if not db_brands or (db_brands & text_brands):
                        direct_matched_hotels.append(h['hotel'])
    
    if "playamar" in text_clean_for_search:
        direct_matched_hotels = [name for name in direct_matched_hotels if "playamar" in name.lower() or "playamar" not in " ".join(direct_matched_hotels).lower()]
    
    logger.info(f"Direct matching found: {direct_matched_hotels}")

    async def _do_targeted_extract(text_to_parse):
        # NO MORE CHUNKING: Send the candidates list to save context and speed up
        db_names = "\n".join([h['hotel'] for h in candidate_hotels])
        extraction_content = f"ТЕКСТ МЕНЕДЖЕРА:\n{text_to_parse}\n\nНАПРЯМОК: {clean_dest_name}\n\nБАЗА:\n{db_names}"
        
        if expected_count > 0:
            extraction_content += f"\n\nВАЖЛИВО: Я очікую знайти РІВНО {expected_count} готелів."
        
        raw = await _call_llm_with_retry(
            messages=[{"role": "system", "content": _EXTRACT_PROMPT}, {"role": "user", "content": extraction_content}],
            models=["openai/gpt-4o-mini", "google/gemini-2.5-flash"],
            timeout=60, # Increased timeout for larger context
            max_tokens=1500,
            response_format={"type": "json_object"}
        )
        if raw:
            try:
                return json.loads(raw).get("hotels", [])
            except: pass
        return []

    extracted_hotels = await _do_targeted_extract(hotel_search_text)
    logger.info(f"LLM extracted {len(extracted_hotels)} hotels: {extracted_hotels}")
    
    # Merge broad extraction results ONLY if the LLM found very few hotels and we expect more
    if expected_count > 0 and broad_hotels and len(extracted_hotels) < expected_count:
        logger.info(f"LLM found {len(extracted_hotels)} but we expect {expected_count}. Merging with broad search results...")
        existing_lower = {h.replace('[NOT_FOUND]', '').strip().lower() for h in extracted_hotels}
        for bh in broad_hotels:
            bh_clean = bh.replace('[NOT_FOUND]', '').strip().lower()
            if bh_clean not in existing_lower:
                extracted_hotels.append(bh)
                existing_lower.add(bh_clean)
            if len(extracted_hotels) >= expected_count: break
        logger.info(f"After merge: {len(extracted_hotels)} hotels")
    
    if not extracted_hotels or (expected_count > 0 and len(extracted_hotels) < expected_count):
        logger.info(f"LLM extraction found {len(extracted_hotels)} but expected {expected_count}. Trying fallback search...")
        fallback_hotels = _fallback_hotel_extraction(hotel_search_text, candidate_hotels)
        
        # If fallback found more or better matches, use it
        if len(fallback_hotels) >= expected_count:
            logger.info(f"Fallback found {len(fallback_hotels)} hotels, which meets expected count.")
            extracted_hotels = fallback_hotels
        elif len(fallback_hotels) > len(extracted_hotels):
            logger.info(f"Fallback found {len(fallback_hotels)} hotels, more than LLM. Using fallback.")
            extracted_hotels = fallback_hotels
        elif not extracted_hotels:
             # Try one more time with broader candidate list
             logger.info("Trying broad fallback search...")
             extracted_hotels = _fallback_hotel_extraction(hotel_search_text, relevant_hotels[:300])
             
    if not extracted_hotels and raw_voice_text:
        # If extraction from raw failed, try the cleaned text as last resort
        logger.info("Retrying extraction from cleaned text...")
        extracted_hotels = await _do_targeted_extract(user_text)
        if not extracted_hotels:
            extracted_hotels = _fallback_hotel_extraction(user_text, candidate_hotels)
    
    # Final check: if we have prices but fewer hotels, try to find missing hotels by simple word search
    if expected_count > 0 and len(extracted_hotels) < expected_count:
        logger.info(f"Still missing {expected_count - len(extracted_hotels)} hotels. Searching for unmatched candidates...")
        # Get blocks from text to find which "N готель" is missing
        text_lower = hotel_search_text.lower()
        
        recovered_hotels = [None] * expected_count
        # Fill in what we already have by checking their positions or just simple assignment if count matches
        # For now, let's try to find which "N готель" matches which extracted hotel
        for h in extracted_hotels:
            # Simple heuristic: if we can't find position, we'll fill gaps later
            recovered_hotels[extracted_hotels.index(h)] = h

        for i in range(expected_count):
            if recovered_hotels[i] is not None: continue
            
            # Try to find a hotel that is mentioned near "i+1 готель"
            ordinal_pattern = rf"(?:{i+1}|{['перший','другий','третій','четвертий','п’ятий','шостий','сьомий','восьмий','дев’ятий','десятий'][i]})\s*(?:готель|отель|варіант)"
            context_match = re.search(ordinal_pattern + r"(.*?)(?:\d+\s*(?:готель|отель|варіант)|$)", text_lower, re.DOTALL)
            
            if context_match:
                context_text = context_match.group(1)
                # Search for any hotel from DB in this specific context
                for h in relevant_hotels:
                    # We need to use a normalization that is available in this scope
                    # fuzzy_match_hotel has a nested normalize_name, but we can't call it directly.
                    # We'll use a simplified version or use fuzzy_match_hotel itself.
                    h_name_clean = re.sub(r'[^a-z0-9\s]', ' ', h['hotel'].lower())
                    unique_words = set(h_name_clean.split()) - BRANDS - _NOISE_TOKENS
                    if unique_words and all(word in context_text for word in unique_words):
                        recovered_hotels[i] = h['hotel']
                        break
        
        # Filter out None and keep the original LLM order
        extracted_hotels = [h for h in recovered_hotels if h is not None]
    else:
        # Keep extracted hotels in their original LLM order
        pass
    
    # Final sync and price extraction refinement
    hotel_prices_map = price_data.get("hotel_prices", {}) if price_data else {}
    if isinstance(hotel_prices_map, list):
        new_map = {}
        for idx, p in enumerate(hotel_prices_map):
            new_map[f"Hotel {idx+1}"] = p
        hotel_prices_map = new_map

    hotel_link_map = {}
    all_hotels_list = [hotel for hotels in db.values() for hotel in hotels]
    matched_hotels = []
    seen_hotels = set()
    final_hotel_prices_raw = []
    prices_dict = price_data.get("hotel_prices", {}) if price_data else {}
    
    # Matching extracted names with DB to get links and full names
    for h_name in extracted_hotels:
        match, score = fuzzy_match_hotel(h_name, relevant_hotels)
        if score < 0.65 and all_hotels_list:
            global_match, g_score = fuzzy_match_hotel(h_name, all_hotels_list)
            if g_score > 0.65:
                match, score = global_match, g_score
        
        display_name = match["hotel"]
        stars = _extract_allowed_stars(display_name)
        display_name = re.sub(r'\s*[1-5]\s*(?:\*|★)', '', display_name).strip()

        if score >= 0.90:
            # We found a match
            pass
        elif "[NOT_FOUND]" in h_name:
            display_name = h_name.replace("[NOT_FOUND]", "").strip() + " ⚠️ (немає в базі)"
            match = {"hotel": display_name, "link": "Посилання відсутнє ⚠️"}
        elif score < 0.90:
            display_name = f"{h_name} ⚠️ (немає в базі)"
            match = {"hotel": display_name, "link": "Посилання відсутнє ⚠️"}

        if stars and stars not in display_name:
            display_name = f"{display_name} {stars}"

        key = display_name.strip().lower()
        if key in seen_hotels: continue
        seen_hotels.add(key)
        
        match["hotel"] = display_name
        matched_hotels.append(match)
        hotel_link_map[display_name.lower()] = match['link']

    # Syncing prices with the matched_hotels list using extracted_hotels as keys
    # Якщо нейромережа все одно повернула масив (підстраховка)
    if isinstance(prices_dict, list):
        # Якщо модель помилилась і повернула масив замість словника
        for p in prices_dict:
            try:
                p_clean = re.sub(r'[^\d.]', '', str(p).replace(',', '.'))
                final_hotel_prices_raw.append(float(p_clean) if p_clean else 0.0)
            except:
                pass
    else:
        # СУВОРИЙ МАПІНГ: Шукаємо ціну саме для цього готелю
        for h_info in matched_hotels:
            hotel_name = h_info['hotel'].replace("⚠️", "").strip()
            # Clean stars/ratings from name for search
            hotel_name = re.sub(r'\s*[1-5]\s*(?:\*|★)', '', hotel_name).strip()
            
            val = prices_dict.get(hotel_name, 0)
            
            # Нечіткий пошук у словнику, якщо ключ трохи відрізняється
            if not val and prices_dict:
                best_match_key = None
                best_score = 0
                generic_words = {"hotel", "resort", "beach", "spa", "village", "apartments", "apartamentos", "suites", "boutique", "готель"}
                n_norm_words = set(re.sub(r'[^a-zа-яіїєґ0-9\s]', ' ', hotel_name.lower()).split()) - generic_words
                
                for k, v in prices_dict.items():
                    k_norm_str = re.sub(r'[^a-zа-яіїєґ0-9]', '', k.lower())
                    n_norm_str = re.sub(r'[^a-zа-яіїєґ0-9]', '', hotel_name.lower())
                    if k_norm_str in n_norm_str or n_norm_str in k_norm_str:
                        best_match_key = k
                        break
                    
                    # Word overlap
                    k_norm_words = set(re.sub(r'[^a-zа-яіїєґ0-9\s]', ' ', k.lower()).split()) - generic_words
                    if k_norm_words and n_norm_words:
                        overlap = len(k_norm_words & n_norm_words) / max(len(k_norm_words), len(n_norm_words))
                        if overlap > best_score:
                            best_score = overlap
                            best_match_key = k
                
                if best_match_key and (best_score >= 0.35 or best_match_key == k):
                    val = prices_dict[best_match_key]
                        
            try:
                p_clean = re.sub(r'[^\d.]', '', str(val).replace(',', '.'))
                final_hotel_prices_raw.append(float(p_clean) if p_clean else 0.0)
            except:
                final_hotel_prices_raw.append(0.0) # Якщо не знайшли, ставимо 0, а не дублюємо чужу ціну

    # Добиваємо нулями, якщо готелів більше, ніж цін
    while len(final_hotel_prices_raw) < len(matched_hotels):
        final_hotel_prices_raw.append(0.0)
        
    final_hotel_prices_raw = final_hotel_prices_raw[:len(matched_hotels)]

    computed_prices = []
    has_children = False
    price_label = "💰 загальна вартість туру за особу"

    if price_data and final_hotel_prices_raw:
        adults = _safe_int(price_data.get("adults"), 2)
        children = _safe_int(price_data.get("children"), 0)
        infants = _safe_int(price_data.get("infants"), 0)
        total_people = adults + children
        has_children = (children + infants) > 0
        
        # NEW: Raw flight price (might be total or per person)
        flight_raw_val = price_data.get("flight_price") or price_data.get("flight_total") or price_data.get("flight_per_person") or 0
        flight_per_person = 0.0
        try:
            f_clean = float(re.sub(r'[^\d.]', '', str(flight_raw_val).replace(',', '.')) or 0)
            is_per_person = price_data.get("flight_is_per_person", False)
            
            if not is_per_person and f_clean > 500 and total_people > 1:
                flight_per_person = f_clean / total_people
            elif not is_per_person and total_people > 1:
                # If they specify a small flight price but don't explicitly say "per person", assume it's total if we have >1 people
                flight_per_person = f_clean / total_people
            else:
                flight_per_person = f_clean
        except: pass
        
        extras_list = price_data.get("extras", [])
        extras_per_person = 0.0
        for e in extras_list:
            try:
                if "price_per_person" in e:
                    extras_per_person += float(e["price_per_person"])
                else:
                    p = float(e.get("price", 0))
                    if e.get("is_per_person"):
                        extras_per_person += p
                    else:
                        extras_per_person += p / total_people if total_people > 0 else p
            except: pass
        
        nights = _safe_int(price_data.get("nights"), 7)
        month = _safe_int(price_data.get("check_in_month"), 6)
        hotel_stars_list = price_data.get("hotel_stars") or []
        
        for idx, hotel_total in enumerate(final_hotel_prices_raw):
            stars_val = 0
            db_stars_str = _extract_allowed_stars(matched_hotels[idx]['hotel']) if idx < len(matched_hotels) else ""
            if db_stars_str:
                m_stars = re.search(r'\d', db_stars_str)
                stars_val = int(m_stars.group()) if m_stars else 0
            elif hotel_stars_list and idx < len(hotel_stars_list):
                stars_val = _safe_int(hotel_stars_list[idx])
            
            tax_info = get_tax_info(selected_dest or "", stars_val, month)
            if tax_info['per_room']:
                total_tax_for_stay = tax_info['rate'] * nights
            else:
                total_tax_for_stay = tax_info['rate'] * nights * adults
            
            tax_per_person_share = total_tax_for_stay / total_people if total_people > 0 else 0
            
            # MATH LOGIC FROM APRIL 28TH + EXTRAS + TAX
            hotel_per_person = hotel_total / total_people if total_people > 0 else hotel_total
            base_cost_no_tax = hotel_per_person + flight_per_person + extras_per_person + tax_per_person_share
            
            if base_cost_no_tax < 350:
                final_no_tax = base_cost_no_tax + 150
            else:
                final_no_tax = base_cost_no_tax * 1.43
            
            # Tax is already inside base_cost_no_tax, so it was multiplied.
            # Add +5 as insurance
            final = round(final_no_tax) + 5
            
            if has_children:
                computed_prices.append(round(final * total_people))
            else:
                computed_prices.append(final)
        
        price_label = "💰 загальна вартість туру за всіх" if has_children else "💰 загальна вартість туру за особу"

    # Prepare data for LLM formatting
    hotels_info = []
    for i, h in enumerate(matched_hotels):
        stars = _extract_allowed_stars(h['hotel'])
        meal = extracted_meals[i] if extracted_meals and i < len(extracted_meals) else "не вказано"
        price = computed_prices[i] if i < len(computed_prices) else "не вказано"
        hotels_info.append(f"{i+1}) {h['hotel']} (ЗІРКИ: {stars if stars else 'немає'}) | Харчування: {meal} | Посилання: {h['link']} | ЦІНА: {price}€")
    
    db_text = "\n".join(hotels_info)
    
    extras_desc = ""
    _baggage = price_data.get("baggage_info") if price_data else None
    _extras = price_data.get("extras", []) if price_data else []
    
    # If priority/baggage is in extras but not in baggage_info, sync them
    if not _baggage:
        prio_item = next((e for e in _extras if e.get('type') in ['priority', 'baggage']), None)
        if prio_item:
            _baggage = prio_item.get('original_name')

    if _baggage:
        extras_desc += f"\nБАГАЖ ТА ПРІОРІТІ: {_baggage}"
    
    if _extras:
        items = []
        for e in _extras:
            etype = str(e.get('type', 'послуга')).upper()
            ename = e.get('original_name', 'послуга')
            items.append(f"- {etype}: {ename}")
        extras_desc += f"\nДОДАТКОВІ ПОСЛУГИ ТЕКСТОМ:\n" + "\n".join(items)
    
    combined_content = f"ТЕКСТ МЕНЕДЖЕРА:\n{user_text}\n\nНАПРЯМОК: {clean_dest_name}\n\n"
    combined_content += f"БАЗА ГОТЕЛІВ ТА ЦІНИ (ВИКОРИСТОВУЙ ВСЕ):\n{db_text}\n{extras_desc}\n\n"
    combined_content += f"РОЗРАХОВАНІ ЦІНИ (ДЛЯ РЯДКА З ЦІНАМИ):\n{price_label} - {', '.join([f'{i+1}){p}€' for i, p in enumerate(computed_prices)])}"

    result = await _call_llm_with_retry(
        messages=[{"role": "system", "content": _FORMAT_PROMPT}, {"role": "user", "content": combined_content}],
        models=smart_models,
        timeout=90,
        max_tokens=3000
    )
    
    # Prepare meals alignment for the programmatic block
    hotel_meal_list = []
    for i in range(len(matched_hotels)):
        if extracted_meals and i < len(extracted_meals):
            hotel_meal_list.append(extracted_meals[i])
        else:
            hotel_meal_list.append("не вказано")

    if result:
        # 1. Start with the LLM-generated intro and recommendations
        # (We assume LLM followed the instruction to provide Intro and Recommendations)
        
        # 2. Build the "Options" block programmatically (100% precision)
        options_block = "🏠 варіанти проживання:\n\n"
        for i, hotel_data in enumerate(matched_hotels, 1):
            name = hotel_data['hotel']
            # Link mapping: use the exact name as stored in matched_hotels
            link = hotel_data['link']
            meal = hotel_meal_list[i-1] # Use the prepared list
            
            options_block += f"{i}) {name}\n"
            options_block += f"🥑 {meal}\n"
            options_block += f"{link}\n\n"
            
        # 3. Build the "Footer" block programmatically
        footer_block = "✔️ путівник + тур страхування\n"
        footer_block += "🤓 онлайн підтримка 24/7\n"
        footer_block += f"{price_label} - "
        
        price_strings = []
        for i, p in enumerate(computed_prices, 1):
            price_strings.append(f"{i}){p}€")
        footer_block += ", ".join(price_strings) + "\n\n"
        footer_block += "❗️Ціна актуальна на момент розрахунку подорожі\n\n"

        # 4. Assemble the final message
        # Split LLM result into Intro and Recommendations using the delimiter
        if "===END_INTRO===" in result:
            parts = result.split("===END_INTRO===")
            intro = parts[0].strip()
            recommendations_raw = parts[1].strip()
        else:
            # Fallback if LLM didn't use the delimiter
            lines = result.split("\n")
            intro_lines = []
            recommendation_lines = []
            is_recommendation = False
            for line in lines:
                if "**" in line and not is_recommendation:
                    is_recommendation = True
                if is_recommendation:
                    recommendation_lines.append(line)
                else:
                    intro_lines.append(line)
            intro = "\n".join(intro_lines).strip()
            recommendations_raw = "\n".join(recommendation_lines).strip()
        
        # Жорсткий контроль екскурсійної програми та трансферів
        text_for_check = (user_text + " " + (raw_voice_text or "")).lower()
        has_excursion = "екскурс" in text_for_check
        has_transfer = "трансфер" in text_for_check
        has_shuttle = "шатл" in text_for_check or "shuttle" in text_for_check
        
        intro_lines_split = intro.split("\n")
        new_intro_lines = []
        excursion_present = False
        
        for line in intro_lines_split:
            line_lower = line.lower()
            if "екскурс" in line_lower or "⭐️ екскурсійна" in line_lower:
                if has_excursion:
                    excursion_present = True
                    new_intro_lines.append(line)
                # Якщо не казав, а LLM додала - рядок просто видаляється
            elif "трансфер" in line_lower or "🚖 індивідуальний" in line_lower:
                if has_transfer:
                    new_intro_lines.append(line)
            elif "шатл" in line_lower or "shuttle" in line_lower or "🚍 шатл" in line_lower:
                if has_shuttle:
                    new_intro_lines.append(line)
            else:
                new_intro_lines.append(line)
                
        if has_excursion and not excursion_present:
            # Якщо менеджер казав, а LLM забула - додаємо обов'язково
            new_intro_lines.append("   - ⭐️ Екскурсійна програма")
                
        intro = "\n".join(new_intro_lines).strip()

        # Programmatically limit recommendations to 3
        rec_parts = re.split(r'\n(?=\*\*)', recommendations_raw)
        recommendations = "\n".join(rec_parts[:3]).strip()
        
        # Assemble: Intro -> Options -> Footer -> Recommendations
        final_message = f"{intro}\n{options_block}{footer_block}\n{recommendations}"
        
        # Final cleanup of multiple newlines and extra symbols
        final_message = re.sub(r'\n{3,}', '\n\n', final_message).strip()
        
        return final_message
    
    return "❌ Помилка генерації тексту."
