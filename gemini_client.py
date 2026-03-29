import json
import logging
import os
import re

import google.generativeai as genai
from dotenv import load_dotenv

load_dotenv()

logger = logging.getLogger(__name__)

genai.configure(api_key=os.getenv("GEMINI_API_KEY"))

_model = genai.GenerativeModel("gemini-2.0-flash")

_SYSTEM_PROMPT = """
Ты — профессиональный бизнес-аналитик, который составляет коммерческие предложения (КП).
Твоя задача — сгенерировать контент для КП строго в формате JSON.

Верни ТОЛЬКО валидный JSON без каких-либо пояснений, markdown-блоков или дополнительного текста.
Структура JSON:
{
  "kp_number": "SP26-XX",
  "kp_date": "дата в формате ДД.ММ.ГГГГ",
  "company_name": "Название компании клиента",
  "contact_person": "Контактное лицо клиента",
  "service_title": "Название услуги/продукта",
  "service_description": "Краткое описание услуги (2-3 абзаца, которое идёт в таблицу рядом с ценой)",
  "service_description_full": "Полное подробное описание услуги с пронумерованными пунктами (раздел Описание услуг)",
  "intro_text": "Вводный текст КП перед таблицей",
  "outro_text": "Заключительный текст после таблицы / призыв к действию",
  "price_monthly": "сумма в месяц (только число с пробелами, например 500 000,00)",
  "price_annual": "сумма за 12 месяцев (только число с пробелами, например 6 000 000,00)",
  "total_amount": "итоговая сумма с НДС (например 6 960 000,00 тенге)",
  "validity_period": "30 календарных дней",
  "price_table": [
    {"item": "Название позиции", "qty": "1", "unit": "шт", "price": "500 000,00", "total": "6 000 000,00"}
  ]
}"""


def _extract_json(text: str) -> dict:
    """Extract JSON from model response, stripping markdown fences if present."""
    text = text.strip()
    # Remove markdown code fences
    text = re.sub(r"^```(?:json)?\s*", "", text)
    text = re.sub(r"\s*```$, "", text)
    return json.loads(text)


def generate_kp_content(user_request: str, history: list) -> dict:
    """Generate KP content via Gemini API.

    Args:
        user_request: The user's request describing the KP to generate.
        history: List of previous messages for context
                 (each item is {"role": "user"|"model", "parts": [str]}).

    Returns:
        A dict with KP content fields.

    Raises:
        ValueError: If Gemini returns invalid JSON after retries.
    """    
    messages = [{"role": "user", "parts": [_SYSTEM_PROMPT]}]
    messages.extend(history)
    messages.append({"role": "user", "parts": [user_request]})

    last_error = None
    for attempt in range(3):
        try:
            response = _model.generate_content(messages)
            raw = response.text
            logger.debug("Gemini raw response (attempt %d): %s", attempt + 1, raw)
            content = _extract_json(raw)
            # Basic validation — ensure required keys are present
            required_keys = {
                "kp_number", "kp_date",
                "company_name", "contact_person", "service_title",
                "service_description", "service_description_full",
                "price_monthly", "price_annual",
                "total_amount", "validity_period",
                "intro_text", "outro_text", "price_table",
            }
            missing = required_keys - content.keys()
            if missing:
                raise ValueError(f"Missing keys in response: {missing}")
            return content
        except (json.JSONDecodeError, ValueError) as exc:
            last_error = exc
            logger.warning("Attempt %d failed: %s", attempt + 1, exc)
            messages.append(
                {
                    "role": "user",
                    "parts": [
                        "Ошибка: ты вернул невалидный JSON или пропустил обязательные поля. "
                        "Пожалуйста, верни ТОЛЬКО валидный JSON без markdown и пояснений, "
                        "со всеми обязательными полями включая kp_number, kp_date, "
                        "price_monthly, price_annual, service_description_full."
                    ],
                }
            )

    raise ValueError(f"Не удалось получить валидный JSON от Gemini: {last_error}")