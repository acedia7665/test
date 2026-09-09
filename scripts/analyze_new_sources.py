#!/usr/bin/env python3
"""
Разовый скрипт: находит в папке с PDF-статьями файлы, которых ещё нет
в итоговом Excel-файле, прогоняет каждый через Claude по промпту
prompts/industry-digest-source-analysis.md и дописывает по нему строку
в digest.xlsx.

Запуск (после каждого добавления новых PDF в папку):

    python scripts/analyze_new_sources.py

Настройки — через .env (см. .env.example) или переменные окружения:
    ANTHROPIC_API_KEY  — обязателен
    PDF_DIR            — папка с PDF (по умолчанию: pdfs)
    OUTPUT_XLSX        — итоговый файл (по умолчанию: digest.xlsx)
    CLAUDE_MODEL       — модель Claude (по умолчанию: claude-sonnet-5)
"""

from __future__ import annotations

import json
import os
import sys
from dataclasses import dataclass, fields
from pathlib import Path

import pdfplumber
from anthropic import Anthropic
from dotenv import load_dotenv
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

REPO_ROOT = Path(__file__).resolve().parent.parent
PROMPT_PATH = REPO_ROOT / "prompts" / "industry-digest-source-analysis.md"

# Ограничение на длину текста, отправляемого в модель (символы).
# Этого с большим запасом хватает на любую статью/отчёт.
MAX_TEXT_CHARS = 150_000

JSON_OUTPUT_INSTRUCTIONS = """
## Формат вывода для автоматической обработки (ВАЖНО, переопределяет раздел выше)

Для этого запуска НЕ используй человекочитаемый формат из раздела
«Формат ответа для каждого источника» — вместо этого верни ОДИН валидный
JSON-объект и больше ничего (без markdown-разметки, без ```json, без
пояснений до или после). Структура объекта строго такая:

{
  "title": "название материала",
  "link": "ссылка или пустая строка, если не указана в тексте",
  "date": "дата публикации в свободном формате, как указана в источнике",
  "author_org": "автор или организация; если не указано — 'Автор не указан'",
  "source_type": "тип источника",
  "relevance_rating": "Высокая | Средняя | Низкая",
  "relevance_comment": "1-2 предложения",
  "reliability_rating": "Высокая | Средняя | Низкая",
  "reliability_comment": "короткое объяснение",
  "theses": ["тезис 1", "тезис 2", "..."],
  "main_interest": "главный интерес для дайджеста одним предложением",
  "images_found": "краткое описание найденных изображений/графиков/схем",
  "images_rating": "Хорошие для использования | Можно использовать при необходимости | Малоинформативные | Полезных изображений нет",
  "final_score": 1,
  "final_comment": "короткое объяснение итоговой оценки"
}

final_score — целое число от 1 до 5. Все текстовые поля — строки на русском
языке. Если что-то невозможно определить из текста — прямо укажи это словами
внутри соответствующего поля (не выдумывай данные), но JSON-структуру не нарушай.
"""


@dataclass
class SourceAnalysis:
    title: str = ""
    link: str = ""
    date: str = ""
    author_org: str = ""
    source_type: str = ""
    relevance_rating: str = ""
    relevance_comment: str = ""
    reliability_rating: str = ""
    reliability_comment: str = ""
    theses: str = ""  # склеенные через "; " для одной ячейки
    main_interest: str = ""
    images_found: str = ""
    images_rating: str = ""
    final_score: str = ""
    final_comment: str = ""


COLUMN_HEADERS = [
    "Файл",
    "Название",
    "Ссылка",
    "Дата",
    "Автор / организация",
    "Тип источника",
    "Оценка актуальности",
    "Комментарий (актуальность)",
    "Оценка надёжности",
    "Комментарий (надёжность)",
    "Основные тезисы",
    "Главный интерес для дайджеста",
    "Изображения: что найдено",
    "Оценка изображений",
    "Итоговая оценка (1-5)",
    "Комментарий (итог)",
    "Дата обработки",
]


def load_system_prompt() -> str:
    if not PROMPT_PATH.exists():
        sys.exit(f"Не найден файл промпта: {PROMPT_PATH}")
    base_prompt = PROMPT_PATH.read_text(encoding="utf-8")
    return base_prompt + "\n" + JSON_OUTPUT_INSTRUCTIONS


def extract_pdf_text(pdf_path: Path) -> str:
    chunks: list[str] = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            if text:
                chunks.append(text)
    full_text = "\n\n".join(chunks).strip()
    if len(full_text) > MAX_TEXT_CHARS:
        full_text = full_text[:MAX_TEXT_CHARS] + "\n\n[...текст обрезан по длине...]"
    return full_text


def call_claude(client: Anthropic, model: str, system_prompt: str, pdf_name: str, pdf_text: str) -> dict:
    if not pdf_text.strip():
        raise ValueError("не удалось извлечь текст из PDF (возможно, скан без OCR)")

    message = client.messages.create(
        model=model,
        max_tokens=2000,
        system=system_prompt,
        messages=[
            {
                "role": "user",
                "content": (
                    f"Файл источника: {pdf_name}\n\n"
                    f"Текст источника:\n\n{pdf_text}"
                ),
            }
        ],
    )
    raw = "".join(block.text for block in message.content if getattr(block, "type", "") == "text").strip()

    # На случай, если модель всё же обернёт JSON в ```json ... ```.
    if raw.startswith("```"):
        raw = raw.strip("`")
        if raw.lower().startswith("json"):
            raw = raw[4:]
        raw = raw.strip()

    try:
        return json.loads(raw)
    except json.JSONDecodeError as exc:
        raise ValueError(f"модель вернула невалидный JSON: {exc}\nОтвет модели:\n{raw[:1000]}") from exc


def to_row(pdf_name: str, data: dict, processed_at: str) -> list[str]:
    theses = data.get("theses") or []
    if isinstance(theses, list):
        theses_str = "; ".join(str(t) for t in theses)
    else:
        theses_str = str(theses)

    return [
        pdf_name,
        str(data.get("title", "")),
        str(data.get("link", "")),
        str(data.get("date", "")),
        str(data.get("author_org", "")),
        str(data.get("source_type", "")),
        str(data.get("relevance_rating", "")),
        str(data.get("relevance_comment", "")),
        str(data.get("reliability_rating", "")),
        str(data.get("reliability_comment", "")),
        theses_str,
        str(data.get("main_interest", "")),
        str(data.get("images_found", "")),
        str(data.get("images_rating", "")),
        str(data.get("final_score", "")),
        str(data.get("final_comment", "")),
        processed_at,
    ]


def load_or_create_workbook(output_path: Path):
    if output_path.exists():
        wb = load_workbook(output_path)
        ws = wb.active
        already_processed = set()
        for row in ws.iter_rows(min_row=2, max_col=1, values_only=True):
            if row and row[0]:
                already_processed.add(row[0])
        return wb, ws, already_processed

    wb = Workbook()
    ws = wb.active
    ws.title = "Дайджест"
    ws.append(COLUMN_HEADERS)
    for i, header in enumerate(COLUMN_HEADERS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = max(14, min(40, len(header) + 4))
    return wb, ws, set()


def main() -> None:
    load_dotenv(REPO_ROOT / ".env")

    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        sys.exit(
            "ANTHROPIC_API_KEY не задан. Скопируй .env.example в .env и вставь ключ "
            "(получить: https://console.anthropic.com/settings/keys)."
        )

    pdf_dir = REPO_ROOT / os.environ.get("PDF_DIR", "pdfs")
    output_path = REPO_ROOT / os.environ.get("OUTPUT_XLSX", "digest.xlsx")
    model = os.environ.get("CLAUDE_MODEL", "claude-sonnet-5")

    if not pdf_dir.exists():
        sys.exit(f"Папка с PDF не найдена: {pdf_dir}")

    pdf_files = sorted(p for p in pdf_dir.iterdir() if p.suffix.lower() == ".pdf")
    if not pdf_files:
        print(f"В папке {pdf_dir} нет PDF-файлов.")
        return

    wb, ws, already_processed = load_or_create_workbook(output_path)
    new_files = [p for p in pdf_files if p.name not in already_processed]

    if not new_files:
        print("Новых PDF нет — все файлы уже есть в Excel.")
        return

    system_prompt = load_system_prompt()
    client = Anthropic(api_key=api_key)

    from datetime import datetime, timezone

    processed_count = 0
    error_count = 0

    for pdf_path in new_files:
        print(f"Обрабатываю: {pdf_path.name} ...")
        try:
            pdf_text = extract_pdf_text(pdf_path)
            data = call_claude(client, model, system_prompt, pdf_path.name, pdf_text)
            processed_at = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")
            ws.append(to_row(pdf_path.name, data, processed_at))
            wb.save(output_path)  # сохраняем после каждого файла — прогресс не теряется
            processed_count += 1
            print(f"  готово, оценка: {data.get('final_score', '?')}/5")
        except Exception as exc:  # noqa: BLE001 — умышленно широкий catch, чтобы не прерывать пакет
            error_count += 1
            print(f"  ОШИБКА: {exc}", file=sys.stderr)
            # Файл не отмечается как обработанный — при следующем запуске попытка повторится.

    print(
        f"\nГотово. Обработано: {processed_count}, ошибок: {error_count}, "
        f"пропущено (уже в Excel): {len(pdf_files) - len(new_files)}."
    )
    print(f"Файл: {output_path}")


if __name__ == "__main__":
    main()
