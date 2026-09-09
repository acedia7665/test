# test

## Industry digest source analysis

[`prompts/industry-digest-source-analysis.md`](prompts/industry-digest-source-analysis.md)
содержит системный промпт/шаблон для AI-агента, который выполняет предварительный
анализ источников (статей, отчётов, пресс-релизов, патентов и т.д.) для отраслевого
дайджеста: определяет основную информацию об источнике, оценивает актуальность и
надёжность, выделяет ключевые тезисы и полезные изображения, и выставляет итоговую
оценку пригодности от 1 до 5 — по единому формату, готовому к использованию.

## Автоматическая обработка PDF → Excel

[`scripts/analyze_new_sources.py`](scripts/analyze_new_sources.py) прогоняет
PDF-статьи из папки `pdfs/` через промпт выше и дописывает результат строками
в общий файл `digest.xlsx`. Уже добавленные ранее файлы не обрабатываются
повторно — скрипт сверяется с тем, что уже есть в Excel.

### Установка

```bash
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
cp .env.example .env
```

Открой `.env` и вставь свой ключ Gemini API в `GEMINI_API_KEY`
(получить бесплатный ключ: https://aistudio.google.com/apikey).

### Использование

1. Положи PDF-статьи в папку `pdfs/`.
2. Запусти:

   ```bash
   python scripts/analyze_new_sources.py
   ```

3. Открой `digest.xlsx` — там появятся новые строки по новым статьям.

Как только добавляешь в `pdfs/` новые файлы — просто запускаешь скрипт снова,
он обработает только их и дозапишет строки в тот же `digest.xlsx`.

Настройки (папка с PDF, имя итогового файла, модель) переопределяются через
переменные окружения в `.env` — см. `.env.example`.
