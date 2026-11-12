#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Обёртка над distribute.py: запускает скрипт навыка и пересказывает результат через GigaChat.

Запуск:
    export GIGACHAT_CREDENTIALS="<ключ>"
    python run_agent.py <файл1.xlsx> <файл2.xlsx> <файл3.xlsx>
"""
import os
import subprocess
import sys

from gigachat import GigaChat

# путь к distribute.py — эталонная копия из архива навыка
DIST = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "Агент_распределения_ФУЛ",
    "logist-distribution",
    "logist-distribution",
    "scripts",
    "distribute.py",
)

OUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "out_agent")

SYSTEM_PROMPT = """Ты помощник сотрудника, который распределяет заказы между логистами.
Тебе дают вывод скрипта распределения ФУЛ. Перескажи его коротко по шаблону:
1. Первая строка: сколько новых строк Cloud, повторных отправок и вопросов (числа из вывода скрипта).
2. Если есть «Требует внимания» — перечисли каждую позицию с причиной (номер ФУЛ + что не так).
3. Если есть возможные пропуски — укажи их количество и что они на отдельном листе.
4. В конце — путь к файлу результата.
Не добавляй числа, номера ФУЛ, имена логистов или причины, которых нет в выводе скрипта.
Если чего-то нет в выводе — не упоминай это."""


def main():
    if len(sys.argv) < 4:
        print("Использование: python run_agent.py <файл1.xlsx> <файл2.xlsx> <файл3.xlsx>")
        sys.exit(2)

    files = sys.argv[1:4]
    for f in files:
        if not os.path.isfile(f):
            print(f"ОШИБКА: файл не найден: {f}", file=sys.stderr)
            sys.exit(2)

    # 1. Запуск distribute.py
    cmd = [sys.executable, DIST, "--files", *files, "--out-dir", OUT_DIR]
    print(">>> Запуск скрипта навыка...")
    result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")
    if result.returncode != 0:
        print("ОШИБКА скрипта навыка:", file=sys.stderr)
        print(result.stderr, file=sys.stderr)
        sys.exit(2)

    print(">>> Скрипт навыка отработал. Вывод:")
    print(result.stdout)
    print(">>> Отправка в GigaChat для пересказа...\n")

    # 2. GigaChat пересказывает
    creds = os.environ.get("GIGACHAT_CREDENTIALS")
    if not creds:
        print("ОШИБКА: не задан GIGACHAT_CREDENTIALS", file=sys.stderr)
        sys.exit(2)

    giga = GigaChat(
        credentials=creds,
        model="GigaChat-2",  # Lite-модель, входит в Freemium [citation:3]
        scope="GIGACHAT_API_PERS",
        verify_ssl_certs=False,  # для Windows без сертификата НУЦ Минцифры [citation:1]
        timeout=60,
    )

    user_prompt = f"Вывод скрипта:\n\n{result.stdout}\n\nПерескажи по шаблону."
    response = giga.chat({
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": user_prompt},
        ]
    })

    # формат ответа в актуальном SDK: messages[0].content[0].text [citation:1]
    answer = response.choices[0].message.content
    print("=" * 60)
    print("ОТВЕТ АГЕНТА (GigaChat):")
    print("=" * 60)
    print(answer)
    print("=" * 60)


if __name__ == "__main__":
    main()