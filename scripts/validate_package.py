#!/usr/bin/env python3
"""Проверить локальные инварианты пакета; не заменяет проверку каталога OpenAI."""

from __future__ import annotations

import ast
import json
from pathlib import Path
import re
import sys

import yaml


ROOT = Path(__file__).resolve().parents[1]


def require(condition: bool, message: str) -> None:
    if not condition:
        raise ValueError(message)


def read_json(relative: str) -> dict:
    path = ROOT / relative
    return json.loads(path.read_text(encoding="utf-8"))


def main() -> int:
    try:
        manifest = read_json("plugin.json")
        require(manifest.get("$schema") == "https://agent-plugins.org/schemas/1.0.0/plugin.schema.json", "Не указан поддерживаемый формат манифеста.")
        require(re.fullmatch(r"[a-z0-9]+(?:-[a-z0-9]+)*", manifest.get("name", "")) is not None, "Неверное техническое имя плагина.")
        require(re.fullmatch(r"\d+\.\d+\.\d+", manifest.get("version", "")) is not None, "Неверный номер версии.")
        interface = manifest["extensions"]["com.openai"]["interface"]
        for key in ("displayName", "shortDescription", "longDescription"):
            require(bool(re.search("[А-Яа-яЁё]", interface.get(key, ""))), f"Не заполнен русский текст интерфейса: {key}.")
        require(isinstance(interface.get("defaultPrompt"), list) and bool(interface["defaultPrompt"]), "Не заданы начальные запросы.")
        legacy = read_json(".codex-plugin/plugin.json")
        for key in ("name", "version"):
            require(legacy.get(key) == manifest[key], f"Манифесты расходятся по полю {key}.")
        require(legacy.get("skills") == "./skills/", "Неверный путь к навыкам в манифесте совместимости.")

        market = read_json(".agents/plugins/marketplace.json")
        require(len(market["plugins"]) == 1, "В этом пакете должен быть один плагин.")
        entry = market["plugins"][0]
        require(entry["name"] == manifest["name"], "Каталог ссылается на другое имя.")
        require(entry["source"]["source"] == "local", "Для каталога репозитория ожидается локальный путь.")
        relative = entry["source"]["path"]
        require(relative.startswith("./"), "Путь каталога должен начинаться с ./.")
        resolved = (ROOT / relative).resolve()
        require(resolved == ROOT, "Каталог должен указывать на корень этого плагина.")
        require(entry["policy"]["installation"] == "AVAILABLE", "Плагин должен быть доступен для установки.")
        require(entry["policy"]["authentication"] == "ON_INSTALL", "Не задана политика подключения каталога.")
        require(entry["category"] == "Productivity", "Не задана категория каталога.")

        skills = list((ROOT / "skills").glob("*/SKILL.md"))
        require(len(skills) == 1, "В пакете должен быть один основной навык.")
        skill = skills[0]
        content = skill.read_text(encoding="utf-8")
        match = re.match(r"\A---\n(.*?)\n---\n(.*)\Z", content, re.S)
        require(match is not None, "Не найдены метаданные навыка.")
        metadata = yaml.safe_load(match.group(1))
        require(set(metadata) == {"name", "description"}, "Метаданные навыка должны содержать name и description.")
        require(metadata["name"] == manifest["name"] == skill.parent.name, "Имена навыка и плагина не согласованы.")
        require(len(metadata["description"]) <= 1024, "Слишком длинное описание навыка.")
        require(bool(match.group(2).strip()), "Инструкция навыка пуста.")
        require("TODO" not in content, "В инструкции осталась незаполненная заготовка.")
        for relative in ("references/formats.md", "references/checks.md", "references/report.md", "scripts/extract_text.py", "scripts/build_report.py", "scripts/requirements.txt", "assets/report-example.json", "agents/openai.yaml"):
            require((skill.parent / relative).is_file(), f"Отсутствует ресурс навыка: {relative}.")
            if relative != "agents/openai.yaml":
                require(relative in content, f"Ресурс не связан с основной инструкцией: {relative}.")
        agent = yaml.safe_load((skill.parent / "agents/openai.yaml").read_text(encoding="utf-8"))
        require("$proverka-tekstov" in agent["interface"]["default_prompt"], "Начальный запрос должен ссылаться на навык.")
        short = agent["interface"]["short_description"]
        require(25 <= len(short) <= 64, "Краткое описание навыка должно содержать 25–64 символа.")

        for path in ROOT.rglob("*.py"):
            if not any(part in {".git", ".venv", "__pycache__"} for part in path.relative_to(ROOT).parts):
                ast.parse(path.read_text(encoding="utf-8"), filename=str(path.relative_to(ROOT)))
        for path in (ROOT / "tests").glob("test_*.py"):
            require("/root/.codex/" not in path.read_text(encoding="utf-8"), f"В тесте остался путь среды автора: {path.name}.")
        print("Пакет проверен: манифесты, каталог, инструкция, ресурсы и синтаксис Python согласованы.")
        print("Это локальная проверка структуры; установка в приложении проверяется отдельно.")
        return 0
    except (ValueError, KeyError, TypeError, OSError, SyntaxError, yaml.YAMLError) as error:
        print(f"Ошибка проверки пакета: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
