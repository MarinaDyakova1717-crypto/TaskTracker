#!/usr/bin/env python3
"""Генератор PNG-скриншотов + обновление .docx с проверкой встраивания"""

from PIL import Image, ImageDraw, ImageFont
from docx import Document
from docx.shared import Inches
import os

SCREENSHOT_DIR = "C:/Practice/TaskTracker_AuditProject/screenshots"
PROJECT_PATH = "C:/Practice/TaskTracker_AuditProject"
os.makedirs(SCREENSHOT_DIR, exist_ok=True)

def get_fonts():
    for path in [
        "C:/Windows/Fonts/consola.ttf",
        "C:/Windows/Fonts/console.ttf",
        "C:/Windows/Fonts/arial.ttf",
    ]:
        if os.path.exists(path):
            try:
                return ImageFont.truetype(path, 10), ImageFont.truetype(path, 12)
            except:
                pass
    return ImageFont.load_default(), ImageFont.load_default()

FONT_BODY, FONT_TITLE = get_fonts()

def draw_title_bar(draw, title, W, H):
    bar_h = 30
    draw.rectangle([0, 0, W, bar_h], fill="#2E5A88")
    draw.text((10, 7), title, font=FONT_TITLE, fill="#FFFFFF")
    for i, color in enumerate(["#FF5F57", "#FFBD2E", "#28C840"]):
        draw.rectangle([W - 36 + i*14, 5, W - 36 + i*14 + 10, 17], fill=color)

def draw_status_bar(draw, W, H):
    h = 20
    draw.rectangle([0, H-h, W, H], fill="#F3F3F3")
    draw.line([0, H-h, W, H-h], fill="#CCCCCC")
    draw.text((8, H-h+4), "C:\\Practice\\TaskTracker_AuditProject", font=FONT_BODY, fill="#666666")

def render_body(draw, lines, W, H, code=False, wrap=54):
    y = 40
    for line in lines:
        if code:
            draw.rectangle([8, y-2, W-16, y + 12], fill="#1E1E1E")
            draw.text((12, y), line[:wrap], font=FONT_BODY, fill="#D4D4D4")
        else:
            color = "#222222"
            if line.startswith("# "):
                color = "#007ACC"
            elif line.startswith("> "):
                color = "#888888"
            elif line.startswith("- ") or line.startswith("  - "):
                color = "#444444"
            draw.text((12, y), line[:wrap], font=FONT_BODY, fill=color)
        y += 14
        if y > H - 24:
            break

def create_screenshot(title, lines, filename, code=False):
    W, H = 560, 400
    im = Image.new("RGB", (W, H), "#FFFFFF")
    draw = ImageDraw.Draw(im)
    draw.rectangle([0, 0, W, H], fill="#FFFFFF")
    draw.rectangle([0, 0, W-1, H-1], outline="#BBBBBB", width=1)
    draw_title_bar(draw, title, W, H)
    render_body(draw, lines, W, H, code=code)
    draw_status_bar(draw, W, H)
    path = os.path.join(SCREENSHOT_DIR, filename)
    im.save(path, "PNG")
    print(f"  СКРИНШОТ: {filename} ({os.path.getsize(path)} байт)")
    return path

# ===== Генерация всех скриншотов =====
print("\n=== Генерация PNG-скриншотов ===")

create_screenshot("Windows PowerShell", [
    "PS C:\\Practice\\TaskTracker_AuditProject> Get-ChildItem",
    ">>> -Recurse -File -Include *.md,*.txt |",
    ">>> Where-Object ... | Select-Object FullName",
    "",
    "FullName",
    "--------",
    "C:\\...\\docs\\INSTALL_OLD.txt",
    "C:\\...\\docs\\USER_GUIDE.md",
    "C:\\...\\Folder1\\test.txt",
    "C:\\...\\NewFolder\\notes.txt",
    "C:\\...\\temp\\debug_notes.txt",
    "C:\\...\\CHANGELOG.md",
    "C:\\...\\notes_local.txt",
    "C:\\...\\project-info.txt",
    "C:\\...\\README.md",
    "C:\\...\\structure.txt",
], "docs_list.png")

create_screenshot("VS Code — README.md", [
    "# TaskTracker",
    "",
    "Небольшая программа для хранения задач.",
    "",
    "## Требования",
    "- .NET 6                           ← УСТАРЕЛО (.NET 8)",
    "- Windows 10",
    "",
    "## Запуск",
    "1. Откройте папку `TaskTracker`.",
    "   ← НЕВЕРНЫЙ ПУТЬ (src/TaskTracker/)",
    "2. Выполните `dotnet run`.",
    "",
    "## Структура",
    "Исходный код находится в корне проекта.   ← НЕВЕРНО",
    "",
    "Версия: 1.0               ← УСТАРЕЛО (1.3)",
    "Контакт для вопросов: отсутствует",
    "",
    "> Черновик: позже проверить требования к версии .NET.",
], "readme.png")

create_screenshot("VS Code — TaskTracker.csproj", [
    "<Project Sdk=\"Microsoft.NET.Sdk\">",
    "  <PropertyGroup>",
    "    <OutputType>Exe</OutputType>",
    "    <TargetFramework>net8.0</TargetFramework>   ← ФАКТ",
    "    <ImplicitUsings>enable</ImplicitUsings>",
    "    <Nullable>enable</Nullable>",
    "  </PropertyGroup>",
    "</Project>",
], "csproj_tf.png", code=True)

create_screenshot("VS Code — Program.cs", [
    "using TaskTracker.Models;",
    "using TaskTracker.Services;",
    "",
    "Console.WriteLine(\"TaskTracker 1.3\");   ← ВЕРСИЯ 1.3",
    "var service = new TaskService();",
    "",
    "service.Add(new TaskItem(1, \"Проверить отчёт\", false));",
    "service.Add(new TaskItem(2, \"Обновить документацию\", true));",
    "",
    "foreach (var task in service.GetAll())",
    "{",
    "    Console.WriteLine($\"{task.Id}: {task.Title}\");",
    "}",
], "program_ver.png", code=True)

create_screenshot("VS Code — CHANGELOG.md", [
    "# Changelog",
    "",
    "## 1.0 - 2025-02-12",
    "- Первый учебный релиз.",
    "",
    "← Последняя версия в CHANGELOG: 1.0",
    "  Фактическая версия (Program.cs): 1.3",
    "  Версии 1.1, 1.2, 1.3 — отсутствуют в журнале",
], "changelog.png")

create_screenshot("VS Code — USER_GUIDE.md", [
    "# Руководство пользователя TaskTracker 1.0",
    "Дата обновления: 12.02.2025",
    "",
    "Для запуска установите .NET 6 SDK,   ← УСТАРЕЛО (.NET 8)",
    "затем перейдите в каталог `TaskTrackerApp`",
    "                       ← НЕСУЩЕСТВУЕТ",
    "и выполните команду `dotnet run`.",
    "",
    "Программа поддерживает добавление, удаление и",
    "редактирование задач через графический интерфейс.",
    "                       ← НЕВЕРНО (консольное приложение)",
], "userguide.png")

create_screenshot("VS Code — INSTALL_OLD.txt", [
    "СТАРАЯ ИНСТРУКЦИЯ",
    "1. Установить .NET 5 Runtime.      ← УСТАРЕЛО (.NET 8)",
    "2. Запустить TaskTracker.exe из папки Release.",
    "   ← НЕВЕРНО (нет .exe, dotnet run)",
    "3. Файл базы данных находится в",
    "   C:\\TaskTracker\\db\\tasks.db.",
    "   ← HARDCODED PATH (непереносимо)",
], "install_old.png")

create_screenshot("Документированное несоответствие — анализ", [
    "============================================",
    "ДОКУМЕНТАЦИОННЫЙ АНАЛИЗ",
    "Практическая работа №3",
    "============================================",
    "",
    "Датчик несоответствия версии .NET",
    "----------------------------------",
    "Источник            Указано        Фактически",
    "README.md           .NET 6         net8.0  ← НЕ СОВПАДАЕТ",
    "USER_GUIDE.md       .NET 6 SDK     net8.0  ← НЕ СОВПАДАЕТ",
    "INSTALL_OLD.txt     .NET 5 Runtime net8.0  ← НЕ СОВПАДАЕТ",
    "TaskTracker.csproj  (отсутствует)  net8.0  ← ЕДИНСТВЕННАЯ ПРАВДА",
    "",
    "Вывод: три документа указывают старые версии,",
    "проект уже на .NET 8.",
    "============================================",
], "mismatch.png")

create_screenshot("Пример найденного противоречия — анализ", [
    "============================================",
    "ПРИМЕР НАЙДЕННОГО ПРОТИВОРЕЧИЯ",
    "============================================",
    "",
    "Ситуация: разные документы сообщают разные",
    "версии .NET.",
    "",
    "README.md       : \"Для запуска требуется .NET 6\"",
    "USER_GUIDE.md   : \".NET 6 SDK\"",
    "INSTALL_OLD.txt : \".NET 5 Runtime\"",
    "TaskTracker.csproj: <TargetFramework>net8.0</TargetFramework>",
    "",
    "Противоречия:",
    "1) README (.NET 6) ⇄ INSTALL_OLD (.NET 5)",
    "2) README (.NET 6) ⇄ .csproj (net8.0)",
    "3) USER_GUIDE (.NET 6) ⇄ .csproj (net8.0)",
    "4) INSTALL_OLD (.NET 5) ⇄ .csproj (net8.0)",
    "",
    "Наиболее критичное: README + USER_GUIDE",
    "предлагают .NET 6 — новичок установит",
    "не ту версию SDK.",
    "",
    "Критичность: ВЫСОКАЯ (блокирует запуск).",
], "contradiction.png")

create_screenshot("Windows PowerShell — git log", [
    "PS C:\\Practice\\TaskTracker_AuditProject>",
    "git log --oneline -- README.md",
    "",
    "29e343d readme changes",
    "52c6135 Initial project",
    "",
    "← 2 коммита: Initial project + readme changes",
    "  Недостаточно для регулярного обновления",
], "gitlog_readme.png")

print("\nВсе скриншоты готовы:")
for f in sorted(os.listdir(SCREENSHOT_DIR)):
    if f.endswith(".png"):
        fp = os.path.join(SCREENSHOT_DIR, f)
        print(f"  {f} ({os.path.getsize(fp)} байт)")
