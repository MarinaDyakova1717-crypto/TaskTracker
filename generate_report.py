#!/usr/bin/env python3
"""Генератор отчёта по Практической работе №3 в формате .docx"""

from docx import Document
from docx.shared import Pt, Inches, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml
import os

PROJECT_PATH = "C:/Practice/TaskTracker_AuditProject"
SCREENSHOT_DIR = os.path.join(PROJECT_PATH, "screenshots")

def set_cell_shading(cell, color):
    """Установить фон ячейки"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shading_element = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color}"/>')
    tcPr.append(shading_element)

def add_table(doc, headers, rows, col_widths=None):
    """Добавить таблицу с заголовком"""
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # Заголовок
    for i, header in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = header
        for paragraph in cell.paragraphs:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in paragraph.runs:
                run.bold = True
                run.font.size = Pt(10)
        set_cell_shading(cell, "2E5A88")
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.color.rgb = RGBColor(255, 255, 255)
    
    # Данные
    for row_idx, row_data in enumerate(rows):
        for col_idx, cell_text in enumerate(row_data):
            cell = table.rows[row_idx + 1].cells[col_idx]
            cell.text = str(cell_text) if cell_text is not None else ""
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(10)
            if row_idx % 2 == 1:
                set_cell_shading(cell, "E8EEF4")
    
    if col_widths:
        for i, width in enumerate(col_widths):
            for row in table.rows:
                row.cells[i].width = Cm(width)
    
    return table

def add_heading(doc, text, level=1):
    h = doc.add_heading(text, level=level)
    for run in h.runs:
        run.font.color.rgb = RGBColor(0, 51, 102)
    return h

def add_para(doc, text, bold=False, italic=False, size=11):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = bold
    run.italic = italic
    run.font.size = Pt(size)
    return p

def add_bullet(doc, text, level=0):
    p = doc.add_paragraph(text, style='List Bullet')
    p.paragraph_format.left_indent = Cm(1.27 * (level + 1))
    for run in p.runs:
        run.font.size = Pt(11)
    return p

def add_image(doc, filepath, width_inches=5.5):
    """Добавить изображение с überwachen"""
    if os.path.exists(filepath):
        try:
            doc.add_picture(filepath, width=Inches(width_inches))
            last_paragraph = doc.paragraphs[-1]
            last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            # Добавить подпись
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(f"[Скриншот: {os.path.basename(filepath)}]")
            run.italic = True
            run.font.size = Pt(9)
            run.font.color.rgb = RGBColor(100, 100, 100)
        except Exception as e:
            add_para(doc, f"[Ошибка загрузки изображения {filepath}: {e}]", italic=True)
    else:
        add_para(doc, f"[Скриншот не найден: {filepath}]", italic=True)

def create_report():
    doc = Document()
    
    # ===== ШРИФТЫ =====
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Calibri'
    font.size = Pt(11)
    
    # ===== ТИТУЛЬНЫЙ ЛИСТ =====
    for _ in range(4):
        doc.add_paragraph()
    
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("ПРАКТИЧЕСКАЯ РАБОТА №3")
    run.bold = True
    run.font.size = Pt(24)
    run.font.color.rgb = RGBColor(0, 51, 102)
    
    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = subtitle.add_run("Анализ актуальности документации и информации\nпрограммного проекта")
    run.font.size = Pt(16)
    run.font.color.rgb = RGBColor(60, 60, 60)
    
    doc.add_paragraph()
    doc.add_paragraph()
    
    info_lines = [
        ("ФИО:", "___________________"),
        ("Группа:", "___________________"),
        ("Дата:", "___________________"),
        ("Номер работы:", "Практическая работа №3"),
        ("Тема:", "Анализ актуальности документации и информации программного проекта"),
        ("Курс:", "3 курс"),
        ("Исходный проект:", "TaskTracker_AuditProject"),
    ]
    
    for label, value in info_lines:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run_label = p.add_run(f"{label:20s}")
        run_label.bold = True
        run_label.font.size = Pt(12)
        run_value = p.add_run(value)
        run_value.font.size = Pt(12)
    
    doc.add_page_break()
    
    # ===== ЦЕЛЬ =====
    add_heading(doc, "Цель работы", 1)
    add_para(doc, "Научиться определять, насколько документация программного проекта соответствует его текущему состоянию, находить устаревшие, противоречивые, неполные и недостоверные сведения и оценивать влияние этих проблем на сопровождение программного продукта.")
    
    doc.add_page_break()
    
    # ===== 1. ИНВЕНТАРИЗАЦИЯ ДОКУМЕНТАЦИИ =====
    add_heading(doc, "1. Инвентаризация документации", 1)
    add_para(doc, "Смотри также: скриншот «Список документов» (Рис. 1).", italic=True)
    
    # Скриншот списка документов
    add_image(doc, os.path.join(SCREENSHOT_DIR, "docs_list.png"), 5.0)
    doc.add_paragraph()
    
    headers_1 = ["№", "Документ", "Расположение", "Формат", "Предполагаемое назначение", "Категория"]
    rows_1 = [
        ["1", "README.md", "Корень проекта", "md", "Общая информация о проекте", "GENERAL"],
        ["2", "CHANGELOG.md", "Корень проекта", "md", "История изменений версий", "CHANGE"],
        ["3", "USER_GUIDE.md", "docs/", "md", "Руководство пользователя", "USER"],
        ["4", "INSTALL_OLD.txt", "docs/", "txt", "Старая инструкция установки (.NET 5)", "OLD"],
        ["5", "project-info.txt", "Корень проекта", "txt", "Служебная информация (нет даты проверки)", "LOCAL"],
        ["6", "notes_local.txt", "Корень проекта", "txt", "Личные заметки разработчика (не в Git)", "LOCAL"],
        ["7", "temp/debug_notes.txt", "temp/", "txt", "Локальные отладочные заметки (не в Git)", "LOCAL"],
        ["8", "Folder1/test.txt", "Folder1/", "txt", "Тестовый файл (вопрос о назначении)", "OTHER"],
        ["9", "NewFolder/notes.txt", "NewFolder/", "txt", "TODO: перенести документацию", "OTHER"],
        ["10", "structure.txt", "Корень проекта", "bin", "Описание структуры (бинарный файл)", "OTHER"],
    ]
    add_table(doc, headers_1, rows_1, [0.8, 2.5, 2.5, 1.2, 4.0, 1.5])
    
    doc.add_paragraph()
    add_para(doc, "Примечание: docs/INSTALL.md отсутствует — современная инструкция установки не создана. Вместо неё присутствует только устаревший INSTALL_OLD.txt.", bold=False, italic=True)
    
    doc.add_page_break()
    
    # ===== 2. АНАЛИЗ README =====
    add_heading(doc, "2. Анализ README", 1)
    add_para(doc, "Смотри также: скриншоты «README», «TargetFramework в .csproj», «Версия в Program.cs» (Рис. 2, 3, 4).", italic=True)
    
    add_image(doc, os.path.join(SCREENSHOT_DIR, "readme.png"), 4.5)
    doc.add_paragraph()
    add_image(doc, os.path.join(SCREENSHOT_DIR, "csproj_tf.png"), 4.5)
    doc.add_paragraph()
    add_image(doc, os.path.join(SCREENSHOT_DIR, "program_ver.png"), 4.5)
    doc.add_paragraph()
    
    # Таблица 2 — Проверка версии платформы
    add_heading(doc, "Таблица 2 — Проверка версии платформы (.NET)", 2)
    headers_2 = ["Источник", "Значение", "Совпадает"]
    rows_2 = [
        ["README.md", ".NET 6", "нет"],
        ["TaskTracker.csproj", "net8.0", "— (фактически)"],
    ]
    add_table(doc, headers_2, rows_2, [3.5, 3.5, 2.5])
    
    doc.add_paragraph()
    
    # Таблица 3 — Проверка версии приложения
    add_heading(doc, "Таблица 3 — Проверка версии приложения", 2)
    headers_3 = ["Источник", "Версия", "Совпадает"]
    rows_3 = [
        ["README.md", "1.0", "нет"],
        ["Program.cs", "1.3", "— (фактически)"],
        ["CHANGELOG.md", "1.0", "нет (нет записи про 1.1-1.3)"],
    ]
    add_table(doc, headers_3, rows_3, [3.5, 3.5, 3.5])
    
    doc.add_paragraph()
    
    # Таблица 4 — Аудит README
    add_heading(doc, "Таблица 4 — Аудит README", 2)
    add_para(doc, "Статусы: АКТУАЛЬНО / НЕАКТУАЛЬНО / НЕПОЛНО / НЕВОЗМОЖНО ПРОВЕРИТЬ", italic=True, size=10)
    
    headers_4 = ["Проверяемый элемент", "Указано в README", "Фактически", "Статус", "Комментарий"]
    rows_4 = [
        ["Название проекта", "TaskTracker", "TaskTracker", "АКТУАЛЬНО", "Совпадает"],
        [".NET", ".NET 6", "net8.0 (csproj)", "НЕАКТУАЛЬНО", "Указано .NET 6, проект использует .NET 8. Критично для новичка."],
        ["Версия приложения", "1.0", "1.3 (Program.cs)", "НЕАКТУАЛЬНО", "README устарел на 2 версии."],
        ["Исходный код", "В корне проекта", "src/TaskTracker/", "НЕАКТУАЛЬНО", "README неверно описывает расположение исходного кода."],
        ["Команда запуска", "dotnet run", "dotnet run (в каталоге src/TaskTracker)", "НЕПОЛНО", "Указана команда, но не указан правильный каталог."],
        ["Каталог запуска", "папка `TaskTracker`", "src/TaskTracker/", "НЕАКТУАЛЬНО", "README указывает несуществующий путь к проекту."],
        ["Описание функций", "хранение задач", "консольное приложение: добавление, вывод задач", "НЕПОЛНО", "Нет описания доступных операций."],
        ["Контакт", "отсутствует", "отсутствует", "НЕПОЛНО", "Нет контакта для вопросов."],
    ]
    add_table(doc, headers_4, rows_4, [2.0, 2.5, 2.5, 2.0, 3.5])
    
    doc.add_page_break()
    
    # ===== 3. АНАЛИЗ USER_GUIDE =====
    add_heading(doc, "3. Анализ USER_GUIDE", 1)
    add_para(doc, "Смотри также: скриншоты «USER_GUIDE», «Пример неверного пути» (Рис. 5, 6).", italic=True)
    
    add_image(doc, os.path.join(SCREENSHOT_DIR, "userguide.png"), 4.5)
    doc.add_paragraph()
    add_image(doc, os.path.join(SCREENSHOT_DIR, "wrong_path.png"), 4.5)
    doc.add_paragraph()
    
    # Таблица 5 — Функциональная актуальность USER_GUIDE
    add_heading(doc, "Таблица 5 — Функциональная актуальность USER_GUIDE", 2)
    headers_5 = ["Описанная функция", "Есть в текущем проекте", "Совпадает описание", "Комментарий"]
    rows_5 = [
        ["Программа с GUI", "Нет (консольное приложение)", "НЕ СОВПАДАЕТ", "USER_GUIDE описывает графический интерфейс, но проект — консольный. Критическая ошибка."],
        ["Добавление задач", "Да (TaskService.Add)", "СОВПАДАЕТ", "Функция реализована."],
        ["Удаление задач", "Нет в текущей реализации", "НЕТ ФУНКЦИИ", "в Program.cs нет удаления. В USER_GUIDE упоминается — нереализовано."],
        ["Редактирование задач", "Нет в текущей реализации", "НЕТ ФУНКЦИИ", "в Program.cs нет редактирования. В USER_GUIDE упоминается."],
        ["Переход в TaskTrackerApp", "Каталог не существует", "НЕСУЩЕСТВУЕТ", "USER_GUIDE указывает несуществующий каталог."],
        ["Установка .NET 6 SDK", ".NET 8 требуется", "НЕАКТУАЛЬНО", "Указана неправильная версия SDK."],
    ]
    add_table(doc, headers_5, rows_5, [2.5, 2.5, 2.5, 4.0])
    
    doc.add_paragraph()
    add_para(doc, "Проверка пути TaskTrackerApp:", bold=True)
    add_bullet(doc, "Команда: Get-ChildItem -Recurse -Directory | Where-Object { $_.Name -eq 'TaskTrackerApp' }")
    add_bullet(doc, "Результат: пустой вывод (каталог не найден)")
    add_bullet(doc, "Вывод: USER_GUIDE содержит указание на несуществующий каталог — доказательством расхождения.")
    
    doc.add_page_break()
    
    # ===== 4. АНАЛИЗ ИНСТРУКЦИИ УСТАНОВКИ =====
    add_heading(doc, "4. Анализ инструкции установки", 1)
    add_para(doc, "Смотри также: скриншоты «Версия .NET в .csproj», «Пример неверного пути» (Рис. 3, 6).", italic=True)
    
    add_para(doc, "Ситуация: docs/INSTALL.md отсутствует. Единственный файл установки — docs/INSTALL_OLD.txt, который является устаревшей инструкцией для .NET 5.", bold=True)
    
    add_image(doc, os.path.join(SCREENSHOT_DIR, "install_old.png"), 4.5)
    doc.add_paragraph()
    
    # Таблица 6 — Аудит установки
    add_heading(doc, "Таблица 6 — Аудит установки", 2)
    headers_6 = ["Шаг документа", "Фактически выполним", "Проблема", "Критичность"]
    rows_6 = [
        ["1. Установить .NET 5 Runtime", "Нет — проект требует .NET 8 SDK", "Указана устаревшая версия. Пользователь установит неподходящую среду.", "ВЫСОКАЯ"],
        ["2. Запустить TaskTracker.exe из папки Release", "Нет — проект собирается через dotnet run (Executable), .exe не предоставляется", "Инструкция предназначена для старой версии, которая не соответствует текущей.", "ВЫСОКАЯ"],
        ["3. База данных: C:\\TaskTracker\\db\\tasks.db", "Нет — абсолютный путь C:\\TaskTracker\\ привязан к конкретной машине. В проекте база данных хранится в data/tasks.json.", "Hardcoded path. Непереносимая инструкция. Привязка к диску C: и пользователю.", "СРЕДНЯЯ"],
        ["Отсутствие INSTALL.md", "—", "Нет современной инструкции установки. Новому разработчику негде взять актуальные сведения.", "ВЫСОКАЯ"],
    ]
    add_table(doc, headers_6, rows_6, [3.0, 3.0, 4.0, 1.5])
    
    doc.add_page_break()
    
    # ===== 5. АНАЛИЗ CHANGELOG =====
    add_heading(doc, "5. Анализ CHANGELOG", 1)
    add_para(doc, "Смотри также: скриншоты «CHANGELOG», «Git log README» (Рис. 7, 10).", italic=True)
    
    add_image(doc, os.path.join(SCREENSHOT_DIR, "changelog.png"), 4.5)
    doc.add_paragraph()
    add_image(doc, os.path.join(SCREENSHOT_DIR, "gitlog_readme.png"), 4.5)
    doc.add_paragraph()
    
    # Таблица 7 — Анализ CHANGELOG
    add_heading(doc, "Таблица 7 — Анализ CHANGELOG", 2)
    headers_7 = ["Показатель", "Результат"]
    rows_7 = [
        ["Последняя версия в CHANGELOG", "1.0 (от 2025-02-12)"],
        ["Версия программы (Program.cs)", "1.3"],
        ["Совпадает", "НЕТ — CHANGELOG не обновлялся с момента 1.0"],
        ["Есть описание промежуточных версий (1.1, 1.2, 1.3)", "НЕТ — отсутствует вся история после 1.0"],
        ["Общая актуальность", "НЕАКТУАЛЬНО — CHANGELOG не отражает текущее состояние проекта"],
    ]
    add_table(doc, headers_7, rows_7, [4.5, 4.5])
    
    doc.add_page_break()
    
    # ===== 6. МАТРИЦА ПРОТИВОРЕЧИЙ =====
    add_heading(doc, "6. Матрица противоречий", 1)
    add_para(doc, "Смотри также: скриншот «Пример найденного противоречия» (Рис. 8).", italic=True)
    
    # Скриншот противоречия
    add_image(doc, os.path.join(SCREENSHOT_DIR, "contradiction.png"), 4.5)
    doc.add_paragraph()
    
    # Таблица 8 — Матрица противоречий
    add_heading(doc, "Таблица 8 — Матрица противоречий", 2)
    headers_8 = ["Информация", "README", "USER_GUIDE", "INSTALL_OLD", "CHANGELOG", "Фактически", "Статус"]
    rows_8 = [
        [".NET", ".NET 6", ".NET 6 SDK", ".NET 5 Runtime", "отсутствует", "net8.0", "ЕСТЬ ПРОТИВОРЕЧИЕ (3 документа расходятся с фактом)"],
        ["Версия программы", "1.0", "1.0", "не указана", "1.0", "1.3", "ЕСТЬ ПРОТИВОРЕЧИЕ (README, USER_GUIDE, CHANGELOG показывают 1.0, факт — 1.3)"],
        ["Каталог проекта", "TaskTracker (корень)", "TaskTrackerApp", "C:\\TaskTracker\\ (старый)", "отсутствует", "src/TaskTracker/", "ЕСТЬ ПРОТИВОРЕЧИЕ / НЕСУЩЕСТВУЮЩИЙ ПУТЬ"],
        ["Тип приложения", "консольное (подразумевается)", "GUI (графический интерфейс)", "exe-файл", "не указано", "консольное приложение (Console.WriteLine)", "ЕСТЬ ПРОТИВОРЕЧИЕ (USER_GUIDE ≠ факт)"],
        ["Способ запуска", "dotnet run из TaskTracker/", "dotnet run из TaskTrackerApp/", "Запустить TaskTracker.exe из Release", "отсутствует", "dotnet run из src/TaskTracker/", "ЕСТЬ ПРОТИВОРЕЧИЕ (неправильные пути)"],
    ]
    add_table(doc, headers_8, rows_8, [1.8, 1.8, 1.8, 1.8, 1.5, 2.0, 1.8])
    
    doc.add_paragraph()
    add_para(doc, "Наиболее критичное противоречие:", bold=True)
    add_bullet(doc, "README требует .NET 6, USER_GUIDE требует .NET 6 SDK, INSTALL_OLD требует .NET 5 Runtime — а проект использует .NET 8.")
    add_bullet(doc, "Новичок, следуя документации, установит неподходящую версию .NET и не сможет запустить проект.")
    add_bullet(doc, "Проблема важна, потому что она блокирует начало работы с проектом.")
    
    doc.add_page_break()
    
    # ===== 7. ИСТОРИЯ ОБНОВЛЕНИЯ ДОКУМЕНТАЦИИ =====
    add_heading(doc, "7. История обновления документации", 1)
    add_para(doc, "Смотри также: скриншоты «Git log README», «Git log CHANGELOG», «Git log USER_GUIDE» (Рис. 9, 10, 11).", italic=True)
    
    add_image(doc, os.path.join(SCREENSHOT_DIR, "gitlog_readme.png"), 3.5)
    doc.add_paragraph()
    add_image(doc, os.path.join(SCREENSHOT_DIR, "gitlog_changelog.png"), 3.5)
    doc.add_paragraph()
    add_image(doc, os.path.join(SCREENSHOT_DIR, "gitlog_userguide.png"), 3.5)
    doc.add_paragraph()
    
    # Таблица 9 — История обновления
    add_heading(doc, "Таблица 9 — История обновления документов", 2)
    headers_9 = ["Документ", "Количество изменений", "Последнее сообщение", "Есть признаки регулярного обновления"]
    rows_9 = [
        ["README.md", "2", "readme changes (29e343d)", "НЕТ — только 2 коммита, один из них initial."],
        ["CHANGELOG.md", "1", "update (0d5287c)", "НЕТ — единственный коммит, дата 1.0 без обновлений."],
        ["USER_GUIDE.md", "1", "docs (c836114)", "НЕТ — единственный коммит, дата 12.02.2025 без изменений."],
    ]
    add_table(doc, headers_9, rows_9, [2.5, 2.5, 3.5, 3.5])
    
    doc.add_paragraph()
    add_para(doc, "Важно: не делайте выводы только на основании количества коммитов. Однако здесь количество коммитов минимально (1-2), и они не отражают текущее состояние проекта — это подтверждает нерегулярность обновлений.", italic=True)
    
    doc.add_page_break()
    
    # ===== 8. АНАЛИЗ ПУТЕЙ =====
    add_heading(doc, "8. Анализ путей", 1)
    add_para(doc, "Смотри также: скриншоты «Пример неверного пути», «Пример найденного противоречия» (Рис. 6, 8).", italic=True)
    
    add_image(doc, os.path.join(SCREENSHOT_DIR, "wrong_path.png"), 4.0)
    doc.add_paragraph()
    
    add_para(doc, "Поиск hardcoded path (C:\\):", bold=True)
    add_bullet(doc, "Команда: Get-ChildItem -Recurse -File -Include *.md,*.txt | Select-String -Pattern 'C:\\\\'")
    add_bullet(doc, "Результат: найден в docs/INSTALL_OLD.txt — строка 3: 'Файл базы данных находится в C:\\TaskTracker\\db\\tasks.db.'")
    add_bullet(doc, "Анализ: путь привязан к диску C: и не существует в общем виде. Непереносим на другие машины. Критичность: СРЕДНЯЯ.")
    
    doc.add_paragraph()
    add_para(doc, "Таблица найденных путей:", bold=True)
    headers_paths = ["Путь", "Где найден", "Тип", "Проблема", "Критичность"]
    rows_paths = [
        ["C:\\TaskTracker\\db\\tasks.db", "INSTALL_OLD.txt", "Hardcoded (abсолютный)", "Привязан к конкретной машине; непереносимый", "СРЕДНЯЯ"],
        ["TaskTrackerApp", "USER_GUIDE.md", "Относительный (неверный)", "Каталог не существует в проекте", "ВЫСОКАЯ"],
        ["TaskTracker/ (как корень)", "README.md", "Относительный (неверный)", "Исходный код находится в src/TaskTracker/, не в корне", "СРЕДНЯЯ"],
        ["папка Release (для exe)", "INSTALL_OLD.txt", "Относительный", "Текущий проект не предоставляет .exe; запускается через dotnet run", "СРЕДНЯЯ"],
    ]
    add_table(doc, headers_paths, rows_paths, [2.5, 2.0, 2.0, 3.5, 1.5])
    
    doc.add_page_break()
    
    # ===== 9. АНАЛИЗ ВЕРСИЙ =====
    add_heading(doc, "9. Анализ версий", 1)
    add_para(doc, "Смотри также: скриншоты «Версия в Program.cs», «TargetFramework в .csproj», «Пример несоответствия версии» (Рис. 3, 4, 2).", italic=True)
    
    add_image(doc, os.path.join(SCREENSHOT_DIR, "mismatch.png"), 4.0)
    doc.add_paragraph()
    
    add_para(doc, "Таблица упоминаний платформы (.NET):", bold=True)
    headers_ver = ["Документ", "Найденный текст", "Фактическая версия", "Актуально"]
    rows_ver = [
        ["README.md", ".NET 6", "net8.0", "НЕАКТУАЛЬНО"],
        ["USER_GUIDE.md", ".NET 6 SDK", "net8.0", "НЕАКТУАЛЬНО"],
        ["INSTALL_OLD.txt", ".NET 5 Runtime", "net8.0", "НЕАКТУАЛЬНО (устаревший документ)"],
    ]
    add_table(doc, headers_ver, rows_ver, [2.5, 2.5, 2.5, 2.5])
    
    doc.add_paragraph()
    add_para(doc, "Версионный анализ:", bold=True)
    add_bullet(doc, "Платформа: README (.NET 6) ≠ .csproj (net8.0) — рассинхронизация.")
    add_bullet(doc, "Версия приложения: README (1.0) ≠ Program.cs (1.3) — документ не обновлялся.")
    add_bullet(doc, "CHANGELOG: последняя запись 1.0, но проект на 1.3 — отсутствуют записи о версиях 1.1, 1.2, 1.3.")
    add_bullet(doc, "Semantic versioning: проект использует MAJOR.MINOR.PATCH (1.3 = major 1, minor 3, patch 0). CHANGELOG должен отражать все минорные изменения.")
    
    doc.add_page_break()
    
    # ===== 10. ОЦЕНКА ПОЛНОТЫ =====
    add_heading(doc, "10. Оценка полноты документации", 1)
    add_para(doc, "Таблица 11 — Полнота документации", bold=True)
    
    headers_full = ["Необходимая информация", "Присутствует", "Где находится", "Достаточно подробно"]
    rows_full = [
        ["Назначение проекта", "Да (частично)", "README.md — 'Небольшая программа для хранения задач'", "НЕДОСТАТОЧНО — нет описания функций, ограничений"],
        ["Требования", "Частично", "README.md — .NET 6, Windows 10", "НЕАКТУАЛЬНО — требования устарели"],
        ["Установка", "Нет", "docs/INSTALL_OLD.txt (устарела), INSTALL.md отсутствует", "НЕТ — нет актуальной инструкции"],
        ["Запуск", "Частично", "README.md — dotnet run из TaskTracker/", "НЕПОЛНО — неверный каталог, не указан .NET 8"],
        ["Структура", "Неверно", "README.md — 'в корне проекта'", "НЕАКТУАЛЬНО — структура изменена"],
        ["Возможности", "Нет", "USER_GUIDE.md — но неверно (GUI вместо консоли)", "НЕТ — не соответствует реальности"],
        ["Версии", "Частично", "README.md (1.0), CHANGELOG.md (1.0)", "НЕПОЛНО — нет записи о 1.1-1.3"],
        ["Известные ограничения", "Нет", "отсутствует", "НЕТ — нет раздела об ограничениях"],
    ]
    add_table(doc, headers_full, rows_full, [2.5, 1.5, 3.5, 2.0])
    
    doc.add_page_break()
    
    # ===== 11. ПРОВЕРКА ГЛАЗАМИ НОВОГО РАЗРАБОТЧИКА =====
    add_heading(doc, "11. Проверка документации глазами нового разработчика", 1)
    add_para(doc, "Практический сценарий: новый разработчик получает папку проекта и должен самостоятельно разобраться, что это, как установить, куда перейти, что запустить и что должно произойти.")
    
    add_para(doc, "Маршрут нового разработчика (реальный):", bold=True)
    add_bullet(doc, "1. Открыть проект — читает README.md")
    add_bullet(doc, "2. Читает требования: .NET 6, Windows 10")
    add_bullet(doc, "3. Ищет исходный код — README пишет 'в корне' — не находит (код в src/TaskTracker/)")
    add_bullet(doc, "4. Ищет инструкцию установки — находит docs/, но INSTALL.md нет, INSTALL_OLD.txt устарел")
    add_bullet(doc, "5. Читает USER_GUIDE — просит перейти в TaskTrackerApp — каталога нет")
    add_bullet(doc, "6. Читает CHANGELOG — последняя версия 1.0, но код говорит 1.3 — непонятно, что произошло")
    
    doc.add_paragraph()
    add_para(doc, "Таблица 12 — Проблемы знакомства с проектом", bold=True)
    headers_newdev = ["Этап", "Что должен сделать разработчик", "Какая информация мешает", "Результат"]
    rows_newdev = [
        ["1. Понимание проекта", "Прочитать README, понять, что это", "README: 'черновик', нет контакта, нет полного описания", "Новичок не понимает, что за проект"],
        ["2. Установка среды", "Установить .NET", "README: .NET 6; USER_GUIDE: .NET 6; INSTALL_OLD: .NET 5 — все неверно (нужен .NET 8)", "Новичок установит неверную версию SDK"],
        ["3. Поиск исходного кода", "Найти, где код", "README: 'в корне' — неверно (код в src/TaskTracker/)", "Новичок не найдет код"],
        ["4. Запуск", "Запустить проект", "USER_GUIDE: TaskTrackerApp — каталог не существует; README: TaskTracker/ — неверный путь", "Новичок не сможет запустить"],
        ["5. Понимание возможностей", "Понять, что программа умеет", "USER_GUIDE: GUI — неверно (консоль); CHANGELOG: только 1.0 — не видно развития", "Новичок получит неверное представление о функциональности"],
    ]
    add_table(doc, headers_newdev, rows_newdev, [1.8, 2.5, 3.5, 2.5])
    
    doc.add_page_break()
    
    # ===== 12. РЕЕСТР ПРОБЛЕМ =====
    add_heading(doc, "12. Реестр проблем документации", 1)
    add_para(doc, "Минимум 10 обоснованных проблем. Участвуют категории: VERSION, PATH, FUNCTION, MISSING, CONFLICT, OUTDATED, UNCLEAR, SECURITY.", italic=True)
    
    # Таблица 13 — Реестр проблем
    add_heading(doc, "Таблица 13 — Реестр проблем документации", 2)
    headers_reg = ["№", "Документ", "Категория", "Проблема", "Доказательство", "Последствие", "Критичность"]
    rows_reg = [
        ["1", "README.md", "VERSION", "Указана платформа .NET 6 вместо .NET 8", "README: '.NET 6'; csproj: '<TargetFramework>net8.0</TargetFramework>'", "Новый разработчик установит неподходящую версию SDK и не сможет запустить проект", "ВЫСОКАЯ"],
        ["2", "README.md", "VERSION", "Версия приложения 1.0 вместо 1.3", "README: 'Версия: 1.0'; Program.cs: 'TaskTracker 1.3'", "Неправильное представление о версии проекта; CHANGELOG не отражает изменений", "СРЕДНЯЯ"],
        ["3", "README.md", "PATH", "Неверное расположение исходного кода", "README: 'Исходный код находится в корне проекта'; факт: код в src/TaskTracker/", "Разработчик не найдет исходный код по инструкции", "СРЕДНЯЯ"],
        ["4", "README.md", "PATH", "Неверный каталог запуска", "README: 'Откройте папку TaskTracker'; фактический каталог: src/TaskTracker/", "Команда dotnet run не выполнится из указанного каталога", "СРЕДНЯЯ"],
        ["5", "USER_GUIDE.md", "VERSION", "Указана .NET 6 SDK вместо .NET 8", "USER_GUIDE: '.NET 6 SDK'; csproj: 'net8.0'", "Неправильная инструкция для установки SDK", "ВЫСОКАЯ"],
        ["6", "USER_GUIDE.md", "FUNCTION", "Описан графический интерфейс, но проект консольный", "USER_GUIDE: 'через графический интерфейс'; Program.cs: Console приложение (нет UI)", "Новичок ожидает GUI, не найдет его — непонимание того, что программа делает", "ВЫСОКАЯ"],
        ["7", "USER_GUIDE.md", "PATH", "Указан несуществующий каталог TaskTrackerApp", "USER_GUIDE: 'перейдите в каталог TaskTrackerApp'; Get-ChildItem -Recurse -Directory | where Name -eq 'TaskTrackerApp' — пустой вывод", "Команда не выполнима — каталога не существует", "ВЫСОКАЯ"],
        ["8", "INSTALL_OLD.txt", "VERSION", "Указана .NET 5 Runtime вместо .NET 8 SDK", "INSTALL_OLD.txt: '.NET 5 Runtime'; csproj: 'net8.0'", "Установка неподходящей среды; проект не запустится", "ВЫСОКАЯ"],
        ["9", "INSTALL_OLD.txt", "SECURITY", "Hardcoded path: C:\\TaskTracker\\db\\tasks.db", "INSTALL_OLD.txt строка 3: 'Файл базы данных находится в C:\\TaskTracker\\db\\tasks.db.'", "Непереносимая инструкция; путь не существует на другой машине", "СРЕДНЯЯ"],
        ["10", "INSTALL_OLD.txt", "OUTDATED", "Инструкция устарела (exe, Release, .NET 5)", "INSTALL_OLD.txt описывает старую версию; текущий проект — dotnet run, net8.0", "Инструкция бесполезна для текущей версии", "СРЕДНЯЯ"],
        ["11", "Все документы (отсутствие)", "MISSING", "Отсутствует современная инструкция установки (INSTALL.md)", "docs/INSTALL.md — файл не найден; только INSTALL_OLD.txt (устарел)", "Новому разработчику негде получить актуальные сведения об установке", "ВЫСОКАЯ"],
        ["12", "CHANGELOG.md", "OUTDATED", "CHANGELOG не обновлялся с 1.0, нет записей о 1.1-1.3", "CHANGELOG: последняя запись '1.0 - 2025-02-12'; Program.cs: '1.3'", "Невозможно понять, что изменилось между версиями", "СРЕДНЯЯ"],
        ["13", "README.md / USER_GUIDE.md / CHANGELOG.md", "OUTDATED", "Все документы показывают версию 1.0, но проект на 1.3", "README: 1.0; USER_GUIDE: 1.0; CHANGELOG: 1.0; Program.cs: 1.3", "Рассинхронизация версий во всех документах", "СРЕДНЯЯ"],
        ["14", "project-info.txt", "UNCLEAR", "Отсутствует дата последней проверки", "project-info.txt: 'Последняя проверка: неизвестно.'", "Невозможно оценить актуальность информации", "НИЗКАЯ"],
        ["15", "README.md", "UNCLEAR", "Раздел 'черновик' без даты", "README: '> Черновик: позже проверить требования к версии .NET.' — без даты и статуса", "Непонятно, когда планировалось обновление", "НИЗКАЯ"],
    ]
    add_table(doc, headers_reg, rows_reg, [0.5, 2.0, 1.2, 2.5, 2.5, 2.5, 1.2])
    
    doc.add_page_break()
    
    # ===== 13. ОЦЕНКА КАЧЕСТВА =====
    add_heading(doc, "13. Оценка качества документации", 1)
    add_para(doc, "Таблица 14 — Оценка документации (шкала: 5 — отлично, 1 — плохо)", italic=True)
    
    headers_qual = ["Критерий", "Оценка (1-5)", "Обоснование"]
    rows_qual = [
        ["Актуальность", "1", "Версии .NET неверны (6→8), версия приложения устарела (1.0→1.3), пути неверны (TaskTrackerApp не существует), CHANGELOG не обновлен."],
        ["Достоверность", "1", "USER_GUIDE описывает GUI, хотя проект консольный. README утверждает код в корне — неправда. Четкое расхождение с реальностью."],
        ["Полнота", "2", "Отсутствует INSTALL.md, нет описания ограничений, нет контакта, CHANGELOG обрывается на 1.0, нет описания функций в README."],
        ["Непротиворечивость", "1", "README (.NET 6) ⇄ USER_GUIDE (.NET 6) ⇄ INSTALL_OLD (.NET 5) ⇄ .csproj (net8.0) — три документа противоречат факту. Версии программы расходятся."],
        ["Понятность", "2", "README имеет пометку 'черновик', PROJECT-INFO — 'неизвестно'. Инструкции содержат несуществующие пути — запутать новичка легко."],
        ["Удобство запуска по инструкции", "1", "README: 'TaskTracker/' — неверный путь. USER_GUIDE: 'TaskTrackerApp' — несуществует. INSTALL_OLD: exe из Release — файл не предоставляется. Ни одна инструкция не приводит к запуску."],
        ["Описание структуры", "1", "README: 'в корне' — неверно. Правильная структура: src/TaskTracker/ (код), docs/ (документация), data/ (данные)."],
        ["Описание функций", "1", "USER_GUIDE: GUI + удаление + редактирование — не реализовано. README: нет описания функций. Программа: консольное приложение с добавлением и выводом задач."],
        ["История изменений", "1", "CHANGELOG содержит только 1.0, проект на 1.3. Git-логи показывают минимальную активность (1-2 коммита на документ)."],
        ["Общая готовность документации", "1", "Документация не позволяет правильно разобраться в проекте. Блокирует запуск, содержит противоречия, отсутствует современная установка."],
    ]
    add_table(doc, headers_qual, rows_qual, [2.2, 1.0, 5.0])
    
    doc.add_paragraph()
    add_para(doc, "Средняя оценка: 1.1 (по шкале 1-5). Документация практически не позволяет правильно разобраться в проекте.", bold=True)
    
    doc.add_page_break()
    
    # ===== 14. ПРЕДЛОЖЕНИЯ ПО ОБНОВЛЕНИЮ =====
    add_heading(doc, "14. Предложения по обновлению документации", 1)
    add_para(doc, "Таблица 15 — Предложения по обновлению", italic=True)
    add_para(doc, "Важно: документы пока не изменяем, только спроектируем изменения.", bold=True)
    
    headers_prop = ["№", "Что необходимо изменить", "Документ", "Причина", "Ожидаемый результат", "Приоритет"]
    rows_prop = [
        ["1", "Обновить требования к .NET: заменить .NET 6 на .NET 8", "README.md", "README содержит .NET 6; проект использует .NET 8", "Новый разработчик установит правильную версию SDK", "P1"],
        ["2", "Обновить версию приложения: заменить 1.0 на 1.3", "README.md", "README показывает 1.0; Program.cs — 1.3", "Версия в документации совпадет с фактической", "P1"],
        ["3", "Исправить описание расположения исходного кода", "README.md", "README: 'в корне'; факт: src/TaskTracker/", "Разработчик найдет код по инструкции", "P1"],
        ["4", "Исправить каталог запуска: вместо 'TaskTracker/' указать 'src/TaskTracker/'", "README.md", "README указывает неверный путь; команда dotnet run должна выполняться в src/TaskTracker/", "Инструкция запуска будет работать", "P1"],
        ["5", "Заменить .NET 6 SDK на .NET 8 SDK", "USER_GUIDE.md", "USER_GUIDE требует .NET 6; проект — .NET 8", "Установка SDK будет успешной", "P1"],
        ["6", "Удалить / заменить описание GUI на описание консольного приложения", "USER_GUIDE.md", "USER_GUIDE описывает GUI, но проект консольный. Нет кнопок, экранов, интерфейса.", "Новичок получит правильное представление о типе приложения", "P1"],
        ["7", "Убрать упоминание каталога TaskTrackerApp (не существует) и заменить на src/TaskTracker/", "USER_GUIDE.md", "USER_GUIDE: 'перейдите в TaskTrackerApp' — каталога нет; фактический: src/TaskTracker/", "Инструкция будет выполнима", "P1"],
        ["8", "Создать новый INSTALL.md (или переименовать INSTALL_OLD.txt в историю и создать актуальный)", "docs/", "INSTALL.md отсутствует; INSTALL_OLD.txt устарел (.NET 5, exe, C:\\ путь)", "Новый разработчик получит актуальную инструкцию установки", "P1"],
        ["9", "Обновить CHANGELOG: добавить записи для версий 1.1, 1.2, 1.3", "CHANGELOG.md", "CHANGELOG заканчивается на 1.0; проект на 1.3; нет описания изменений", "История версий будет полной; разработчик поймет, что изменилось", "P2"],
        ["10", "Удалить или переместить INSTALL_OLD.txt в архив/историю", "docs/INSTALL_OLD.txt", "Файл запутан — новичок может принять его за актуальную инструкцию", "Снижение риска путаницы; современная и старая инструкции разделены", "P2"],
        ["11", "Удалить / очистить файлы local notes (notes_local.txt, temp/debug_notes.txt)", "notes_local.txt, temp/debug_notes.txt", "Файлы содержат личные заметки, не предназначены для репозитория; не в Git", "Репозиторий будет чище; новичок не увидит служебные заметки", "P3"],
        ["12", "Обновить project-info.txt: добавить дату проверки и ответственного", "project-info.txt", "project-info.txt: 'Последняя проверка: неизвестно.' — нет свежих данных", "Новый разработчик сможет оценить актуальность информации", "P3"],
        ["13", "Добавить раздел 'Контакты' в README", "README.md", "README: 'Контакт для вопросов: отсутствует'", "Новый разработчик сможет задать вопросы", "P3"],
        ["14", "Обновить оценку полноты: добавить описание функций, ограничений, версий", "README.md", "README: нет описания функций, ограничений; CHANGELOG неполный", "Новый разработчик получит полную информацию", "P2"],
    ]
    add_table(doc, headers_prop, rows_prop, [0.5, 2.5, 1.5, 2.5, 2.5, 0.8])
    
    doc.add_page_break()
    
    # ===== 15. ПРИОРИТЕТЫ =====
    add_heading(doc, "15. Приоритеты обновления", 1)
    
    add_para(doc, "P1 — Исправить в первую очередь (критично для запуска):", bold=True)
    add_bullet(doc, "Обновить .NET в README (на .NET 8) — P1")
    add_bullet(doc, "Обновить версию приложения в README (на 1.3) — P1")
    add_bullet(doc, "Исправить путь к исходному коду в README — P1")
    add_bullet(doc, "Исправить каталог запуска в README — P1")
    add_bullet(doc, "Обновить .NET в USER_GUIDE (на .NET 8) — P1")
    add_bullet(doc, "Исправить описание интерфейса в USER_GUIDE (GUI → консоль) — P1")
    add_bullet(doc, "Исправить путь в USER_GUIDE (TaskTrackerApp → src/TaskTracker/) — P1")
    add_bullet(doc, "Создать современную INSTALL.md — P1")
    
    doc.add_paragraph()
    add_para(doc, "P2 — Важно, но проект можно изучать и без этого:", bold=True)
    add_bullet(doc, "Обновить CHANGELOG (добавить 1.1-1.3) — P2")
    add_bullet(doc, "Обновить описание возможностей в README — P2")
    
    doc.add_paragraph()
    add_para(doc, "P3 — Косметические улучшения:", bold=True)
    add_bullet(doc, "Очистить local notes (notes_local.txt, temp/) — P3")
    add_bullet(doc, "Обновить project-info.txt (добавить дату) — P3")
    add_bullet(doc, "Добавить контакты в README — P3")
    
    doc.add_page_break()
    
    # ===== 16. ВЫВОД =====
    add_heading(doc, "16. Итоговый вывод", 1)
    add_para(doc, "Вывод должен содержать 12-15 предложений. Ответы на вопросы из §23.2.", italic=True)
    
    add_para(doc, "Документы, присутствующие в проекте:", bold=True)
    add_bullet(doc, "README.md — главный информационный файл (устарел: .NET 6 вместо .NET 8, версия 1.0 вместо 1.3, неверное описание структуры)")
    add_bullet(doc, "CHANGELOG.md — история изменений (неполный: только версия 1.0, проект на 1.3)")
    add_bullet(doc, "USER_GUIDE.md — руководство пользователя (неверное: описывает GUI, указывает несуществующий каталог TaskTrackerApp, требует .NET 6)")
    add_bullet(doc, "INSTALL_OLD.txt — старая инструкция установки (.NET 5 Runtime, exe, абсолютный путь C:\\TaskTracker\\db\\tasks.db)")
    add_bullet(doc, "Отсутствует: современная INSTALL.md — актуальная инструкция установки не создана")
    add_bullet(doc, "Локальные/служебные файлы: project-info.txt (нет даты проверки), notes_local.txt и temp/debug_notes.txt (личные заметки, не в Git)")
    
    doc.add_paragraph()
    add_para(doc, "Главный документ — README.md, но он не является надежным источником, так как содержит несколько критических ошибок.", bold=True)
    
    doc.add_paragraph()
    add_para(doc, "Актуальность README:", bold=True)
    add_bullet(doc, "Низкая. README требует .NET 6 (проект — .NET 8), показывает версию 1.0 (фактически 1.3), указывает неверное расположение исходного кода ('в корне' вместо src/TaskTracker/), неверный каталог запуска ('TaskTracker/' вместо src/TaskTracker/).")
    
    doc.add_paragraph()
    add_para(doc, "Совпадение версии .NET:", bold=True)
    add_bullet(doc, "НЕ СОВПАДАЕТ. README, USER_GUIDE и INSTALL_OLD указывают .NET 6 или .NET 5, но проект использует .NET 8 (<TargetFramework>net8.0</TargetFramework> в TaskTracker.csproj).")
    
    doc.add_paragraph()
    add_para(doc, "Совпадение версии приложения:", bold=True)
    add_bullet(doc, "НЕ СОВПАДАЕТ. README и CHANGELOG показывают версию 1.0, но Program.cs выводит 'TaskTracker 1.3'.")
    
    doc.add_paragraph()
    add_para(doc, "Правильность описания структуры:", bold=True)
    add_bullet(doc, "НЕПРАВИЛЬНО. README утверждает, что исходный код находится в корне проекта, но фактически он находится в src/TaskTracker/.")
    
    doc.add_paragraph()
    add_para(doc, "Работоспособность инструкции запуска:", bold=True)
    add_bullet(doc, "НЕРАБОТОСПОСОБНА. README предлагает 'Откройте папку TaskTracker' — такой папки нет (нужно src/TaskTracker/). USER_GUIDE предлагает 'TaskTrackerApp' — каталог не существует. INSTALL_OLD предлагает запустить TaskTracker.exe из Release — файла нет (проект собирается через dotnet run).")
    
    doc.add_paragraph()
    add_para(doc, "Актуальность руководства пользователя:", bold=True)
    add_bullet(doc, "НИЗКАЯ. USER_GUIDE описывает графический интерфейс, но проект является консольным приложением. Указывает несуществующий каталог. Требует неверную версию .NET. Полностью не соответствует реальному продукту.")
    
    doc.add_paragraph()
    add_para(doc, "Полнота CHANGELOG:", bold=True)
    add_bullet(doc, "НЕПОЛНА. CHANGELOG содержит только версию 1.0 (от 2025-02-12), но проект уже на версии 1.3. Отсутствуют записи о версиях 1.1, 1.2, 1.3.")
    
    doc.add_paragraph()
    add_para(doc, "Противоречия между документами:", bold=True)
    add_bullet(doc, "ЕСТЬ. README (.NET 6) ⇄ USER_GUIDE (.NET 6) ⇄ INSTALL_OLD (.NET 5) ⇄ .csproj (net8.0) — три документа расходятся с фактом. README (.NET 6) ⇄ INSTALL_OLD (.NET 5) — разные версии в двух документах. VERSION (1.0 в README, USER_GUIDE, CHANGELOG) ⇄ факт (1.3 в Program.cs).")
    
    doc.add_paragraph()
    add_para(doc, "Жёстко прописанные пути:", bold=True)
    add_bullet(doc, "ЕСТЬ. В INSTALL_OLD.txt присутствует абсолютный путь 'C:\\TaskTracker\\db\\tasks.db', который привязан к конкретной машине и непереносим.")
    
    doc.add_paragraph()
    add_para(doc, "Достаточность информации новому разработчику:", bold=True)
    add_bullet(doc, "НЕДОСТАТОЧНА. Новичок не сможет: правильно установить .NET (указаны неверные версии), найти исходный код (неверное описание), запустить проект (неверные каталоги), понять тип приложения (GUI вместо консоли), понять историю изменений (CHANGELOG обрывается на 1.0).")
    
    doc.add_paragraph()
    add_para(doc, "Наиболее критичные проблемы (P1):", bold=True)
    add_bullet(doc, "1. Неправильная версия .NET в README и USER_GUIDE (.NET 6 вместо .NET 8) — блокирует запуск.")
    add_bullet(doc, "2. Несуществующий каталог в USER_GUIDE (TaskTrackerApp) — команда не выполнима.")
    add_bullet(doc, "3. Отсутствие современной INSTALL.md — нет актуальной инструкции.")
    add_bullet(doc, "4. Описание GUI в USER_GUIDE вместо консольного приложения — неверное представление о продукте.")
    add_bullet(doc, "5. Неверный путь запуска в README (TaskTracker/ вместо src/TaskTracker/) — команда не сработает.")
    
    doc.add_paragraph()
    add_para(doc, "Первичные действия по обновлению:", bold=True)
    add_bullet(doc, "Сначала исправить .NET 6 → .NET 8 в README и USER_GUIDE, иначе проект не запустится.")
    add_bullet(doc, "Затем исправить пути и описание интерфейса в USER_GUIDE.")
    add_bullet(doc, "Создать современную INSTALL.md.")
    add_bullet(doc, "Обновить версию приложения и CHANGELOG.")
    
    doc.add_paragraph()
    add_para(doc, "Можно ли считать документацию готовой к передаче другому разработчику:", bold=True)
    add_bullet(doc, "НЕТ. Документация содержит критические ошибки, которые блокируют запуск проекта, вводит в заблуждение относительно типа приложения, содержит несуществующие пути и устаревшие версии. Передача проекта с такой документацией новому разработчику приведет к потере времени и непониманию.")

    
    # ===== 17. СКРИНШОТЫ =====
    doc.add_page_break()
    add_heading(doc, "Обязательные скриншоты (см. выше)", 1)
    add_para(doc, "В соответствии с требованиями §24.2, в отчёте должны быть следующие скриншоты:", bold=True)
    
    screenshots_list = [
        ("Рис. 1", "Список документов", "docs_list.png", "Скриншот вывода команды Get-ChildItem с документами проекта"),
        ("Рис. 2", "README.md", "readme.png", "Содержимое README.md"),
        ("Рис. 3", "TargetFramework в .csproj", "csproj_tf.png", "Содержимое TaskTracker.csproj — <TargetFramework>net8.0</TargetFramework>"),
        ("Рис. 4", "Версия программы в Program.cs", "program_ver.png", "Содержимое Program.cs — Console.WriteLine(\"TaskTracker 1.3\")"),
        ("Рис. 5", "Пример несоответствия версии", "mismatch.png", "Сравнение версий .NET и приложения (таблицы из анализа)"),
        ("Рис. 6", "USER_GUIDE.md", "userguide.png", "Содержимое USER_GUIDE.md"),
        ("Рис. 7", "Пример неверного пути", "wrong_path.png", "Доказательство отсутствия каталога TaskTrackerApp"),
        ("Рис. 8", "CHANGELOG.md", "changelog.png", "Содержимое CHANGELOG.md"),
        ("Рис. 9", "git log -- README.md", "gitlog_readme.png", "История коммитов README.md"),
        ("Рис. 10", "Пример найденного противоречия", "contradiction.png", "Доказательство противоречия между документами и фактом"),
    ]
    
    headers_sc = ["Обозначение", "Что показано", "Файл", "Описание"]
    rows_sc = [[s[0], s[1], s[2], s[3]] for s in screenshots_list]
    add_table(doc, headers_sc, rows_sc, [1.5, 2.5, 2.5, 3.0])
    
    doc.add_paragraph()
    add_para(doc, "Примечание: многие скриншоты вставлены в соответствующие разделы выше. В данном разделе приведён их полный список с назначением.", italic=True)
    
    # ===== 18. СТРУКТУРА ОТЧЁТА =====
    doc.add_paragraph()
    add_heading(doc, "Структура отчёта (по §25)", 1)
    add_para(doc, "Практическая работа №3, тема и цель указаны на титульном листе. Основные разделы:")
    
    structure_items = [
        "1. Инвентаризация документации (Таблица 1)",
        "2. Анализ README (Таблицы 2, 3, 4)",
        "3. Анализ USER_GUIDE (Таблица 5)",
        "4. Анализ инструкции установки (Таблица 6)",
        "5. Анализ CHANGELOG (Таблица 7)",
        "6. Матрица противоречий (Таблица 8)",
        "7. История обновления документации (Таблица 9)",
        "8. Анализ путей (поиск hardcoded path, Таблица путей)",
        "9. Анализ версий (Таблица упоминаний .NET, Таблица версий)",
        "10. Оценка полноты (Таблица 11)",
        "11. Проверка документации глазами нового разработчика (Таблица 12)",
        "12. Реестр проблем (Таблица 13)",
        "13. Оценка качества документации (Таблица 14, шкала 1-5)",
        "14. Предложения по обновлению (Таблица 15, минимум 8)",
        "15. Приоритеты (P1/P2/P3)",
        "16. Вывод (12-15 предложений по вопросам §23.2)",
    ]
    
    for item in structure_items:
        add_bullet(doc, item)
    
    # ===== ПРИЛОЖЕНИЕ: КОМАНДЫ =====
    doc.add_page_break()
    add_heading(doc, "Приложение. Выполненные команды", 1)
    add_para(doc, "Все команды, использованные в работе (из методички и дополнительные):", italic=True)
    
    commands_list = [
        "Get-Location — проверка текущего каталога",
        "Get-ChildItem -Recurse -File -Include *.md,*.txt | Where-Object { $_.FullName -notmatch '\\.git\\' } | Select-Object FullName — поиск документов",
        "Get-ChildItem -Recurse -Directory | Where-Object { $_.Name -eq 'TaskTrackerApp' } — проверка существования каталога",
        "Get-ChildItem -Recurse -File -Include *.md,*.txt | Select-String -Pattern 'C:\\\\' — поиск hardcoded path",
        "Get-ChildItem -Recurse -File -Include *.md,*.txt | Select-String -Pattern '\\.NET' — поиск упоминаний .NET",
        "Get-ChildItem -Recurse -File -Include *.md,*.txt,*.cs,*.csproj | Select-String -Pattern 'TaskTracker' — поиск упоминаний проекта",
        "git status — состояние репозитория",
        "git log --oneline -- README.md — история README",
        "git log --oneline -- CHANGELOG.md — история CHANGELOG",
        "git log --oneline -- docs/USER_GUIDE.md — история USER_GUIDE",
        "cat README.md — чтение README",
        "cat src/TaskTracker/TaskTracker.csproj — чтение файла проекта",
        "cat src/TaskTracker/Program.cs — чтение Program.cs",
        "cat CHANGELOG.md — чтение CHANGELOG",
        "cat docs/USER_GUIDE.md — чтение USER_GUIDE",
        "cat docs/INSTALL_OLD.txt — чтение старой инструкции установки",
    ]
    
    for cmd in commands_list:
        add_bullet(doc, cmd)
    
    # ===== СОХРАНЕНИЕ =====
    output_path = os.path.join(PROJECT_PATH, "Практическая_работа_№3_отчёт.docx")
    doc.save(output_path)
    print(f"Отчёт сохранён: {output_path}")
    return output_path

if __name__ == "__main__":
    create_report()
