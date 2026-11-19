import csv
import html
import io
import json
import os
import re
import shutil
import sys
import threading
import traceback

import flet as ft
import openpyxl
from docx import Document
from fpdf import FPDF
from openpyxl.styles import Alignment, Font
from pptx import Presentation

try:
    import flet_permission_handler as fph
except ImportError:
    fph = None

DOWNLOAD_DIR = "/storage/emulated/0/Download"
ASSETS_DIR = os.path.join(os.path.dirname(__file__), "assets")
FONT_PATH = os.path.join(ASSETS_DIR, "DejaVuSans.ttf")


def ensure_download_dir():
    if not os.path.exists(DOWNLOAD_DIR):
        try:
            os.makedirs(DOWNLOAD_DIR, exist_ok=True)
        except Exception:
            pass


def parse_csv(csv_string):
    if not csv_string or not str(csv_string).strip():
        return []
    reader = csv.reader(io.StringIO(str(csv_string)))
    return [row for row in reader]


def save_file_to_download(filename, content_bytes_or_path, is_binary=True):
    ensure_download_dir()
    safe_name = os.path.basename(str(filename).strip()) or "document"
    full_path = os.path.join(DOWNLOAD_DIR, safe_name)
    if is_binary:
        data = (
            content_bytes_or_path
            if isinstance(content_bytes_or_path, (bytes, bytearray))
            else str(content_bytes_or_path).encode("utf-8")
        )
        with open(full_path, "wb") as f:
            f.write(data)
    else:
        shutil.copy(content_bytes_or_path, full_path)
    return full_path


def _escape_csv_cell(value):
    s = str(value or "")
    if any(ch in s for ch in (",", '"', "\n")):
        return '"' + s.replace('"', '""') + '"'
    return s


def parse_layout_from_text(raw_text, source_filename=""):
    trimmed = str(raw_text or "").strip()
    if not trimmed:
        return "", "", "", ""

    lower_name = str(source_filename or "").lower()

    # 1. JSON format
    if lower_name.endswith(".json") or (trimmed.startswith("{") and trimmed.endswith("}")):
        try:
            parsed = json.loads(trimmed)
            if isinstance(parsed, dict):
                list_val = (
                    parsed.get("listItems")
                    or parsed.get("list")
                    or parsed.get("items")
                    or ""
                )
                if isinstance(list_val, list):
                    list_str = "\n".join(str(x) for x in list_val)
                else:
                    list_str = str(list_val)

                table_val = parsed.get("tableCsv") or parsed.get("table") or ""
                if isinstance(table_val, list):
                    rows = []
                    for r in table_val:
                        if isinstance(r, list):
                            rows.append(",".join(_escape_csv_cell(c) for c in r))
                        else:
                            rows.append(str(r))
                    table_str = "\n".join(rows)
                else:
                    table_str = str(table_val)

                return (
                    str(parsed.get("title", "")),
                    str(parsed.get("text") or parsed.get("body") or ""),
                    list_str,
                    table_str,
                )
        except Exception:
            pass

    # 2. CSV file
    if lower_name.endswith(".csv"):
        base_title = os.path.splitext(os.path.basename(source_filename))[0]
        return base_title, "", "", trimmed

    # 3. Markdown / structured text
    title = ""
    text_lines = []
    list_lines = []
    table_rows = []
    in_table_block = False

    for raw_line in str(raw_text).splitlines():
        line = raw_line.rstrip()
        stripped = line.strip()

        if re.match(r"^TABLE:\s*$", stripped, re.IGNORECASE):
            in_table_block = True
            continue

        if in_table_block:
            if not stripped:
                in_table_block = False
                continue
            if stripped.startswith("|") and stripped.endswith("|"):
                if re.match(r"^\|[\s:\-|]+\|$", stripped):
                    continue
                cells = [_escape_csv_cell(c.strip()) for c in stripped[1:-1].split("|")]
                table_rows.append(",".join(cells))
            else:
                table_rows.append(stripped)
            continue

        title_match = re.match(
            r"^(?:TITLE|Заголовок)\s*:\s*(.+)$", stripped, re.IGNORECASE
        )
        if title_match and not title:
            title = title_match.group(1).strip()
            continue

        h1_match = re.match(r"^#\s+(.+)$", stripped)
        if h1_match and not title:
            title = h1_match.group(1).strip()
            continue

        sub_match = re.match(r"^#{2,6}\s+(.+)$", stripped)
        if sub_match:
            text_lines.append(sub_match.group(1).strip())
            continue

        if stripped.startswith("|") and stripped.endswith("|") and len(stripped) > 2:
            if not re.sub(r"[\s|:\-]", "", stripped):
                continue
            cells = [_escape_csv_cell(c.strip()) for c in stripped[1:-1].split("|")]
            table_rows.append(",".join(cells))
            continue

        bullet_match = re.match(r"^(?:[-*•]|\d+[.)])\s+(.+)$", stripped)
        if bullet_match:
            list_lines.append(bullet_match.group(1).strip())
            continue

        text_lines.append(line)

    if not title and source_filename:
        title = os.path.splitext(os.path.basename(source_filename))[0]

    return (
        title,
        "\n".join(text_lines).strip(),
        "\n".join(list_lines),
        "\n".join(table_rows),
    )


def create_docx(title, text, list_items, table_csv, filename):
    doc = Document()
    if title:
        doc.add_heading(str(title), level=1)
    if text:
        doc.add_paragraph(str(text))
    if list_items:
        for item in str(list_items).splitlines():
            if item.strip():
                doc.add_paragraph(item.strip(), style="List Bullet")
    table_data = parse_csv(table_csv)
    if table_data:
        rows = len(table_data)
        cols = len(table_data[0]) if rows > 0 else 0
        if cols > 0:
            table = doc.add_table(rows=rows, cols=cols)
            table.style = "Table Grid"
            for i, row in enumerate(table_data):
                for j, cell_value in enumerate(row):
                    if j < cols:
                        table.cell(i, j).text = str(cell_value)
    temp_path = os.path.join(os.path.dirname(__file__), os.path.basename(filename))
    doc.save(temp_path)
    save_file_to_download(filename, temp_path, is_binary=False)
    os.remove(temp_path)
    return f"Файл {filename} сохранен в Downloads"


def create_pptx(title, text, list_items, filename):
    prs = Presentation()
    slide_layout = prs.slide_layouts[1]
    slide = prs.slides.add_slide(slide_layout)
    slide.shapes.title.text = str(title or "")
    body_shape = slide.placeholders[1]
    tf = body_shape.text_frame
    tf.text = str(text or "")
    if list_items:
        for item in str(list_items).splitlines():
            if item.strip():
                p = tf.add_paragraph()
                p.text = item.strip()
                p.level = 1
    temp_path = os.path.join(os.path.dirname(__file__), os.path.basename(filename))
    prs.save(temp_path)
    save_file_to_download(filename, temp_path, is_binary=False)
    os.remove(temp_path)
    return f"Файл {filename} сохранен в Downloads"


def create_xlsx(table_csv, filename):
    table_data = parse_csv(table_csv)
    if not table_data:
        raise ValueError("Нет данных для таблицы Excel.")
    wb = openpyxl.Workbook()
    ws = wb.active
    for i, row in enumerate(table_data):
        for j, value in enumerate(row):
            cell = ws.cell(row=i + 1, column=j + 1, value=value)
            cell.alignment = Alignment(wrap_text=True)
            if i == 0:
                cell.font = Font(bold=True)
    temp_path = os.path.join(os.path.dirname(__file__), os.path.basename(filename))
    wb.save(temp_path)
    save_file_to_download(filename, temp_path, is_binary=False)
    os.remove(temp_path)
    return f"Файл {filename} сохранен в Downloads"


def _is_valid_ttf(font_path):
    if not os.path.exists(font_path):
        return False
    try:
        with open(font_path, "rb") as f:
            header = f.read(4)
        return header in (b"\x00\x01\x00\x00", b"OTTO", b"true")
    except Exception:
        return False


def create_pdf(title, text, list_items, filename):
    pdf = FPDF()
    pdf.add_page()
    if _is_valid_ttf(FONT_PATH):
        pdf.add_font("DejaVu", "", FONT_PATH)
        pdf.set_font("DejaVu", size=16)
    else:
        pdf.set_font("Helvetica", size=16)
    pdf.cell(0, 10, str(title or ""), ln=True, align="C")
    pdf.ln(10)
    pdf.set_font_size(12)
    pdf.multi_cell(0, 8, str(text or ""))
    if list_items:
        for item in str(list_items).splitlines():
            if item.strip():
                pdf.multi_cell(0, 8, f"• {item.strip()}")
    temp_path = os.path.join(os.path.dirname(__file__), os.path.basename(filename))
    pdf.output(temp_path)
    save_file_to_download(filename, temp_path, is_binary=False)
    os.remove(temp_path)
    return f"Файл {filename} сохранен в Downloads"


def create_txt(title, text, list_items, table_csv, filename):
    parts = []
    if title:
        parts.append(str(title).upper())
        parts.append("=" * min(max(len(str(title)), 3), 60))
        parts.append("")
    if text:
        parts.append(str(text))
        parts.append("")
    if list_items:
        for item in str(list_items).splitlines():
            if item.strip():
                parts.append(f"• {item.strip()}")
        parts.append("")
    for row in parse_csv(table_csv):
        parts.append("\t|\t".join(row))
    content = "\n".join(parts).strip() + "\n"
    save_file_to_download(filename, content.encode("utf-8"), is_binary=True)
    return f"Файл {filename} сохранен в Downloads"


def create_md(title, text, list_items, table_csv, filename):
    parts = []
    if title:
        parts.append(f"# {title}\n")
    if text:
        parts.append(f"{text}\n")
    if list_items:
        for item in str(list_items).splitlines():
            if item.strip():
                parts.append(f"- {item.strip()}")
        parts.append("")
    table_data = parse_csv(table_csv)
    if table_data:
        cols = len(table_data[0])
        parts.append("| " + " | ".join(table_data[0]) + " |")
        parts.append("| " + " | ".join(["---"] * cols) + " |")
        for row in table_data[1:]:
            padded = [row[i] if i < len(row) else "" for i in range(cols)]
            parts.append("| " + " | ".join(padded) + " |")
    content = "\n".join(parts).strip() + "\n"
    save_file_to_download(filename, content.encode("utf-8"), is_binary=True)
    return f"Файл {filename} сохранен в Downloads"


def create_html(title, text, list_items, table_csv, filename):
    sections = []
    if title:
        sections.append(f"<h1>{html.escape(str(title))}</h1>")
    if text:
        for para in str(text).splitlines():
            if para.strip():
                sections.append(f"<p>{html.escape(para)}</p>")
    if list_items:
        items = [
            f"<li>{html.escape(i.strip())}</li>"
            for i in str(list_items).splitlines()
            if i.strip()
        ]
        if items:
            sections.append("<ul>" + "".join(items) + "</ul>")
    table_data = parse_csv(table_csv)
    if table_data:
        rows_html = []
        for idx, row in enumerate(table_data):
            tag = "th" if idx == 0 else "td"
            rows_html.append(
                "<tr>" + "".join(f"<{tag}>{html.escape(c)}</{tag}>" for c in row) + "</tr>"
            )
        sections.append("<table>" + "".join(rows_html) + "</table>")

    doc_html = (
        f'<!doctype html><html lang="ru"><head><meta charset="UTF-8">'
        f"<title>{html.escape(str(title or filename))}</title>"
        f"<style>body{{font-family:sans-serif;max-width:800px;margin:40px auto;padding:0 16px}}"
        f"table{{border-collapse:collapse;width:100%}}th,td{{border:1px solid #ccc;padding:6px 8px}}</style>"
        f"</head><body>{''.join(sections)}</body></html>"
    )
    save_file_to_download(filename, doc_html.encode("utf-8"), is_binary=True)
    return f"Файл {filename} сохранен в Downloads"


def create_from_text(raw_text, fmt="docx", filename="layout"):
    title, text, list_items, table_csv = parse_layout_from_text(raw_text)
    if not filename.endswith(f".{fmt}"):
        filename = f"{filename}.{fmt}"
    if fmt == "docx":
        return create_docx(title, text, list_items, table_csv, filename)
    if fmt == "pptx":
        return create_pptx(title, text, list_items, filename)
    if fmt == "xlsx":
        return create_xlsx(table_csv, filename)
    if fmt == "pdf":
        return create_pdf(title, text, list_items, filename)
    if fmt == "txt":
        return create_txt(title, text, list_items, table_csv, filename)
    if fmt == "md":
        return create_md(title, text, list_items, table_csv, filename)
    if fmt == "html":
        return create_html(title, text, list_items, table_csv, filename)
    raise ValueError(f"Неизвестный формат: {fmt}")


def execute_user_script(code, log_output):
    def run():
        old_stdout = sys.stdout
        try:
            namespace = {
                "create_docx": create_docx,
                "create_pptx": create_pptx,
                "create_xlsx": create_xlsx,
                "create_pdf": create_pdf,
                "create_txt": create_txt,
                "create_md": create_md,
                "create_html": create_html,
                "create_from_text": create_from_text,
                "parse_layout": parse_layout_from_text,
                "save_file_to_download": save_file_to_download,
                "os": os,
                "print": lambda *args: log_output(f"[PRINT] {' '.join(map(str, args))}"),
            }
            sys.stdout = io.StringIO()
            exec(code, namespace)
            output = sys.stdout.getvalue()
            if output:
                log_output(f"Вывод скрипта:\n{output}")
            log_output("✅ Скрипт выполнен успешно.")
        except Exception as e:
            log_output(f"❌ Ошибка: {e}\n{traceback.format_exc()}")
        finally:
            sys.stdout = old_stdout

    threading.Thread(target=run, daemon=True).start()


def main(page: ft.Page):
    page.title = "DocBuilder Mobile"
    page.theme_mode = ft.ThemeMode.LIGHT
    page.scroll = ft.ScrollMode.ADAPTIVE

    ph = fph.PermissionHandler() if fph is not None else None
    if ph is not None:
        page.services.append(ph)

    file_picker = ft.FilePicker()
    page.services.append(file_picker)

    script_log = ft.Text("", selectable=True, size=12)
    layout_log = ft.Text("", selectable=True, size=12)

    def show_snack(msg):
        page.show_dialog(ft.SnackBar(ft.Text(msg)))

    def log_to_ui(message):
        script_log.value += f"{message}\n"
        page.update()

    def log_layout(message):
        layout_log.value += f"{message}\n"
        page.update()

    async def request_storage_permission(e):
        if ph is None:
            log_to_ui("Управление разрешениями выполняется в настройках Android.")
            show_snack("Модуль разрешений недоступен в этой сборке.")
            return
        try:
            status = await ph.request(fph.Permission.MANAGE_EXTERNAL_STORAGE)
            if status == fph.PermissionStatus.GRANTED:
                log_to_ui("Разрешение на управление внешним хранилищем получено.")
                show_snack("Разрешение получено.")
            else:
                log_to_ui(f"Статус разрешения: {status}")
                show_snack("Разрешение не получено. Сохранение может не работать.")
        except Exception as ex:
            log_to_ui(f"Ошибка при запросе разрешения: {ex}")

    # Вкладка 1: Шаблоны
    title_field = ft.TextField(label="Заголовок", hint_text="Введите заголовок документа")
    text_field = ft.TextField(label="Текст", multiline=True, min_lines=3, max_lines=6)
    list_field = ft.TextField(
        label="Список (каждый пункт с новой строки)",
        multiline=True,
        min_lines=2,
        max_lines=4,
    )
    table_field = ft.TextField(
        label="Таблица (CSV: строки через Enter, колонки через запятую)",
        multiline=True,
        min_lines=3,
        max_lines=6,
    )
    filename_field = ft.TextField(label="Имя файла (с расширением)", value="document.docx")

    def make_create_handler(format_type):
        def handler(e):
            try:
                fname = filename_field.value.strip()
                if not fname:
                    raise ValueError("Укажите имя файла.")
                if not fname.endswith(f".{format_type}"):
                    fname += f".{format_type}"
                t = title_field.value
                txt = text_field.value
                lst = list_field.value
                tbl = table_field.value
                if format_type == "docx":
                    res = create_docx(t, txt, lst, tbl, fname)
                elif format_type == "pptx":
                    res = create_pptx(t, txt, lst, fname)
                elif format_type == "xlsx":
                    res = create_xlsx(tbl, fname)
                elif format_type == "pdf":
                    res = create_pdf(t, txt, lst, fname)
                elif format_type == "md":
                    res = create_md(t, txt, lst, tbl, fname)
                else:
                    res = create_txt(t, txt, lst, tbl, fname)
                show_snack(res)
            except Exception as ex:
                show_snack(f"Ошибка: {ex}")

        return handler

    doc_tab = ft.Column(
        controls=[
            ft.Text("Создание документа", size=20, weight=ft.FontWeight.BOLD),
            title_field,
            text_field,
            list_field,
            table_field,
            filename_field,
            ft.Row(
                controls=[
                    ft.Button(content="Word (.docx)", on_click=make_create_handler("docx")),
                    ft.Button(content="PowerPoint (.pptx)", on_click=make_create_handler("pptx")),
                    ft.Button(content="Excel (.xlsx)", on_click=make_create_handler("xlsx")),
                    ft.Button(content="PDF (.pdf)", on_click=make_create_handler("pdf")),
                    ft.Button(content="Markdown (.md)", on_click=make_create_handler("md")),
                    ft.Button(content="Текст (.txt)", on_click=make_create_handler("txt")),
                ],
                wrap=True,
            ),
            ft.Divider(),
            ft.Text("Все файлы сохраняются в папку Download.", size=12, color=ft.Colors.GREY),
            ft.Button(
                content="Запросить разрешение на запись",
                on_click=request_storage_permission,
            ),
        ],
        scroll=ft.ScrollMode.ADAPTIVE,
        spacing=10,
    )

    # Вкладка 2: Верстка из файла
    layout_content_field = ft.TextField(
        label="Содержимое текстового файла / разметка",
        multiline=True,
        min_lines=8,
        max_lines=16,
        hint_text="# Заголовок\nТекст\n- Пункт 1\n| A | B |\n| --- | --- |\n| 1 | 2 |",
    )
    layout_filename_field = ft.TextField(
        label="Имя выходного файла (без расширения)",
        value="layout",
    )
    loaded_source_name = {"name": ""}

    async def pick_layout_file(e):
        files = await file_picker.pick_files(
            allow_multiple=False,
            allowed_extensions=["txt", "md", "py", "js", "json", "csv", "yaml", "yml", "log"],
        )
        if files:
            f = files[0]
            if f.path and os.path.exists(f.path):
                with open(f.path, "r", encoding="utf-8", errors="replace") as fp:
                    layout_content_field.value = fp.read()
                loaded_source_name["name"] = f.name
                layout_filename_field.value = os.path.splitext(f.name)[0] or "layout"
                layout_log.value = f"Загружен файл: {f.name}\n"
                page.update()

    def make_typeset_handler(fmt):
        def handler(e):
            try:
                raw = layout_content_field.value
                if not raw or not raw.strip():
                    raise ValueError("Текстовый файл пуст.")
                base = os.path.splitext(layout_filename_field.value.strip() or "layout")[0]
                t, txt, lst, tbl = parse_layout_from_text(raw, loaded_source_name["name"])
                fname = f"{base}.{fmt}"
                if fmt == "docx":
                    res = create_docx(t, txt, lst, tbl, fname)
                elif fmt == "pptx":
                    res = create_pptx(t, txt, lst, fname)
                elif fmt == "xlsx":
                    res = create_xlsx(tbl, fname)
                elif fmt == "pdf":
                    res = create_pdf(t, txt, lst, fname)
                elif fmt == "md":
                    res = create_md(t, txt, lst, tbl, fname)
                elif fmt == "html":
                    res = create_html(t, txt, lst, tbl, fname)
                else:
                    res = create_txt(t, txt, lst, tbl, fname)
                layout_log.value = f"{res}\n"
                show_snack(res)
            except Exception as ex:
                layout_log.value = f"❌ Ошибка: {ex}\n"
                show_snack(f"Ошибка: {ex}")

        return handler

    def run_layout_as_script(e):
        layout_log.value = "Выполнение...\n"
        page.update()
        execute_user_script(layout_content_field.value, log_layout)

    layout_tab = ft.Column(
        controls=[
            ft.Text("Верстка из текстового файла", size=20, weight=ft.FontWeight.BOLD),
            ft.Button(
                content="Выбрать файл (.txt, .md, .py, .json, .csv)",
                on_click=pick_layout_file,
            ),
            layout_content_field,
            layout_filename_field,
            ft.Row(
                controls=[
                    ft.Button(content="Word (.docx)", on_click=make_typeset_handler("docx")),
                    ft.Button(content="PDF (.pdf)", on_click=make_typeset_handler("pdf")),
                    ft.Button(content="PowerPoint (.pptx)", on_click=make_typeset_handler("pptx")),
                    ft.Button(content="Excel (.xlsx)", on_click=make_typeset_handler("xlsx")),
                    ft.Button(content="Markdown (.md)", on_click=make_typeset_handler("md")),
                    ft.Button(content="HTML (.html)", on_click=make_typeset_handler("html")),
                    ft.Button(content="Текст (.txt)", on_click=make_typeset_handler("txt")),
                    ft.Button(content="Выполнить как скрипт", on_click=run_layout_as_script),
                ],
                wrap=True,
            ),
            ft.Divider(),
            layout_log,
        ],
        scroll=ft.ScrollMode.ADAPTIVE,
        spacing=10,
    )

    # Вкладка 3: Свой скрипт
    script_code_field = ft.TextField(
        label="Python-код",
        multiline=True,
        min_lines=10,
        max_lines=20,
        hint_text="Доступны: create_docx, create_pptx, create_xlsx, create_pdf, create_from_text",
        value="# Пример:\n# create_docx('Заголовок', 'Текст', 'Пункт 1\\nПункт 2', 'A,B\\n1,2', 'my.docx')\nprint('Привет из скрипта!')",
    )

    async def pick_script_file(e):
        files = await file_picker.pick_files(
            allow_multiple=False,
            allowed_extensions=["py", "txt", "md", "js"],
        )
        if files:
            f = files[0]
            if f.path and os.path.exists(f.path):
                with open(f.path, "r", encoding="utf-8", errors="replace") as fp:
                    script_code_field.value = fp.read()
                script_log.value = f"Загружен скрипт: {f.name}\n"
                page.update()

    def run_script(e):
        script_log.value = "Выполнение...\n"
        page.update()
        execute_user_script(script_code_field.value, log_to_ui)

    script_tab = ft.Column(
        controls=[
            ft.Text("Свой скрипт", size=20, weight=ft.FontWeight.BOLD),
            script_code_field,
            ft.Row(
                controls=[
                    ft.Button(content="Выполнить", on_click=run_script),
                    ft.Button(content="Загрузить файл скрипта", on_click=pick_script_file),
                ],
                wrap=True,
            ),
            ft.Divider(),
            ft.Text("Лог выполнения:", weight=ft.FontWeight.BOLD),
            script_log,
        ],
        scroll=ft.ScrollMode.ADAPTIVE,
        spacing=10,
    )

    tabs = ft.Tabs(
        length=3,
        selected_index=0,
        expand=True,
        content=ft.Column(
            expand=True,
            controls=[
                ft.TabBar(
                    tabs=[
                        ft.Tab(label="Шаблоны"),
                        ft.Tab(label="Верстка из файла"),
                        ft.Tab(label="Свой скрипт"),
                    ]
                ),
                ft.TabBarView(
                    expand=True,
                    controls=[doc_tab, layout_tab, script_tab],
                ),
            ],
        ),
    )
    page.add(ft.SafeArea(content=tabs, expand=True))


if __name__ == "__main__":
    ft.run(main)
