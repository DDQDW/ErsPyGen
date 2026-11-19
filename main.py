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
APP_DIR = os.path.dirname(os.path.abspath(__file__))
ASSETS_DIR = os.path.join(APP_DIR, "assets")
FALLBACK_DIR = os.path.join(APP_DIR, "generated_docs")
FONT_PATH = os.path.join(ASSETS_DIR, "DejaVuSans.ttf")
GENERATED_EXTENSIONS = (".docx", ".pptx", ".xlsx", ".pdf", ".txt", ".md", ".html", ".csv", ".json")

# Thread-local or shared registry of files produced during an action
_created_files_lock = threading.Lock()
_recent_created_files = []


def _record_created_file(filename, data_bytes, saved_path, in_downloads):
    with _created_files_lock:
        _recent_created_files.append(
            {
                "filename": filename,
                "bytes": data_bytes,
                "path": saved_path,
                "in_downloads": in_downloads,
            }
        )


def _pop_recent_created_files():
    with _created_files_lock:
        items = list(_recent_created_files)
        _recent_created_files.clear()
        return items


def ensure_dir(path_dir):
    if not os.path.exists(path_dir):
        try:
            os.makedirs(path_dir, exist_ok=True)
        except Exception:
            pass


def parse_csv(csv_string):
    if not csv_string or not str(csv_string).strip():
        return []
    reader = csv.reader(io.StringIO(str(csv_string)))
    return [row for row in reader]


def save_file_to_download(filename, content_bytes_or_path, is_binary=True):
    safe_name = os.path.basename(str(filename).strip()) or "document"
    if is_binary:
        data_bytes = (
            bytes(content_bytes_or_path)
            if isinstance(content_bytes_or_path, (bytes, bytearray))
            else str(content_bytes_or_path).encode("utf-8")
        )
    else:
        with open(content_bytes_or_path, "rb") as f:
            data_bytes = f.read()

    # Try direct write to Android public Download folder first
    in_downloads = False
    saved_path = None
    try:
        ensure_dir(DOWNLOAD_DIR)
        candidate = os.path.join(DOWNLOAD_DIR, safe_name)
        with open(candidate, "wb") as f:
            f.write(data_bytes)
        saved_path = candidate
        in_downloads = True
    except Exception:
        ensure_dir(FALLBACK_DIR)
        candidate = os.path.join(FALLBACK_DIR, safe_name)
        with open(candidate, "wb") as f:
            f.write(data_bytes)
        saved_path = candidate
        in_downloads = False

    _record_created_file(safe_name, data_bytes, saved_path, in_downloads)
    return saved_path


def _format_save_message(filename, saved_path, in_downloads):
    if in_downloads:
        return f"Файл {filename} сохранен в Download ({saved_path})"
    return (
        f"Файл {filename} создан во внутренней папке ({saved_path}). "
        "Нет доступа к Download — нажмите «Сохранить в Загрузки» ниже."
    )


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
    buf = io.BytesIO()
    doc.save(buf)
    saved_path = save_file_to_download(filename, buf.getvalue(), is_binary=True)
    in_dl = saved_path.startswith(DOWNLOAD_DIR)
    return _format_save_message(os.path.basename(saved_path), saved_path, in_dl)


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
    buf = io.BytesIO()
    prs.save(buf)
    saved_path = save_file_to_download(filename, buf.getvalue(), is_binary=True)
    in_dl = saved_path.startswith(DOWNLOAD_DIR)
    return _format_save_message(os.path.basename(saved_path), saved_path, in_dl)


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
    buf = io.BytesIO()
    wb.save(buf)
    saved_path = save_file_to_download(filename, buf.getvalue(), is_binary=True)
    in_dl = saved_path.startswith(DOWNLOAD_DIR)
    return _format_save_message(os.path.basename(saved_path), saved_path, in_dl)


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
    pdf_bytes = bytes(pdf.output())
    saved_path = save_file_to_download(filename, pdf_bytes, is_binary=True)
    in_dl = saved_path.startswith(DOWNLOAD_DIR)
    return _format_save_message(os.path.basename(saved_path), saved_path, in_dl)


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
    saved_path = save_file_to_download(filename, content.encode("utf-8"), is_binary=True)
    in_dl = saved_path.startswith(DOWNLOAD_DIR)
    return _format_save_message(os.path.basename(saved_path), saved_path, in_dl)


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
    saved_path = save_file_to_download(filename, content.encode("utf-8"), is_binary=True)
    in_dl = saved_path.startswith(DOWNLOAD_DIR)
    return _format_save_message(os.path.basename(saved_path), saved_path, in_dl)


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
    saved_path = save_file_to_download(filename, doc_html.encode("utf-8"), is_binary=True)
    in_dl = saved_path.startswith(DOWNLOAD_DIR)
    return _format_save_message(os.path.basename(saved_path), saved_path, in_dl)


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


def _snapshot_dir(target_dir):
    snap = {}
    try:
        for name in os.listdir(target_dir):
            full = os.path.join(target_dir, name)
            if os.path.isfile(full):
                st = os.stat(full)
                snap[full] = (st.st_mtime_ns, st.st_size)
    except Exception:
        pass
    return snap


def execute_user_script(code, log_output, on_files_ready=None):
    def run():
        old_stdout = sys.stdout
        _pop_recent_created_files()
        watch_dirs = list(dict.fromkeys([os.getcwd(), APP_DIR, DOWNLOAD_DIR]))
        before = {}
        for d in watch_dirs:
            before.update(_snapshot_dir(d))

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
                "DOWNLOAD_DIR": DOWNLOAD_DIR,
                "os": os,
                "print": lambda *args: log_output(f"[PRINT] {' '.join(map(str, args))}"),
            }
            sys.stdout = io.StringIO()
            exec(code, namespace)
            output = sys.stdout.getvalue()
            if output:
                log_output(f"Вывод скрипта:\n{output}")

            # Detect files created directly by user scripts (e.g., doc.save("file.docx"))
            recorded = _pop_recent_created_files()
            recorded_names = {item["filename"] for item in recorded}

            for d in watch_dirs:
                for full_path, stat_val in _snapshot_dir(d).items():
                    if full_path not in before or before[full_path] != stat_val:
                        fname = os.path.basename(full_path)
                        if fname.lower().endswith(GENERATED_EXTENSIONS) and fname not in recorded_names:
                            if d == DOWNLOAD_DIR:
                                with open(full_path, "rb") as fp:
                                    data = fp.read()
                                recorded.append(
                                    {
                                        "filename": fname,
                                        "bytes": data,
                                        "path": full_path,
                                        "in_downloads": True,
                                    }
                                )
                            else:
                                # Copy from private app working dir to Download (or fallback)
                                save_file_to_download(fname, full_path, is_binary=False)
                                recorded.extend(_pop_recent_created_files())
                            recorded_names.add(fname)

            for item in recorded:
                if item["in_downloads"]:
                    log_output(f"📁 Сохранено в Download: {item['path']}")
                else:
                    log_output(
                        f"⚠️ Сохранено во внутреннюю память ({item['path']}). "
                        f"Нажмите «Сохранить {item['filename']} в Загрузки» ниже."
                    )

            if on_files_ready and recorded:
                on_files_ready(recorded)

            log_output("✅ Скрипт выполнен успешно.")
        except Exception as e:
            log_output(f"❌ Ошибка: {e}\n{traceback.format_exc()}")
        finally:
            sys.stdout = old_stdout

    threading.Thread(target=run, daemon=True).start()


PRIMARY_COLOR = "#1976d2"
SURFACE_COLOR = "#ffffff"
BORDER_COLOR = "#c4c7c5"
DIVIDER_COLOR = "#e0e2e0"
TEXT_COLOR = "#1f1f1f"
MUTED_COLOR = "#5f6368"
BTN_BG = "#f0f4fa"
LOG_BG = "#f6f8fa"
DROPZONE_BG = "#fafbfc"

SAMPLE_LAYOUT_TEXT = """# Отчет по проекту DocBuilder
TITLE: Отчет по проекту DocBuilder

Ниже приведены ключевые результаты работы и сводная таблица показателей.

## Основные задачи
- Поддержка верстки из текстовых файлов (.txt, .md, .json, .csv)
- Экспорт в Word, PowerPoint, Excel, PDF, Markdown и HTML
- Выполнение пользовательских скриптов

## Сводная таблица
| Модуль | Статус | Формат |
| --- | --- | --- |
| Шаблоны | Готово | DOCX/PPTX/XLSX/PDF |
| Текстовая верстка | Готово | TXT/MD/JSON/CSV |
| Скрипты | Готово | Python/JS |"""


def main(page: ft.Page):
    page.title = "DocBuilder Mobile"
    page.theme_mode = ft.ThemeMode.LIGHT
    page.theme = ft.Theme(color_scheme_seed=PRIMARY_COLOR)
    page.bgcolor = SURFACE_COLOR
    page.padding = 0

    ph = fph.PermissionHandler() if fph is not None else None
    if ph is not None:
        page.services.append(ph)

    file_picker = ft.FilePicker()
    page.services.append(file_picker)

    def make_field(label, **kwargs):
        return ft.TextField(
            label=label,
            bgcolor=SURFACE_COLOR,
            border_color=BORDER_COLOR,
            focused_border_color=PRIMARY_COLOR,
            border_radius=6,
            text_size=14,
            content_padding=ft.Padding.symmetric(horizontal=12, vertical=10),
            **kwargs,
        )

    def make_elevated_btn(text, on_click):
        return ft.Button(
            content=text,
            on_click=on_click,
            bgcolor=BTN_BG,
            color=PRIMARY_COLOR,
        )

    def make_outlined_btn(text, on_click):
        return ft.OutlinedButton(
            content=text,
            on_click=on_click,
        )

    script_log = ft.Text("", selectable=True, size=12, color=TEXT_COLOR)
    layout_log = ft.Text("", selectable=True, size=12, color=TEXT_COLOR)
    doc_save_buttons = ft.Row(controls=[], wrap=True, spacing=8, run_spacing=8)
    layout_save_buttons = ft.Row(controls=[], wrap=True, spacing=8, run_spacing=8)
    script_save_buttons = ft.Row(controls=[], wrap=True, spacing=8, run_spacing=8)

    def show_snack(msg):
        page.show_dialog(ft.SnackBar(ft.Text(msg)))

    def make_export_button(item):
        async def on_save_click(e):
            try:
                saved = await file_picker.save_file(
                    dialog_title="Сохранить документ",
                    file_name=item["filename"],
                    src_bytes=item["bytes"],
                )
                if saved:
                    show_snack(f"Сохранено: {saved}")
                else:
                    show_snack(f"Файл {item['filename']} сохранён через системный диалог.")
            except Exception as ex:
                show_snack(f"Ошибка сохранения: {ex}")

        return ft.Button(
            content=f"Сохранить {item['filename']} в Загрузки",
            on_click=on_save_click,
            bgcolor="#eef4fc",
            color=PRIMARY_COLOR,
        )

    def populate_save_buttons(row_control, items):
        row_control.controls = [make_export_button(item) for item in items]
        page.update()

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
                show_snack("Разрешение не получено. Используйте кнопку «Сохранить в Загрузки».")
        except Exception as ex:
            log_to_ui(f"Ошибка при запросе разрешения: {ex}")

    # Вкладка 1: Шаблоны
    title_field = make_field("Заголовок", hint_text="Введите заголовок документа")
    text_field = make_field("Текст", multiline=True, min_lines=3, max_lines=5)
    list_field = make_field(
        "Список (каждый пункт с новой строки)",
        multiline=True,
        min_lines=2,
        max_lines=4,
    )
    table_field = make_field(
        "Таблица (CSV: строки через Enter, колонки через запятую)",
        multiline=True,
        min_lines=3,
        max_lines=5,
    )
    filename_field = make_field("Имя файла (с расширением)", value="document.docx")

    def make_create_handler(format_type):
        def handler(e):
            try:
                _pop_recent_created_files()
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
                populate_save_buttons(doc_save_buttons, _pop_recent_created_files())
                show_snack(res)
            except Exception as ex:
                show_snack(f"Ошибка: {ex}")

        return handler

    doc_tab = ft.Container(
        padding=ft.Padding.only(left=16, right=16, top=16, bottom=28),
        expand=True,
        bgcolor=SURFACE_COLOR,
        content=ft.Column(
            controls=[
                ft.Text("Создание документа", size=20, weight=ft.FontWeight.BOLD, color=TEXT_COLOR),
                title_field,
                text_field,
                list_field,
                table_field,
                filename_field,
                ft.Row(
                    controls=[
                        make_elevated_btn("Word (.docx)", make_create_handler("docx")),
                        make_elevated_btn("PowerPoint (.pptx)", make_create_handler("pptx")),
                        make_elevated_btn("Excel (.xlsx)", make_create_handler("xlsx")),
                        make_elevated_btn("PDF (.pdf)", make_create_handler("pdf")),
                        make_elevated_btn("Markdown (.md)", make_create_handler("md")),
                        make_elevated_btn("Текст (.txt)", make_create_handler("txt")),
                    ],
                    wrap=True,
                    spacing=8,
                    run_spacing=8,
                ),
                doc_save_buttons,
                ft.Divider(color=DIVIDER_COLOR, height=1),
                ft.Text(
                    "Все файлы сохраняются в папку Download или через кнопку выше.",
                    size=12,
                    color="#757575",
                ),
                ft.Row(
                    controls=[
                        make_elevated_btn("Запросить разрешение на запись", request_storage_permission),
                    ],
                    wrap=True,
                ),
            ],
            horizontal_alignment=ft.CrossAxisAlignment.STRETCH,
            scroll=ft.ScrollMode.AUTO,
            expand=True,
            spacing=12,
        ),
    )

    # Вкладка 2: Верстка из файла
    dropzone_title = ft.Text(
        "Нажмите, чтобы выбрать текстовый файл / скрипт",
        size=14,
        weight=ft.FontWeight.W_500,
        color=TEXT_COLOR,
        text_align=ft.TextAlign.CENTER,
    )
    layout_content_field = make_field(
        "Содержимое текстового файла / разметка",
        multiline=True,
        min_lines=8,
        max_lines=14,
        hint_text="# Заголовок\nТекст\n- Пункт 1\n| A | B |\n| --- | --- |\n| 1 | 2 |",
    )
    layout_filename_field = make_field(
        "Имя итогового файла (без расширения)",
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
                dropzone_title.value = f"Загружен файл: {f.name}"
                layout_filename_field.value = os.path.splitext(f.name)[0] or "layout"
                layout_log.value = f"Загружен файл: {f.name}\n"
                page.update()

    def load_sample_layout(e):
        loaded_source_name["name"] = "example.md"
        dropzone_title.value = "Загружен файл: example.md"
        layout_content_field.value = SAMPLE_LAYOUT_TEXT
        layout_filename_field.value = "report"
        layout_log.value = "Загружен пример текстовой разметки.\n"
        page.update()

    def make_typeset_handler(fmt):
        def handler(e):
            try:
                _pop_recent_created_files()
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
                populate_save_buttons(layout_save_buttons, _pop_recent_created_files())
                show_snack(res)
            except Exception as ex:
                layout_log.value = f"❌ Ошибка: {ex}\n"
                show_snack(f"Ошибка: {ex}")

        return handler

    def run_layout_as_script(e):
        layout_log.value = "Выполнение...\n"
        layout_save_buttons.controls = []
        page.update()
        execute_user_script(
            layout_content_field.value,
            log_layout,
            on_files_ready=lambda items: populate_save_buttons(layout_save_buttons, items),
        )

    layout_tab = ft.Container(
        padding=ft.Padding.only(left=16, right=16, top=16, bottom=28),
        expand=True,
        bgcolor=SURFACE_COLOR,
        content=ft.Column(
            controls=[
                ft.Text(
                    "Верстка из текстовых файлов и скриптов",
                    size=20,
                    weight=ft.FontWeight.BOLD,
                    color=TEXT_COLOR,
                ),
                ft.Container(
                    bgcolor=DROPZONE_BG,
                    border=ft.Border.all(1, BORDER_COLOR),
                    border_radius=8,
                    padding=16,
                    on_click=pick_layout_file,
                    content=ft.Column(
                        controls=[
                            dropzone_title,
                            ft.Text(
                                "Поддерживаются: .txt, .md, .py, .js, .json, .csv, .yaml",
                                size=12,
                                color=MUTED_COLOR,
                                text_align=ft.TextAlign.CENTER,
                            ),
                        ],
                        horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                        spacing=6,
                    ),
                ),
                ft.Row(
                    controls=[
                        make_outlined_btn("Выбрать файл", pick_layout_file),
                        make_outlined_btn("Загрузить пример разметки", load_sample_layout),
                    ],
                    wrap=True,
                    spacing=8,
                    run_spacing=8,
                ),
                layout_content_field,
                layout_filename_field,
                ft.Row(
                    controls=[
                        make_elevated_btn("Word (.docx)", make_typeset_handler("docx")),
                        make_elevated_btn("PDF (.pdf)", make_typeset_handler("pdf")),
                        make_elevated_btn("PowerPoint (.pptx)", make_typeset_handler("pptx")),
                        make_elevated_btn("Excel (.xlsx)", make_typeset_handler("xlsx")),
                        make_elevated_btn("Markdown (.md)", make_typeset_handler("md")),
                        make_elevated_btn("HTML (.html)", make_typeset_handler("html")),
                        make_elevated_btn("Текст (.txt)", make_typeset_handler("txt")),
                        make_outlined_btn("Выполнить как скрипт", run_layout_as_script),
                    ],
                    wrap=True,
                    spacing=8,
                    run_spacing=8,
                ),
                ft.Divider(color=DIVIDER_COLOR, height=1),
                ft.Text("Результат верстки:", size=14, weight=ft.FontWeight.BOLD, color=TEXT_COLOR),
                ft.Container(
                    bgcolor=LOG_BG,
                    border=ft.Border.all(1, DIVIDER_COLOR),
                    border_radius=6,
                    padding=10,
                    content=layout_log,
                ),
                layout_save_buttons,
            ],
            horizontal_alignment=ft.CrossAxisAlignment.STRETCH,
            scroll=ft.ScrollMode.AUTO,
            expand=True,
            spacing=12,
        ),
    )

    # Вкладка 3: Свой скрипт
    script_code_field = make_field(
        "Python-код",
        multiline=True,
        min_lines=10,
        max_lines=18,
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
        script_save_buttons.controls = []
        page.update()
        execute_user_script(
            script_code_field.value,
            log_to_ui,
            on_files_ready=lambda items: populate_save_buttons(script_save_buttons, items),
        )

    script_tab = ft.Container(
        padding=ft.Padding.only(left=16, right=16, top=16, bottom=28),
        expand=True,
        bgcolor=SURFACE_COLOR,
        content=ft.Column(
            controls=[
                ft.Text("Свой скрипт", size=20, weight=ft.FontWeight.BOLD, color=TEXT_COLOR),
                script_code_field,
                ft.Row(
                    controls=[
                        make_elevated_btn("Выполнить", run_script),
                        make_outlined_btn("Загрузить файл скрипта (.txt, .py, .js)", pick_script_file),
                    ],
                    wrap=True,
                    spacing=8,
                    run_spacing=8,
                ),
                ft.Divider(color=DIVIDER_COLOR, height=1),
                ft.Text("Лог выполнения:", size=14, weight=ft.FontWeight.BOLD, color=TEXT_COLOR),
                ft.Container(
                    bgcolor=LOG_BG,
                    border=ft.Border.all(1, DIVIDER_COLOR),
                    border_radius=6,
                    padding=10,
                    content=script_log,
                ),
                script_save_buttons,
            ],
            horizontal_alignment=ft.CrossAxisAlignment.STRETCH,
            scroll=ft.ScrollMode.AUTO,
            expand=True,
            spacing=12,
        ),
    )

    tabs = ft.Tabs(
        length=3,
        selected_index=0,
        expand=True,
        content=ft.Column(
            expand=True,
            spacing=0,
            controls=[
                ft.TabBar(
                    indicator_color=PRIMARY_COLOR,
                    label_color=PRIMARY_COLOR,
                    unselected_label_color=MUTED_COLOR,
                    tabs=[
                        ft.Tab(label="Шаблоны"),
                        ft.Tab(label="Верстка из файла"),
                        ft.Tab(label="Свой скрипт"),
                    ],
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
