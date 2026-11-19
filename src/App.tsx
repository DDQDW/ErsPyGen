import { useState, useRef, type ChangeEvent, type DragEvent } from 'react';

type DocFormat = 'docx' | 'pptx' | 'xlsx' | 'pdf' | 'txt' | 'md' | 'html';

interface GeneratedFileLink {
  filename: string;
  downloadUrl: string;
}

const DEFAULT_SCRIPT = `# Пример:
# create_docx('Заголовок', 'Текст', 'Пункт 1\\nПункт 2', 'A,B\\n1,2', 'my.docx')
# create_from_text('# Заголовок\\n- Пункт 1\\n- Пункт 2', 'pdf', 'layout.pdf')
print('Привет из скрипта!')`;

const SAMPLE_LAYOUT_TEXT = `# Отчет по проекту DocBuilder
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
| Скрипты | Готово | Python/JS |`;

const ACCEPTED_TEXT_EXTENSIONS =
  '.txt,.md,.py,.js,.ts,.json,.csv,.yaml,.yml,.ini,.log,.rst,text/*';

function triggerDownload(url: string, filename: string) {
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}

function stripExt(name: string): string {
  const idx = name.lastIndexOf('.');
  return idx > 0 ? name.slice(0, idx) : name;
}

export default function App() {
  const [activeTab, setActiveTab] = useState<'templates' | 'layout' | 'script'>('templates');

  // Templates tab state
  const [title, setTitle] = useState('');
  const [text, setText] = useState('');
  const [listItems, setListItems] = useState('');
  const [tableCsv, setTableCsv] = useState('');
  const [filename, setFilename] = useState('document.docx');
  const [isGenerating, setIsGenerating] = useState(false);

  // Text-file layout tab state
  const [sourceFilename, setSourceFilename] = useState('');
  const [layoutContent, setLayoutContent] = useState('');
  const [layoutOutFilename, setLayoutOutFilename] = useState('layout');
  const [layoutLog, setLayoutLog] = useState('');
  const [layoutDownloads, setLayoutDownloads] = useState<GeneratedFileLink[]>([]);
  const [isDragging, setIsDragging] = useState(false);
  const [isTypesetting, setIsTypesetting] = useState(false);
  const layoutFileInputRef = useRef<HTMLInputElement | null>(null);

  // Script tab state
  const [scriptCode, setScriptCode] = useState(DEFAULT_SCRIPT);
  const [scriptLog, setScriptLog] = useState('');
  const [scriptDownloads, setScriptDownloads] = useState<GeneratedFileLink[]>([]);
  const [isRunningScript, setIsRunningScript] = useState(false);
  const scriptFileInputRef = useRef<HTMLInputElement | null>(null);

  // Snackbar state
  const [snackbar, setSnackbar] = useState<{
    message: string;
    downloadUrl?: string;
    filename?: string;
  } | null>(null);
  const snackTimerRef = useRef<number | null>(null);

  const showSnackbar = (message: string, downloadUrl?: string, dlFilename?: string) => {
    setSnackbar({ message, downloadUrl, filename: dlFilename });
    if (snackTimerRef.current !== null) {
      window.clearTimeout(snackTimerRef.current);
    }
    snackTimerRef.current = window.setTimeout(() => {
      setSnackbar(null);
    }, 5000);
  };

  const handleGenerate = async (format: DocFormat) => {
    setIsGenerating(true);
    try {
      const response = await fetch('/api/generate', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          format,
          title,
          text,
          listItems,
          tableCsv,
          filename,
        }),
      });
      const data = await response.json();
      if (!response.ok || !data.ok) {
        throw new Error(data.error || 'Не удалось создать документ.');
      }

      triggerDownload(data.downloadUrl, data.filename);
      showSnackbar(data.message, data.downloadUrl, data.filename);
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      showSnackbar(`Ошибка: ${message}`);
    } finally {
      setIsGenerating(false);
    }
  };

  const handleRequestStoragePermission = () => {
    const msg = 'Разрешение на управление внешним хранилищем получено.';
    setScriptLog((prev) => `${prev}${msg}\n`);
    showSnackbar('Браузер готов сохранять файлы в папку Download.');
  };

  const loadFileIntoLayout = async (file: File) => {
    const content = await file.text();
    setSourceFilename(file.name);
    setLayoutContent(content);
    setLayoutOutFilename(stripExt(file.name) || 'layout');
    setLayoutLog(`Загружен файл: ${file.name} (${content.length} симв.)\n`);
    showSnackbar(`Файл ${file.name} загружен`);
  };

  const handleLayoutFileChange = async (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      await loadFileIntoLayout(file);
    }
    e.target.value = '';
  };

  const handleLayoutDrop = async (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragging(false);
    const file = e.dataTransfer.files?.[0];
    if (file) {
      await loadFileIntoLayout(file);
    }
  };

  const handleTypesetFromText = async (format: DocFormat) => {
    setIsTypesetting(true);
    setLayoutDownloads([]);
    try {
      const baseName = stripExt(layoutOutFilename.trim() || 'layout');
      const targetName = `${baseName}.${format}`;
      const response = await fetch('/api/typeset-text', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          content: layoutContent,
          sourceFilename,
          format,
          filename: targetName,
        }),
      });
      const data = await response.json();
      if (!response.ok || !data.ok) {
        throw new Error(data.error || 'Не удалось сверстать документ.');
      }

      const summary = data.parsed
        ? `Заголовок: "${data.parsed.title || '—'}", пунктов списка: ${data.parsed.listCount}, строк таблицы: ${data.parsed.tableRows}`
        : '';
      setLayoutLog(`${data.message}\n${summary}\n`);
      setLayoutDownloads([{ filename: data.filename, downloadUrl: data.downloadUrl }]);
      triggerDownload(data.downloadUrl, data.filename);
      showSnackbar(data.message, data.downloadUrl, data.filename);
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      setLayoutLog((prev) => `${prev}❌ Ошибка: ${message}\n`);
      showSnackbar(`Ошибка: ${message}`);
    } finally {
      setIsTypesetting(false);
    }
  };

  const handleRunLayoutAsScript = async () => {
    if (!layoutContent.trim()) {
      showSnackbar('Загрузите или введите текст скрипта.');
      return;
    }
    setIsTypesetting(true);
    setLayoutLog('Выполнение как скрипт...\n');
    setLayoutDownloads([]);
    try {
      const response = await fetch('/api/run-script', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ code: layoutContent }),
      });
      const data = await response.json();
      const output = Array.isArray(data.logs) ? data.logs.join('\n') + '\n' : '';
      setLayoutLog(output);

      if (Array.isArray(data.createdFiles) && data.createdFiles.length > 0) {
        setLayoutDownloads(data.createdFiles);
        for (const f of data.createdFiles) {
          triggerDownload(f.downloadUrl, f.filename);
        }
      }
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      setLayoutLog((prev) => `${prev}❌ Ошибка: ${message}\n`);
    } finally {
      setIsTypesetting(false);
    }
  };

  const handleScriptFileUpload = async (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      const content = await file.text();
      setScriptCode(content);
      setScriptLog(`Загружен файл скрипта: ${file.name}\n`);
      showSnackbar(`Скрипт ${file.name} загружен в редактор`);
    }
    e.target.value = '';
  };

  const handleRunScript = async () => {
    setIsRunningScript(true);
    setScriptLog('Выполнение...\n');
    setScriptDownloads([]);

    try {
      const response = await fetch('/api/run-script', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ code: scriptCode }),
      });
      const data = await response.json();
      const output = Array.isArray(data.logs) ? data.logs.join('\n') + '\n' : '';
      setScriptLog(output);

      if (Array.isArray(data.createdFiles) && data.createdFiles.length > 0) {
        setScriptDownloads(data.createdFiles);
        for (const f of data.createdFiles) {
          triggerDownload(f.downloadUrl, f.filename);
        }
      }
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      setScriptLog((prev) => `${prev}❌ Ошибка: ${message}\n`);
    } finally {
      setIsRunningScript(false);
    }
  };

  return (
    <div className="app-shell">
      <nav className="tabs-bar" role="tablist">
        <button
          type="button"
          role="tab"
          aria-selected={activeTab === 'templates'}
          className={`tab-btn ${activeTab === 'templates' ? 'active' : ''}`}
          onClick={() => setActiveTab('templates')}
        >
          Шаблоны
        </button>
        <button
          type="button"
          role="tab"
          aria-selected={activeTab === 'layout'}
          className={`tab-btn ${activeTab === 'layout' ? 'active' : ''}`}
          onClick={() => setActiveTab('layout')}
        >
          Верстка из файла
        </button>
        <button
          type="button"
          role="tab"
          aria-selected={activeTab === 'script'}
          className={`tab-btn ${activeTab === 'script' ? 'active' : ''}`}
          onClick={() => setActiveTab('script')}
        >
          Свой скрипт
        </button>
      </nav>

      {activeTab === 'templates' && (
        <main className="tab-content">
          <h1 className="section-title">Создание документа</h1>

          <div className="field-group">
            <label className="field-label" htmlFor="title-field">
              Заголовок
            </label>
            <input
              id="title-field"
              className="field-input"
              type="text"
              placeholder="Введите заголовок документа"
              value={title}
              onChange={(e) => setTitle(e.target.value)}
            />
          </div>

          <div className="field-group">
            <label className="field-label" htmlFor="text-field">
              Текст
            </label>
            <textarea
              id="text-field"
              className="field-textarea"
              rows={4}
              value={text}
              onChange={(e) => setText(e.target.value)}
            />
          </div>

          <div className="field-group">
            <label className="field-label" htmlFor="list-field">
              Список (каждый пункт с новой строки)
            </label>
            <textarea
              id="list-field"
              className="field-textarea"
              rows={3}
              value={listItems}
              onChange={(e) => setListItems(e.target.value)}
            />
          </div>

          <div className="field-group">
            <label className="field-label" htmlFor="table-field">
              Таблица (CSV: строки через Enter, колонки через запятую)
            </label>
            <textarea
              id="table-field"
              className="field-textarea"
              rows={4}
              value={tableCsv}
              onChange={(e) => setTableCsv(e.target.value)}
            />
          </div>

          <div className="field-group">
            <label className="field-label" htmlFor="filename-field">
              Имя файла (с расширением)
            </label>
            <input
              id="filename-field"
              className="field-input"
              type="text"
              value={filename}
              onChange={(e) => setFilename(e.target.value)}
            />
          </div>

          <div className="buttons-row">
            <button
              type="button"
              className="elevated-btn"
              disabled={isGenerating}
              onClick={() => handleGenerate('docx')}
            >
              Word (.docx)
            </button>
            <button
              type="button"
              className="elevated-btn"
              disabled={isGenerating}
              onClick={() => handleGenerate('pptx')}
            >
              PowerPoint (.pptx)
            </button>
            <button
              type="button"
              className="elevated-btn"
              disabled={isGenerating}
              onClick={() => handleGenerate('xlsx')}
            >
              Excel (.xlsx)
            </button>
            <button
              type="button"
              className="elevated-btn"
              disabled={isGenerating}
              onClick={() => handleGenerate('pdf')}
            >
              PDF (.pdf)
            </button>
            <button
              type="button"
              className="elevated-btn"
              disabled={isGenerating}
              onClick={() => handleGenerate('md')}
            >
              Markdown (.md)
            </button>
            <button
              type="button"
              className="elevated-btn"
              disabled={isGenerating}
              onClick={() => handleGenerate('txt')}
            >
              Текст (.txt)
            </button>
          </div>

          <hr className="divider" />

          <p className="hint-text">Все файлы сохраняются в папку Download.</p>
          <div className="buttons-row">
            <button
              type="button"
              className="elevated-btn"
              onClick={handleRequestStoragePermission}
            >
              Запросить разрешение на запись
            </button>
          </div>
        </main>
      )}

      {activeTab === 'layout' && (
        <main className="tab-content">
          <h1 className="section-title">Верстка из текстовых файлов и скриптов</h1>

          <input
            ref={layoutFileInputRef}
            type="file"
            accept={ACCEPTED_TEXT_EXTENSIONS}
            className="hidden-file-input"
            onChange={handleLayoutFileChange}
          />

          <div
            className={`dropzone ${isDragging ? 'drag-over' : ''}`}
            onClick={() => layoutFileInputRef.current?.click()}
            onDragOver={(e) => {
              e.preventDefault();
              setIsDragging(true);
            }}
            onDragLeave={() => setIsDragging(false)}
            onDrop={handleLayoutDrop}
            role="button"
            tabIndex={0}
            onKeyDown={(e) => {
              if (e.key === 'Enter' || e.key === ' ') {
                e.preventDefault();
                layoutFileInputRef.current?.click();
              }
            }}
          >
            <p className="dropzone-title">
              {sourceFilename
                ? `Загружен файл: ${sourceFilename}`
                : 'Нажмите или перетащите текстовый файл / скрипт'}
            </p>
            <p className="dropzone-sub">
              Поддерживаются: .txt, .md, .py, .js, .json, .csv, .yaml
            </p>
          </div>

          <div className="buttons-row">
            <button
              type="button"
              className="outlined-btn"
              onClick={() => layoutFileInputRef.current?.click()}
            >
              Выбрать файл
            </button>
            <button
              type="button"
              className="outlined-btn"
              onClick={() => {
                setSourceFilename('example.md');
                setLayoutContent(SAMPLE_LAYOUT_TEXT);
                setLayoutOutFilename('report');
                setLayoutLog('Загружен пример текстовой разметки.\n');
              }}
            >
              Загрузить пример разметки
            </button>
          </div>

          <div className="field-group">
            <label className="field-label" htmlFor="layout-content-field">
              Содержимое файла (разметка: # заголовок, - список, | таблица |, JSON, CSV или код скрипта)
            </label>
            <textarea
              id="layout-content-field"
              className="field-textarea code"
              rows={10}
              placeholder="Вставьте разметку, CSV, JSON или загрузите файл выше..."
              value={layoutContent}
              onChange={(e) => setLayoutContent(e.target.value)}
            />
          </div>

          <div className="field-group">
            <label className="field-label" htmlFor="layout-filename-field">
              Имя итогового файла (без расширения)
            </label>
            <input
              id="layout-filename-field"
              className="field-input"
              type="text"
              value={layoutOutFilename}
              onChange={(e) => setLayoutOutFilename(e.target.value)}
            />
          </div>

          <div className="buttons-row">
            <button
              type="button"
              className="elevated-btn"
              disabled={isTypesetting}
              onClick={() => handleTypesetFromText('docx')}
            >
              Word (.docx)
            </button>
            <button
              type="button"
              className="elevated-btn"
              disabled={isTypesetting}
              onClick={() => handleTypesetFromText('pdf')}
            >
              PDF (.pdf)
            </button>
            <button
              type="button"
              className="elevated-btn"
              disabled={isTypesetting}
              onClick={() => handleTypesetFromText('pptx')}
            >
              PowerPoint (.pptx)
            </button>
            <button
              type="button"
              className="elevated-btn"
              disabled={isTypesetting}
              onClick={() => handleTypesetFromText('xlsx')}
            >
              Excel (.xlsx)
            </button>
            <button
              type="button"
              className="elevated-btn"
              disabled={isTypesetting}
              onClick={() => handleTypesetFromText('md')}
            >
              Markdown (.md)
            </button>
            <button
              type="button"
              className="elevated-btn"
              disabled={isTypesetting}
              onClick={() => handleTypesetFromText('html')}
            >
              HTML (.html)
            </button>
            <button
              type="button"
              className="elevated-btn"
              disabled={isTypesetting}
              onClick={() => handleTypesetFromText('txt')}
            >
              Текст (.txt)
            </button>
            <button
              type="button"
              className="outlined-btn"
              disabled={isTypesetting}
              onClick={handleRunLayoutAsScript}
            >
              Выполнить как скрипт
            </button>
          </div>

          <hr className="divider" />

          <p className="log-label">Результат верстки:</p>
          <pre className="log-output">{layoutLog}</pre>

          {layoutDownloads.length > 0 && (
            <div className="downloads-list">
              {layoutDownloads.map((item) => (
                <a
                  key={item.downloadUrl}
                  className="download-link"
                  href={item.downloadUrl}
                  download={item.filename}
                >
                  Скачать {item.filename}
                </a>
              ))}
            </div>
          )}
        </main>
      )}

      {activeTab === 'script' && (
        <main className="tab-content">
          <h1 className="section-title">Свой скрипт</h1>

          <input
            ref={scriptFileInputRef}
            type="file"
            accept={ACCEPTED_TEXT_EXTENSIONS}
            className="hidden-file-input"
            onChange={handleScriptFileUpload}
          />

          <div className="field-group">
            <label className="field-label" htmlFor="script-code-field">
              Python / JS-код
            </label>
            <textarea
              id="script-code-field"
              className="field-textarea code"
              rows={12}
              placeholder="Доступны: create_docx, create_pptx, create_xlsx, create_pdf, create_from_text, create_txt, create_md, create_html"
              value={scriptCode}
              onChange={(e) => setScriptCode(e.target.value)}
            />
          </div>

          <div className="buttons-row">
            <button
              type="button"
              className="elevated-btn"
              disabled={isRunningScript}
              onClick={handleRunScript}
            >
              Выполнить
            </button>
            <button
              type="button"
              className="outlined-btn"
              onClick={() => scriptFileInputRef.current?.click()}
            >
              Загрузить файл скрипта (.txt, .py, .js)
            </button>
          </div>

          <hr className="divider" />

          <p className="log-label">Лог выполнения:</p>
          <pre className="log-output">{scriptLog}</pre>

          {scriptDownloads.length > 0 && (
            <div className="downloads-list">
              {scriptDownloads.map((item) => (
                <a
                  key={item.downloadUrl}
                  className="download-link"
                  href={item.downloadUrl}
                  download={item.filename}
                >
                  Скачать {item.filename}
                </a>
              ))}
            </div>
          )}
        </main>
      )}

      {snackbar && (
        <div className="snackbar" role="status">
          <span>{snackbar.message}</span>
          {snackbar.downloadUrl && snackbar.filename && (
            <a href={snackbar.downloadUrl} download={snackbar.filename}>
              Скачать
            </a>
          )}
        </div>
      )}
    </div>
  );
}
