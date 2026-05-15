import React, { ChangeEvent, useMemo, useState } from "react";
import { createRoot } from "react-dom/client";
import * as XLSX from "xlsx";
import {
  Download,
  FileSpreadsheet,
  Files,
  RefreshCw,
  Table2,
  Trash2,
  Upload,
} from "lucide-react";
import "./styles.css";

type CellValue = string | number | boolean | Date | null;

type SheetData = {
  name: string;
  rows: CellValue[][];
};

type WorkbookFile = {
  id: string;
  fileName: string;
  sheets: SheetData[];
  activeSheet: string;
  headerRow: number;
};

type MergeResult = {
  headers: string[];
  rows: Record<string, CellValue>[];
};

const MAX_PREVIEW_ROWS = 8;

function isBlankCell(value: CellValue) {
  return value === null || String(value).trim() === "";
}

function normalizeCell(cell: XLSX.CellObject | undefined): CellValue {
  if (!cell || cell.v === undefined || cell.v === null) return null;
  if (cell.v instanceof Date) return cell.v;

  if (cell.t === "n" && typeof cell.v === "number") {
    const displayedValue = typeof cell.w === "string" ? cell.w.trim() : "";
    const compactDisplay = displayedValue.replace(/[, ]/g, "");

    if (/^0\d+/.test(compactDisplay) || /^\d{12,}$/.test(compactDisplay) || /e\+/i.test(displayedValue)) {
      return displayedValue || String(cell.v);
    }

    return cell.v;
  }

  if (typeof cell.v === "string" || typeof cell.v === "boolean") {
    return cell.v;
  }

  return String(cell.v);
}

function readSheetRows(sheet: XLSX.WorkSheet): CellValue[][] {
  if (!sheet["!ref"]) return [];

  const range = XLSX.utils.decode_range(sheet["!ref"]);
  const rows: CellValue[][] = [];

  for (let rowIndex = range.s.r; rowIndex <= range.e.r; rowIndex += 1) {
    const row: CellValue[] = [];

    for (let columnIndex = range.s.c; columnIndex <= range.e.c; columnIndex += 1) {
      const address = XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex });
      row.push(normalizeCell(sheet[address]));
    }

    if (row.some((value) => !isBlankCell(value))) {
      rows.push(row);
    }
  }

  return rows;
}

function formatCellValue(value: CellValue) {
  if (value === null) return "";
  if (value instanceof Date) return value.toLocaleDateString("zh-CN");
  return String(value);
}

function makeHeaderName(value: CellValue, index: number) {
  const raw = value === null ? "" : String(value).trim();
  return raw || `未命名列 ${index + 1}`;
}

function getActiveSheet(file: WorkbookFile) {
  return file.sheets.find((sheet) => sheet.name === file.activeSheet) ?? file.sheets[0];
}

function mergeWorkbooks(files: WorkbookFile[]): MergeResult {
  const headers: string[] = [];
  const headerSet = new Set<string>();
  const rows: Record<string, CellValue>[] = [];

  for (const file of files) {
    const sheet = getActiveSheet(file);
    if (!sheet) continue;

    const headerIndex = Math.max(0, file.headerRow - 1);
    const sourceHeaders = (sheet.rows[headerIndex] ?? []).map(makeHeaderName);

    for (const header of sourceHeaders) {
      if (!headerSet.has(header)) {
        headerSet.add(header);
        headers.push(header);
      }
    }

    for (const dataRow of sheet.rows.slice(headerIndex + 1)) {
      const hasValue = dataRow.some((value) => !isBlankCell(value));
      if (!hasValue) continue;

      const row: Record<string, CellValue> = {
        来源文件: file.fileName,
        来源工作表: sheet.name,
      };

      for (let index = 0; index < sourceHeaders.length; index += 1) {
        row[sourceHeaders[index]] = dataRow[index] ?? null;
      }

      rows.push(row);
    }
  }

  const finalHeaders = ["来源文件", "来源工作表", ...headers.filter((header) => !["来源文件", "来源工作表"].includes(header))];
  return { headers: finalHeaders, rows };
}

async function readWorkbook(file: File): Promise<WorkbookFile> {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { cellDates: true });
  const sheets = workbook.SheetNames.map((sheetName) => {
    const sheet = workbook.Sheets[sheetName];

    return {
      name: sheetName,
      rows: readSheetRows(sheet),
    };
  });

  return {
    id: `${file.name}-${file.lastModified}-${crypto.randomUUID()}`,
    fileName: file.name,
    sheets,
    activeSheet: sheets[0]?.name ?? "",
    headerRow: 1,
  };
}

function App() {
  const [files, setFiles] = useState<WorkbookFile[]>([]);
  const [isReading, setIsReading] = useState(false);
  const [error, setError] = useState("");

  const mergeResult = useMemo(() => mergeWorkbooks(files), [files]);

  async function handleFiles(event: ChangeEvent<HTMLInputElement>) {
    const selectedFiles = Array.from(event.target.files ?? []);
    if (selectedFiles.length === 0) return;

    setIsReading(true);
    setError("");

    try {
      const parsed = await Promise.all(selectedFiles.map(readWorkbook));
      setFiles((current) => [...current, ...parsed]);
    } catch (err) {
      setError(err instanceof Error ? err.message : "读取文件失败，请确认文件格式正确。");
    } finally {
      setIsReading(false);
      event.target.value = "";
    }
  }

  function updateFile(id: string, next: Partial<WorkbookFile>) {
    setFiles((current) => current.map((file) => (file.id === id ? { ...file, ...next } : file)));
  }

  function removeFile(id: string) {
    setFiles((current) => current.filter((file) => file.id !== id));
  }

  function clearAll() {
    setFiles([]);
    setError("");
  }

  function exportMergedWorkbook() {
    if (mergeResult.rows.length === 0) {
      setError("还没有可导出的数据。");
      return;
    }

    const exportRows = mergeResult.rows.map((row) =>
      mergeResult.headers.reduce<Record<string, CellValue>>((nextRow, header) => {
        nextRow[header] = row[header] ?? null;
        return nextRow;
      }, {}),
    );

    const worksheet = XLSX.utils.json_to_sheet(exportRows, { header: mergeResult.headers });
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "总表");
    XLSX.writeFile(workbook, `总表-${new Date().toISOString().slice(0, 10)}.xlsx`);
  }

  return (
    <main className="app">
      <header className="topbar">
        <div>
          <div className="eyebrow">本地 Excel 数据处理</div>
          <h1>Excel 总表合并工具</h1>
        </div>
        <div className="actions">
          <label className="button primary">
            <Upload size={18} />
            <span>{isReading ? "读取中" : "上传 Excel"}</span>
            <input
              type="file"
              multiple
              accept=".xlsx,.xls,.csv"
              onChange={handleFiles}
              disabled={isReading}
            />
          </label>
          <button className="button" onClick={exportMergedWorkbook} disabled={mergeResult.rows.length === 0}>
            <Download size={18} />
            <span>导出总表</span>
          </button>
        </div>
      </header>

      <section className="stats" aria-label="处理概览">
        <div className="stat">
          <Files size={20} />
          <div>
            <strong>{files.length}</strong>
            <span>文件</span>
          </div>
        </div>
        <div className="stat">
          <Table2 size={20} />
          <div>
            <strong>{mergeResult.rows.length}</strong>
            <span>合并行</span>
          </div>
        </div>
        <div className="stat">
          <FileSpreadsheet size={20} />
          <div>
            <strong>{mergeResult.headers.length}</strong>
            <span>总表列</span>
          </div>
        </div>
        <button className="button ghost" onClick={clearAll} disabled={files.length === 0}>
          <RefreshCw size={18} />
          <span>清空</span>
        </button>
      </section>

      {error && <div className="notice">{error}</div>}

      {files.length === 0 ? (
        <section className="empty">
          <FileSpreadsheet size={56} />
          <h2>选择几个 Excel 文件开始合并</h2>
          <p>每个文件可以选择工作表和表头行，导出时会自动补齐不同文件里的列。</p>
        </section>
      ) : (
        <section className="workspace">
          <div className="file-list">
            {files.map((file) => {
              const activeSheet = getActiveSheet(file);
              const previewRows = activeSheet?.rows.slice(0, MAX_PREVIEW_ROWS) ?? [];
              const maxHeaderRow = Math.max(1, activeSheet?.rows.length ?? 1);

              return (
                <article className="file-card" key={file.id}>
                  <div className="file-card-header">
                    <div>
                      <h2>{file.fileName}</h2>
                      <p>{activeSheet?.rows.length ?? 0} 行数据</p>
                    </div>
                    <button className="icon-button" onClick={() => removeFile(file.id)} aria-label="删除文件">
                      <Trash2 size={18} />
                    </button>
                  </div>

                  <div className="controls">
                    <label>
                      <span>工作表</span>
                      <select
                        value={file.activeSheet}
                        onChange={(event) => updateFile(file.id, { activeSheet: event.target.value, headerRow: 1 })}
                      >
                        {file.sheets.map((sheet) => (
                          <option value={sheet.name} key={sheet.name}>
                            {sheet.name}
                          </option>
                        ))}
                      </select>
                    </label>
                    <label>
                      <span>表头行</span>
                      <input
                        type="number"
                        min={1}
                        max={maxHeaderRow}
                        value={file.headerRow}
                        onChange={(event) => updateFile(file.id, { headerRow: Number(event.target.value) || 1 })}
                      />
                    </label>
                  </div>

                  <div className="table-wrap">
                    <table>
                      <tbody>
                        {previewRows.map((row, rowIndex) => (
                          <tr className={rowIndex + 1 === file.headerRow ? "selected-row" : ""} key={rowIndex}>
                            <th>{rowIndex + 1}</th>
                            {row.slice(0, 8).map((cell, cellIndex) => (
                              <td key={cellIndex}>{formatCellValue(cell)}</td>
                            ))}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </article>
              );
            })}
          </div>

          <aside className="summary">
            <h2>总表预览</h2>
            <p>{mergeResult.rows.length} 行会被导出，包含来源文件和来源工作表两列。</p>
            <div className="table-wrap compact">
              <table>
                <thead>
                  <tr>
                    {mergeResult.headers.slice(0, 8).map((header) => (
                      <th key={header}>{header}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {mergeResult.rows.slice(0, 10).map((row, rowIndex) => (
                    <tr key={rowIndex}>
                      {mergeResult.headers.slice(0, 8).map((header) => (
                        <td key={header}>{formatCellValue(row[header] ?? null)}</td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </aside>
        </section>
      )}
    </main>
  );
}

createRoot(document.getElementById("root")!).render(<App />);
