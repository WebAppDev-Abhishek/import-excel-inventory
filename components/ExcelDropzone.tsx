"use client";

import { useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx';
import { InventoryCardGrid } from '@/components/InventoryCardGrid';

type RowData = Record<string, string | number | boolean | null>;

type ParsedWorkbook = {
  headers: string[];
  rows: RowData[];
};

const commonNameKeys = ['item name', 'name', 'product', 'product name', 'title'];
const commonStockKeys = ['stock', 'quantity', 'available', 'qty', 'on hand', 'available quantity'];
const commonStatusKeys = ['status', 'state', 'availability'];

function normalize(value: string) {
  return value.trim().toLowerCase();
}

function getFieldValue(row: RowData, keys: string[]) {
  for (const key of keys) {
    const match = Object.keys(row).find((column) => normalize(column) === normalize(key));
    if (match) {
      const value = row[match];
      if (value === null || value === undefined) return null;
      return String(value);
    }
  }
  return null;
}

function sanitizeCellValue(value: unknown): string | number | boolean | null {
  if (value === null || value === undefined) return null;
  if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') {
    return value;
  }
  return String(value);
}

function inferDisplayMeta(row: RowData) {
  const title = getFieldValue(row, commonNameKeys) ?? getFieldValue(row, ['item', 'sku']) ?? 'Inventory item';
  const stockValue = getFieldValue(row, commonStockKeys);
  const statusValue = getFieldValue(row, commonStatusKeys);
  const stockNumber = stockValue ? Number(stockValue) : NaN;

  let badge: { tone: 'default' | 'low' | 'out' | 'info'; label: string } | null = null;
  if (!Number.isNaN(stockNumber)) {
    if (stockNumber <= 0) {
      badge = { tone: 'out', label: 'Out of stock' };
    } else if (stockNumber <= 5) {
      badge = { tone: 'low', label: 'Low stock' };
    }
  } else if (statusValue) {
    badge = { tone: 'info', label: String(statusValue) };
  }

  return { title, badge };
}

function parseWorkbook(file: File): Promise<ParsedWorkbook> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const data = reader.result as ArrayBuffer;
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 }) as unknown[][];

        if (!rows.length) {
          resolve({ headers: [], rows: [] });
          return;
        }

        const headers = (rows[0] as unknown[]).map((header) => String(header ?? '').trim());
        const bodyRows = rows.slice(1);

        const mappedRows = bodyRows
          .filter((row) => row.some((cell) => cell !== undefined && cell !== null && String(cell).trim() !== ''))
          .map((row) => {
            const record: RowData = {};
            headers.forEach((header, index) => {
              const value = row[index];
              record[header || `Column ${index + 1}`] = sanitizeCellValue(value);
            });
            return record;
          });

        resolve({ headers, rows: mappedRows });
      } catch (error) {
        reject(error);
      }
    };
    reader.onerror = () => reject(new Error('Unable to read this file.'));
    reader.readAsArrayBuffer(file);
  });
}

export function ExcelDropzone() {
  const [headers, setHeaders] = useState<string[]>([]);
  const [rows, setRows] = useState<RowData[]>([]);
  const [status, setStatus] = useState('Drop an .xlsx or .xls file to begin.');
  const [error, setError] = useState<string | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const [isLoading, setIsLoading] = useState(false);
  const [search, setSearch] = useState('');
  const inputRef = useRef<HTMLInputElement | null>(null);
  const [exportFormat, setExportFormat] = useState<'xlsx' | 'csv' | 'json'>('xlsx');
  const [lastFileName, setLastFileName] = useState<string | null>(null);

  const filteredRows = useMemo(() => {
    if (!search.trim()) return rows;
    const query = search.trim().toLowerCase();
    return rows.filter((row) =>
      Object.values(row).some((value) => String(value ?? '').toLowerCase().includes(query))
    );
  }, [rows, search]);

  const handleFile = async (file: File | null | undefined) => {
    if (!file) return;
    const extension = file.name.split('.').pop()?.toLowerCase();
    if (extension !== 'xlsx' && extension !== 'xls') {
      setError('Only .xlsx and .xls files are supported.');
      setStatus('Please choose a spreadsheet file.');
      return;
    }

    setIsLoading(true);
    setError(null);
    setStatus(`Parsing ${file.name}...`);

    try {
      const parsed = await parseWorkbook(file);
      if (!parsed.rows.length) {
        setHeaders([]);
        setRows([]);
        setStatus('The workbook was loaded, but no data rows were found.');
      } else {
        setHeaders(parsed.headers);
        setRows(parsed.rows);
        setLastFileName(file.name);
        setStatus(`${parsed.rows.length} items loaded from ${file.name}`);
      }
    } catch (err) {
      setHeaders([]);
      setRows([]);
      setError(err instanceof Error ? err.message : 'Unknown parsing error.');
      setStatus('The file could not be parsed.');
    } finally {
      setIsLoading(false);
    }
  };

  const onDrop = async (event: React.DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    setIsDragging(false);
    const droppedFile = event.dataTransfer.files?.[0];
    await handleFile(droppedFile);
  };

  const onBrowseClick = () => inputRef.current?.click();

  const resetState = () => {
    setHeaders([]);
    setRows([]);
    setSearch('');
    setError(null);
    setStatus('Drop an .xlsx or .xls file to begin.');
    if (inputRef.current) inputRef.current.value = '';
  };

  const cards = filteredRows.map((row, index) => ({
    id: `${index}-${JSON.stringify(row)}`,
    row,
    meta: inferDisplayMeta(row)
  }));

  function buildExportRows() {
    // Convert rows to plain objects ordered by headers
    return rows.map((r) => {
      const obj: Record<string, string | number | boolean | null> = {};
      headers.forEach((h) => {
        obj[h] = r[h] ?? '';
      });
      return obj;
    });
  }

  const downloadParsed = (format: 'xlsx' | 'csv' | 'json') => {
    if (!rows.length) return;
    const baseName = (lastFileName && lastFileName.replace(/\.[^/.]+$/, '')) || 'export';
    const data = buildExportRows();

    if (format === 'json') {
      const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `${baseName}.json`;
      a.click();
      URL.revokeObjectURL(url);
      return;
    }

    // For csv/xlsx use SheetJS
    const sheet = XLSX.utils.json_to_sheet(data, { header: headers });
    if (format === 'csv') {
      const csv = XLSX.utils.sheet_to_csv(sheet);
      const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `${baseName}.csv`;
      a.click();
      URL.revokeObjectURL(url);
      return;
    }

    // default: xlsx
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, sheet, 'Sheet1');
    XLSX.writeFile(wb, `${baseName}.xlsx`);
  };

  return (
    <section>
      <div
        className={`dropzone ${isDragging ? 'is-active' : ''} ${isLoading ? 'is-loading' : ''}`}
        onDragOver={(event) => {
          event.preventDefault();
          setIsDragging(true);
        }}
        onDragLeave={() => setIsDragging(false)}
        onDrop={onDrop}
      >
        <div className="dropzone-inner">
          <strong>Drag and drop your spreadsheet here</strong>
          <p>or click below to browse your computer.</p>
          <input
            ref={inputRef}
            type="file"
            accept=".xlsx,.xls"
            hidden
            onChange={(event) => handleFile(event.target.files?.[0])}
          />
          <button className="upload-btn" type="button" onClick={onBrowseClick}>
            {isLoading ? 'Parsing...' : 'Browse spreadsheet'}
          </button>
          <p>Accepted formats: .xlsx and .xls</p>
        </div>
      </div>

      <div className="status-row">
        <div className={`status-pill ${error ? 'error' : ''}`}>
          {isLoading ? '⏳ Parsing workbook…' : error ? `⚠️ ${error}` : '✅ Ready'}
        </div>
        <div className="meta">{status}</div>
      </div>

          <div style={{ display: rows.length ? 'flex' : 'none', gap: 12, alignItems: 'center', marginTop: 12 }}>
            <label className="meta" style={{ marginRight: 8 }}>Export format</label>
            <select value={exportFormat} onChange={(e) => setExportFormat(e.target.value as any)}>
              <option value="xlsx">Excel (.xlsx)</option>
              <option value="csv">CSV (.csv)</option>
              <option value="json">JSON (.json)</option>
            </select>
            <button
              className="upload-btn"
              type="button"
              onClick={() => downloadParsed(exportFormat)}
              style={{ marginLeft: 12 }}
            >
              Download
            </button>
          </div>

      {rows.length > 0 ? (
        <>
          <div className="input-row">
            <input
              className="search-input"
              value={search}
              onChange={(event) => setSearch(event.target.value)}
              placeholder="Search any column value..."
            />
            <button className="secondary-btn" type="button" onClick={resetState}>
              Clear / load another file
            </button>
          </div>
          <div className="meta">Showing {filteredRows.length} of {rows.length} items</div>
          <InventoryCardGrid headers={headers} rows={cards} />
        </>
      ) : null}
    </section>
  );
}
