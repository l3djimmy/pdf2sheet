import { useState, useCallback } from 'react';
import { Upload, FileSpreadsheet, Download, Trash2, Loader2, CheckCircle } from 'lucide-react';
import * as XLSX from 'xlsx';
import './App.css';

interface ExtractedTable {
  id: string;
  headers: string[];
  rows: string[][];
}

function App() {
  const [file, setFile] = useState<File | null>(null);
  const [tables, setTables] = useState<ExtractedTable[]>([]);
  const [loading, setLoading] = useState(false);
  const [done, setDone] = useState(false);
  const [error, setError] = useState('');

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    const f = e.dataTransfer.files[0];
    if (f?.type === 'application/pdf') {
      setFile(f);
      setError('');
      setDone(false);
      setTables([]);
    } else {
      setError('Only PDF files are supported.');
    }
  }, []);

  const handleFile = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const f = e.target.files?.[0];
    if (f?.type === 'application/pdf') {
      setFile(f);
      setError('');
      setDone(false);
      setTables([]);
    } else {
      setError('Only PDF files are supported.');
    }
  }, []);

  const runExtraction = async () => {
    if (!file) return;
    setLoading(true);
    setError('');
    try {
      const pdfjs = await import('pdfjs-dist');
      pdfjs.GlobalWorkerOptions.workerSrc = `https://cdnjs.cloudflare.com/ajax/libs/pdf.js/${pdfjs.version}/pdf.worker.min.mjs`;
      const buf = await file.arrayBuffer();
      const pdf = await pdfjs.getDocument({ data: buf }).promise;
      const allRows: string[][] = [];
      for (let i = 1; i <= pdf.numPages; i++) {
        const page = await pdf.getPage(i);
        const text = await page.getTextContent();
        const lines = new Map<number, string[]>();
        for (const item of text.items) {
          const it = item as any;
          const y = Math.round(it.transform[5] / 2) * 2;
          if (!lines.has(y)) lines.set(y, []);
          lines.get(y)!.push(it.str);
        }
        const sorted = [...lines.entries()].sort((a, b) => b[0] - a[0]);
        for (const [, words] of sorted) {
          if (words.length > 1) allRows.push(words);
        }
      }
      if (!allRows.length) throw new Error('No tabular data found in PDF.');
      const headers = allRows[0];
      const rows = allRows.slice(1);
      setTables([{ id: 't1', headers, rows }]);
      setDone(true);
    } catch (err: any) {
      setError(err.message || 'Extraction failed.');
    } finally {
      setLoading(false);
    }
  };

  const downloadXLSX = (tid: string) => {
    const t = tables.find(x => x.id === tid);
    if (!t) return;
    const ws = XLSX.utils.aoa_to_sheet([t.headers, ...t.rows]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    XLSX.writeFile(wb, 'extracted.xlsx');
  };

  const downloadCSV = (tid: string) => {
    const t = tables.find(x => x.id === tid);
    if (!t) return;
    const csv = [t.headers, ...t.rows].map(r => r.map(c => `"${c.replace(/"/g, '""')}"`).join(',')).join('\n');
    const blob = new Blob([csv], { type: 'text/csv' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'extracted.csv';
    a.click();
    URL.revokeObjectURL(a.href);
  };

  return (
    <div className="min-h-screen bg-slate-50 text-slate-800">
      <header className="border-b bg-white px-6 py-4 flex items-center gap-2 shadow-sm">
        <FileSpreadsheet className="w-6 h-6 text-emerald-600" />
        <h1 className="text-xl font-bold">PDF2Sheet</h1>
      </header>

      <main className="max-w-3xl mx-auto px-4 py-10">
        {
        !(file || tables.length || error) && <div className="mb-6 rounded-lg bg-white shadow p-4">
          <p className="text-sm text-slate-600">Free, in-browser app. Upload a PDF invoice, statement, or report. Get editable spreadsheets. No signup needed.</p>
        </div>
        }

        <div
          className={"border-2 border-dashed rounded-xl p-10 text-center transition-colors " + (file ? 'bg-white border-emerald-400' : 'bg-white border-slate-300 hover:border-emerald-400')}
          onDragOver={e => e.preventDefault()}
          onDrop={handleDrop}
        >
          <input id="pdf" type="file" accept="application/pdf" className="hidden" onChange={handleFile} />
          <label htmlFor="pdf" className="cursor-pointer flex flex-col items-center gap-3">
            <Upload className="w-10 h-10 text-slate-400" />
            <span className="text-sm text-slate-500">{file ? file.name : 'Click or drop a PDF here'}</span>
          </label>
        </div>

        {file && !done && (
          <div className="mt-4 flex items-center gap-3">
            <button onClick={runExtraction} disabled={loading} className="bg-emerald-600 hover:bg-emerald-700 disabled:opacity-60 text-white px-5 py-2.5 rounded-md font-medium flex items-center gap-2">
              {loading ? <Loader2 className="w-4 h-4 animate-spin" /> : <CheckCircle className="w-4 h-4" />}
              {loading ? 'Extracting...' : 'Extract to Spreadsheet'}
            </button>
            <button onClick={() => { setFile(null); setTables([]); setDone(false); setError(''); }} className="text-red-600 hover:text-red-700 px-3 py-2 flex items-center gap-1">
              <Trash2 className="w-4 h-4" /> Remove
            </button>
          </div>
        )}

        {error && <p className="mt-3 text-red-600 text-sm">{error}</p>}

        {tables.map(t => (
          <div key={t.id} className="mt-6 rounded-lg bg-white shadow overflow-hidden">
            <div className="flex items-center gap-2 px-4 py-3 border-b bg-slate-50">
              <FileSpreadsheet className="w-5 h-5 text-emerald-600" />
              <span className="font-semibold text-sm">Extracted Table</span>
              <div className="ml-auto flex gap-2">
                <button onClick={() => downloadXLSX(t.id)} className="text-xs bg-emerald-600 hover:bg-emerald-700 text-white px-3 py-1.5 rounded flex items-center gap-1"><Download className="w-3 h-3" /> .xlsx</button>
                <button onClick={() => downloadCSV(t.id)} className="text-xs bg-slate-700 hover:bg-slate-800 text-white px-3 py-1.5 rounded flex items-center gap-1"><Download className="w-3 h-3" /> .csv</button>
              </div>
            </div>
            <div className="overflow-auto">
              <table className="w-full text-sm">
                <thead><tr className="bg-slate-50 border-b">{t.headers.map((h, i) => <th key={i} className="text-left px-4 py-2 font-medium text-slate-600 whitespace-nowrap">{h}</th>)}</tr></thead>
                <tbody>{t.rows.map((r, i) => <tr key={i} className="border-b hover:bg-slate-50">{r.map((c, j) => <td key={j} className="px-4 py-2 text-slate-700">{c}</td>)}</tr>)}</tbody>
              </table>
            </div>
          </div>
        ))}
      </main>
    </div>
  );
}

export default App;
