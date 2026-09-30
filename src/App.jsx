import React, { useState, useEffect } from 'react';
import { UploadCloud, FileSpreadsheet, CheckCircle2, AlertCircle, Download, Loader2, Info, SlidersHorizontal, Calculator, Plus, Trash2, RotateCcw, FileDown, FileUp } from 'lucide-react';
import { processExcelFiles, computeRateSuggestions, DEFAULT_UPLIFT_ROWS, buildUpliftMasterBlob, parseUpliftMasterFile } from './utils/excelProcessor';
import './index.css';

const MASTER_KEY = 'nlc_uplift_master_v1';

const loadMaster = () => {
  try {
    const raw = localStorage.getItem(MASTER_KEY);
    if (raw) {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed)) return parsed;
    }
  } catch { /* ignore */ }
  return DEFAULT_UPLIFT_ROWS.map(r => ({ ...r }));
};

function App() {
  const [files, setFiles] = useState({ gt: null, nlc: null, item: null });
  const [isProcessing, setIsProcessing] = useState(false);
  const [error, setError] = useState(null);
  const [downloadUrl, setDownloadUrl] = useState(null);

  // Master 1 — editable uplift % table
  const [upliftRows, setUpliftRows] = useState(loadMaster);
  const [showMaster1, setShowMaster1] = useState(false);
  const [importMsg, setImportMsg] = useState(null);
  const [showProcedure, setShowProcedure] = useState(true);

  // Master 2 — rate -> % calculator
  const [showMaster2, setShowMaster2] = useState(false);
  const [rateText, setRateText] = useState('');
  const [rateResults, setRateResults] = useState(null);
  const [rateBusy, setRateBusy] = useState(false);
  const [rateError, setRateError] = useState(null);

  useEffect(() => {
    try { localStorage.setItem(MASTER_KEY, JSON.stringify(upliftRows)); } catch { /* ignore */ }
  }, [upliftRows]);

  const percentByCode = () => {
    const m = {};
    upliftRows.forEach(r => {
      if (r.code && r.pct !== '' && !isNaN(Number(r.pct))) m[String(r.code).trim().toLowerCase()] = Number(r.pct);
    });
    return m;
  };

  const buildOverrides = () => {
    const m = {};
    upliftRows.forEach(r => {
      if (r.code && r.pct !== '' && !isNaN(Number(r.pct))) m[String(r.code).trim().toLowerCase()] = 1 + Number(r.pct) / 100;
    });
    return m;
  };

  const handleFileChange = (e, type) => {
    const file = e.target.files[0];
    if (file) {
      setFiles(prev => ({ ...prev, [type]: file }));
      setError(null);
      setDownloadUrl(null);
    }
  };

  const handleProcess = async () => {
    if (!files.gt || !files.nlc || !files.item) {
      setError("Please upload all three required files.");
      return;
    }
    setIsProcessing(true);
    setError(null);
    setDownloadUrl(null);
    try {
      const blob = await processExcelFiles(files.gt, files.nlc, files.item, { upliftOverrides: buildOverrides() });
      setDownloadUrl(URL.createObjectURL(blob));
    } catch (err) {
      setError(err.message || "An error occurred while processing the files.");
    } finally {
      setIsProcessing(false);
    }
  };

  // ----- Master 1 handlers -----
  const updateRow = (i, field, val) => setUpliftRows(rows => rows.map((r, idx) => idx === i ? { ...r, [field]: val } : r));
  const addRow = () => setUpliftRows(rows => [...rows, { code: '', name: '', pct: 5 }]);
  const deleteRow = (i) => setUpliftRows(rows => rows.filter((_, idx) => idx !== i));
  const resetRows = () => { setUpliftRows(DEFAULT_UPLIFT_ROWS.map(r => ({ ...r }))); setImportMsg(null); };

  const exportMaster = () => {
    const blob = buildUpliftMasterBlob(upliftRows);
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'NLC_Uplift_Master.xlsx';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  };

  const importMaster = async (e) => {
    const file = e.target.files[0];
    e.target.value = '';
    if (!file) return;
    setImportMsg(null);
    try {
      const rows = await parseUpliftMasterFile(file);
      if (!rows.length) { setImportMsg('No valid rows found — the file needs Item Code and Uplift % columns.'); return; }
      setUpliftRows(rows.map(r => ({ code: r.code, name: r.name || '', pct: r.pct })));
      setImportMsg(`Imported ${rows.length} item${rows.length === 1 ? '' : 's'} (replaced the master).`);
    } catch (err) {
      setImportMsg('Import failed: ' + (err.message || String(err)));
    }
  };

  // ----- Master 2 handlers -----
  const parseRateRows = (text) => {
    return text.split(/\r?\n/).map(l => l.replace(/\s+$/, '')).filter(l => l.trim())
      .map(line => (line.indexOf('\t') > -1 ? line.split('\t') : line.split(/\s{2,}/)).map(p => p.trim()))
      .filter(p => p.length >= 2 && !/^item\s*code$/i.test(p[0]))
      .map(p => ({ code: p[0], name: p[1], mrp: p.length >= 4 ? p[2] : '', rate: p[p.length - 1] }));
  };

  const handleCalcRates = async () => {
    if (!files.gt) { setRateError('Upload the GT Sheet first — it is used to look up NLC per-pkt (sale basis) rates.'); return; }
    const rows = parseRateRows(rateText);
    if (!rows.length) { setRateError('Paste at least one row: Item Code, Item Name, MRP, New Rate (tab-separated from Excel).'); return; }
    setRateBusy(true); setRateError(null);
    try {
      const res = await computeRateSuggestions(files.gt, rows, percentByCode());
      setRateResults(res);
    } catch (e) {
      setRateError(e.message || 'Failed to compute suggestions.');
    } finally {
      setRateBusy(false);
    }
  };

  const applySuggestions = () => {
    if (!rateResults) return;
    setUpliftRows(rows => {
      const next = rows.map(r => ({ ...r }));
      const idxByCode = new Map(next.map((r, i) => [String(r.code).trim().toLowerCase(), i]));
      rateResults.forEach(res => {
        if (res.pct == null || !res.code) return;
        const key = String(res.code).trim().toLowerCase();
        if (idxByCode.has(key)) next[idxByCode.get(key)].pct = res.pct;
        else { next.push({ code: res.code, name: res.name || '', pct: res.pct }); idxByCode.set(key, next.length - 1); }
      });
      return next;
    });
  };

  const FileUploadBox = ({ title, type, file, id }) => (
    <div className="upload-box">
      <input type="file" id={id} accept=".xlsx, .xls, .csv" onChange={(e) => handleFileChange(e, type)} className="hidden-input" />
      <label htmlFor={id} className={`upload-label ${file ? 'has-file' : ''}`}>
        <div className="icon-container">
          {file ? <CheckCircle2 className="icon success" /> : <FileSpreadsheet className="icon" />}
        </div>
        <div className="upload-content">
          <h3>{title}</h3>
          <p>{file ? file.name : 'Click to browse or drag file'}</p>
        </div>
      </label>
    </div>
  );

  // inline style helpers (kept local so index.css is untouched)
  const S = {
    panel: { marginTop: 20, padding: 20, borderRadius: 12, background: 'rgba(255,255,255,0.04)', border: '1px solid rgba(255,255,255,0.10)' },
    head: { display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 12, cursor: 'pointer' },
    headLeft: { display: 'flex', alignItems: 'center', gap: 10 },
    btn: { display: 'inline-flex', alignItems: 'center', gap: 6, padding: '6px 12px', borderRadius: 8, border: '1px solid rgba(255,255,255,0.18)', background: 'rgba(255,255,255,0.06)', color: 'inherit', cursor: 'pointer', fontSize: 13 },
    input: { width: '100%', boxSizing: 'border-box', padding: '6px 8px', borderRadius: 6, border: '1px solid rgba(255,255,255,0.18)', background: 'rgba(0,0,0,0.25)', color: 'inherit', fontSize: 13 },
    th: { textAlign: 'left', padding: '6px 8px', fontSize: 12, opacity: 0.7, borderBottom: '1px solid rgba(255,255,255,0.12)', whiteSpace: 'nowrap' },
    td: { padding: '4px 8px', borderBottom: '1px solid rgba(255,255,255,0.06)', fontSize: 13, verticalAlign: 'middle' },
    tableWrap: { maxHeight: 340, overflow: 'auto', marginTop: 12, border: '1px solid rgba(255,255,255,0.08)', borderRadius: 8 },
    textarea: { width: '100%', boxSizing: 'border-box', minHeight: 120, padding: 10, borderRadius: 8, border: '1px solid rgba(255,255,255,0.18)', background: 'rgba(0,0,0,0.25)', color: 'inherit', fontFamily: 'monospace', fontSize: 12.5 },
  };

  return (
    <div className="app-container">
      <div className="background-gradient"></div>
      <div className="content-wrapper">
        <header className="header">
          <h1>Data Nexus Engine</h1>
          <p>Advanced Excel Processing for GT, NLC, and Item Operations</p>
        </header>

        <main className="main-panel">
          <div className="upload-grid">
            <FileUploadBox title="GT Sheet" type="gt" file={files.gt} id="gt-upload" />
            <FileUploadBox title="NLC Sheet" type="nlc" file={files.nlc} id="nlc-upload" />
            <FileUploadBox title="Item List" type="item" file={files.item} id="item-upload" />
          </div>

          <div className="conditions-panel">
            <div className="conditions-header" style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', cursor: 'pointer' }} onClick={() => setShowProcedure(v => !v)}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                <Info size={20} className="info-icon" />
                <h2 style={{ margin: 0 }}>Procedure — Step by Step</h2>
              </div>
              <span style={{ fontSize: 13, opacity: 0.7 }}>{showProcedure ? 'Hide ▲' : 'Show ▼'}</span>
            </div>

            {showProcedure && (
              <div style={{ marginTop: 6 }}>
                <h3 style={{ fontSize: 15, margin: '14px 0 6px' }}>A. How to use this tool</h3>
                <ol className="conditions-list" style={{ lineHeight: 1.6 }}>
                  <li><strong>Upload three files</strong> above: <strong>GT Sheet</strong> (the Retail Costing workbook), <strong>NLC Sheet</strong>, and <strong>Item List</strong>.</li>
                  <li><em>(Optional)</em> Open <strong>Uplift % Master</strong> to review or change the per-item uplift %. You can <strong>Bulk import</strong> an .xlsx/.csv or <strong>Bulk export</strong> the current master.</li>
                  <li><em>(Optional)</em> Use the <strong>Rate → % Calculator</strong> to turn new selling rates into the % to put in the master, then <strong>Apply to Uplift % Master</strong>.</li>
                  <li>Click <strong>Process Excel Files</strong>, then <strong>Download Processed File</strong> (sheet name <code>NLC_Processed</code>).</li>
                </ol>

                <h3 style={{ fontSize: 15, margin: '16px 0 6px' }}>B. What the tool does to your data</h3>
                <ol className="conditions-list" style={{ lineHeight: 1.6 }}>
                  <li><strong>Reads the files.</strong> From the multi-sheet GT workbook it auto-selects the sheet named <strong>“GT”</strong> (the one holding SKU + the pricing columns), not the first sheet.</li>
                  <li><strong>Matches items</strong> across files using a normalized key — NLC <em>Item Name</em> ↔ GT <em>SKU</em> ↔ Item List <em>Item Name</em> — that ignores extra spaces and non‑breaking‑space “mojibake”, so the same product isn’t treated as two.</li>
                  <li><strong>Appends missing SKUs:</strong> any GT SKU not already in the NLC sheet is added as a new row (Item Group &amp; UOM from GT, Item Code from the Item List).</li>
                  <li><strong>Shifts dates by one month:</strong> <em>From Date</em> → +1 month; <em>To Date</em> → end of that new month; <em>Month</em> → the new month’s name.</li>
                  <li><strong>Forces defaults on every row:</strong> Customer Group = <strong>“GT”</strong>, GST % = <strong>5%</strong> (0.05).</li>
                  <li><strong>Merges from the Item List</strong> (by Item Name): Item Code &amp; Item Group (filled only if blank), <strong>MRP</strong> and <strong>EAN/Barcode</strong>.</li>
                  <li><strong>Sets UOM</strong> = GT <em>“UOM (G)”</em> ÷ 1000 (grams → kg).</li>
                  <li><strong>Pulls GT costs</strong> for the matched SKU: Ex‑Factory, Logistics, Total Cost (per kg), NLC per kg sale &amp; cost basis, and Margin %.</li>
                  <li><strong>Applies the uplift factor</strong> (from the Uplift % Master by Item Code; any item not listed uses <strong>5%</strong>) to: <em>Ex‑Factory</em>, <em>Non‑GST Final</em> (= Total Cost), <em>Grand Final</em> (= NLC sale), and the <em>Cost‑Basis</em> column (= NLC cost). <strong>Logistics Cost is never uplifted</strong> — it is carried over as‑is.</li>
                  <li><strong>Recomputes:</strong> Margin = Margin % × Cost‑Basis; GST Amount = 5% × Non‑GST Final.</li>
                  <li><strong>Writes the output</strong> to the <code>NLC_Processed</code> sheet with <strong>live Excel formulas</strong> (each cell also stores a cached value, so it displays correctly and recalculates when opened).</li>
                </ol>

                <h3 style={{ fontSize: 15, margin: '16px 0 6px' }}>C. The two masters</h3>
                <ul className="conditions-list" style={{ lineHeight: 1.6 }}>
                  <li><strong>Uplift % Master</strong> — per‑Item‑Code uplift %. Edit inline, add/remove rows, reset to defaults, and bulk import/export. Saved in this browser and used automatically on the next Process run.</li>
                  <li><strong>Rate → % Calculator</strong> — paste <em>Item Code, Item Name, MRP, New Rate</em>. It looks up the GT <em>“NLC per pkt (sale basis)”</em> for each SKU and computes <strong>Required % = (New Rate ÷ that base) − 1</strong>, flags items whose % differs from the master, and can apply the results into the Uplift % Master.</li>
                </ul>
              </div>
            )}
          </div>

          {/* ---------- Master 1: Uplift % Master ---------- */}
          <div style={S.panel}>
            <div style={S.head} onClick={() => setShowMaster1(v => !v)}>
              <div style={S.headLeft}>
                <SlidersHorizontal size={18} />
                <h2 style={{ margin: 0, fontSize: 17 }}>Uplift % Master</h2>
                <span style={{ fontSize: 12, opacity: 0.6 }}>({upliftRows.length} items)</span>
              </div>
              <span style={{ fontSize: 13, opacity: 0.7 }}>{showMaster1 ? 'Hide ▲' : 'Adjust ▼'}</span>
            </div>
            {showMaster1 && (
              <div>
                <p style={{ fontSize: 12.5, opacity: 0.7, marginTop: 8 }}>
                  Per-Item-Code uplift %. Applied to GT cost columns (Logistics excluded); any item not listed uses 5%. Saved in this browser.
                </p>
                <div style={{ display: 'flex', gap: 8, marginTop: 8, flexWrap: 'wrap' }}>
                  <button style={S.btn} onClick={addRow}><Plus size={14} /> Add item</button>
                  <button style={S.btn} onClick={resetRows}><RotateCcw size={14} /> Reset to defaults</button>
                  <button style={S.btn} onClick={exportMaster}><FileDown size={14} /> Bulk export (.xlsx)</button>
                  <label style={{ ...S.btn, marginBottom: 0 }}>
                    <FileUp size={14} /> Bulk import
                    <input type="file" accept=".xlsx,.xls,.csv" onChange={importMaster} style={{ display: 'none' }} />
                  </label>
                </div>
                {importMsg && <p style={{ fontSize: 12.5, marginTop: 8, opacity: 0.85 }}>{importMsg}</p>}
                <div style={S.tableWrap}>
                  <table style={{ width: '100%', borderCollapse: 'collapse' }}>
                    <thead>
                      <tr>
                        <th style={S.th}>Item Code</th>
                        <th style={S.th}>Item Name (optional)</th>
                        <th style={{ ...S.th, width: 90 }}>Uplift %</th>
                        <th style={{ ...S.th, width: 44 }}></th>
                      </tr>
                    </thead>
                    <tbody>
                      {upliftRows.map((r, i) => (
                        <tr key={i}>
                          <td style={S.td}><input style={S.input} value={r.code || ''} onChange={e => updateRow(i, 'code', e.target.value)} placeholder="Seeds_11-..." /></td>
                          <td style={S.td}><input style={S.input} value={r.name || ''} onChange={e => updateRow(i, 'name', e.target.value)} /></td>
                          <td style={S.td}><input style={{ ...S.input, textAlign: 'right' }} type="number" step="0.01" value={r.pct} onChange={e => updateRow(i, 'pct', e.target.value)} /></td>
                          <td style={S.td}><button title="Remove" style={{ ...S.btn, padding: 6 }} onClick={() => deleteRow(i)}><Trash2 size={14} /></button></td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}
          </div>

          {/* ---------- Master 2: Rate -> % Calculator ---------- */}
          <div style={S.panel}>
            <div style={S.head} onClick={() => setShowMaster2(v => !v)}>
              <div style={S.headLeft}>
                <Calculator size={18} />
                <h2 style={{ margin: 0, fontSize: 17 }}>Rate → % Calculator</h2>
              </div>
              <span style={{ fontSize: 13, opacity: 0.7 }}>{showMaster2 ? 'Hide ▲' : 'Open ▼'}</span>
            </div>
            {showMaster2 && (
              <div>
                <p style={{ fontSize: 12.5, opacity: 0.7, marginTop: 8 }}>
                  Paste rows as <strong>Item Code, Item Name, MRP, New Rate</strong> (tab-separated, e.g. copied from Excel). Uses the uploaded GT Sheet's
                  <strong> NLC per pkt (sale basis)</strong>. Required % = (New Rate ÷ base) − 1. Rows whose % differs from the master are flagged.
                </p>
                <textarea
                  style={S.textarea}
                  value={rateText}
                  onChange={e => setRateText(e.target.value)}
                  placeholder={"Dates_4-3011\tBarkat dates Farmley Standee Pouch 250 g\t149\t82.52\nDates_4-2472\tBarkat Dates Farmley Standee Pouch 500g\t299\t160.07"}
                />
                <div style={{ display: 'flex', gap: 8, marginTop: 8, flexWrap: 'wrap' }}>
                  <button style={{ ...S.btn, opacity: rateBusy ? 0.6 : 1 }} onClick={handleCalcRates} disabled={rateBusy}>
                    {rateBusy ? <Loader2 size={14} className="spinner" /> : <Calculator size={14} />} Calculate %
                  </button>
                  {rateResults && rateResults.some(r => r.pct != null) && (
                    <button style={S.btn} onClick={applySuggestions}><CheckCircle2 size={14} /> Apply to Uplift % Master</button>
                  )}
                </div>
                {rateError && (
                  <div className="error-message" style={{ marginTop: 10 }}><AlertCircle size={18} /><span>{rateError}</span></div>
                )}
                {rateResults && (
                  <div style={S.tableWrap}>
                    <table style={{ width: '100%', borderCollapse: 'collapse' }}>
                      <thead>
                        <tr>
                          <th style={S.th}>Item Code</th>
                          <th style={S.th}>Item Name</th>
                          <th style={{ ...S.th, textAlign: 'right' }}>New Rate</th>
                          <th style={{ ...S.th, textAlign: 'right' }}>Base (NLC/pkt sale)</th>
                          <th style={{ ...S.th, textAlign: 'right' }}>Required %</th>
                          <th style={{ ...S.th, textAlign: 'right' }}>Current %</th>
                          <th style={S.th}>Status</th>
                        </tr>
                      </thead>
                      <tbody>
                        {rateResults.map((r, i) => (
                          <tr key={i}>
                            <td style={S.td}>{r.code}</td>
                            <td style={{ ...S.td, maxWidth: 260, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{r.name}</td>
                            <td style={{ ...S.td, textAlign: 'right' }}>{r.rate}</td>
                            <td style={{ ...S.td, textAlign: 'right' }}>{r.base == null ? '—' : r.base.toFixed(2)}</td>
                            <td style={{ ...S.td, textAlign: 'right', fontWeight: 600 }}>{r.pct == null ? '—' : r.pct + '%'}</td>
                            <td style={{ ...S.td, textAlign: 'right', opacity: 0.8 }}>{r.current == null ? '—' : r.current + '%'}</td>
                            <td style={S.td}>
                              {!r.matched ? <span style={{ color: '#f0a020' }}>No GT match</span>
                                : r.needsChange ? <span style={{ color: '#ff6b6b' }}>Needs change</span>
                                  : <span style={{ color: '#3ddc84' }}>OK</span>}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>
            )}
          </div>

          {error && (
            <div className="error-message">
              <AlertCircle size={20} />
              <span>{error}</span>
            </div>
          )}

          <div className="action-area">
            {!downloadUrl ? (
              <button
                className={`primary-btn ${isProcessing ? 'processing' : ''}`}
                onClick={handleProcess}
                disabled={isProcessing || !files.gt || !files.nlc || !files.item}
              >
                {isProcessing ? (<><Loader2 className="spinner" /> Processing Data...</>) : (<><UploadCloud size={20} /> Process Excel Files</>)}
              </button>
            ) : (
              <a href={downloadUrl} download={`NLC_Processed_${new Date().getTime()}.xlsx`} className="download-btn">
                <Download size={20} /> Download Processed File
              </a>
            )}
          </div>
        </main>

        <footer className="footer">
          <p>Designed for automated data transformations.</p>
        </footer>
      </div>
    </div>
  );
}

export default App;
