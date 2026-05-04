import { useState, useRef } from 'react';
import type { ChangeEvent, FormEvent } from 'react';
import axios from 'axios';
import './index.css';

const API_BASE = 'http://127.0.0.1:5001/api';
const SINGLE_URL = `${API_BASE}/process`;
const BULK_URL = `${API_BASE}/bulk`;
const MULTI_URL = `${API_BASE}/bulk-multi`;
const PS_URL = `${API_BASE}/bulk-ps`;

type Mode = 'single' | 'bulk' | 'multi' | 'ps';

function App() {
  const [excelFile, setExcelFile] = useState<File | null>(null);
  const [wordFiles, setWordFiles] = useState<FileList | null>(null);
  const [scsFile, setScsFile] = useState<File | null>(null);
  const [defaultFile, setDefaultFile] = useState<File | null>(null);
  const [mode, setMode] = useState<Mode>('single');
  const [isProcessing, setIsProcessing] = useState(false);
  const [status, setStatus] = useState('');
  const [progress, setProgress] = useState(0);

  const [psTemplateFiles, setPsTemplateFiles] = useState<FileList | null>(null);

  const excelInputRef = useRef<HTMLInputElement>(null);
  const wordInputRef = useRef<HTMLInputElement>(null);
  const scsInputRef = useRef<HTMLInputElement>(null);
  const defaultInputRef = useRef<HTMLInputElement>(null);
  const psTemplateInputRef = useRef<HTMLInputElement>(null);

  const handleExcelChange = (e: ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      setExcelFile(e.target.files[0]);
    }
  };

  const handleWordChange = (e: ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      setWordFiles(e.target.files);
    }
  };

  const handleScsChange = (e: ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      setScsFile(e.target.files[0]);
    }
  };

  const handleDefaultChange = (e: ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      setDefaultFile(e.target.files[0]);
    }
  };

  const handlePsTemplateChange = (e: ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      setPsTemplateFiles(e.target.files);
    }
  };

  const handleSubmit = async (e: FormEvent) => {
    e.preventDefault();

    if (mode === 'multi') {
      if (!excelFile || !scsFile || !defaultFile) {
        alert("Please select Excel, SCS template, and Default template.");
        return;
      }
    } else if (mode === 'ps') {
      if (!excelFile || !psTemplateFiles || psTemplateFiles.length === 0) {
        alert("Please select an Excel file and at least one Word template.");
        return;
      }
    } else {
      if (!excelFile || !wordFiles) {
        alert("Please select both an Excel file and Word templates.");
        return;
      }
      if (mode === 'bulk' && wordFiles.length !== 1) {
        alert("Bulk mode expects exactly one Word template.");
        return;
      }
    }

    setIsProcessing(true);
    setStatus('Uploading and processing...');
    setProgress(30);

    const formData = new FormData();
    formData.append('excel', excelFile!);

    if (mode === 'multi') {
      formData.append('word_scs', scsFile!);
      formData.append('word_default', defaultFile!);
    } else if (mode === 'ps') {
      Array.from(psTemplateFiles!).forEach((file) => {
        formData.append('word', file);
      });
    } else {
      Array.from(wordFiles!).forEach((file) => {
        formData.append('word', file);
      });
    }

    const endpoint =
      mode === 'ps' ? PS_URL :
      mode === 'multi' ? MULTI_URL :
      mode === 'bulk' ? BULK_URL : SINGLE_URL;

    try {
      const response = await axios.post(endpoint, formData, {
        responseType: 'blob',
        onUploadProgress: (progressEvent) => {
           if (progressEvent.total) {
               const percentCompleted = Math.round((progressEvent.loaded * 100) / progressEvent.total);
               setProgress(30 + (percentCompleted * 0.7)); // Scale 30-100%
           }
        }
      });

      setProgress(100);
      const scsCount = response.headers['x-scs-count'];
      const defCount = response.headers['x-default-count'];
      const recordCount = response.headers['x-record-count'];
      if (mode === 'multi' && scsCount !== undefined) {
        setStatus(`Success! ${scsCount} SCS + ${defCount} Default = ${Number(scsCount) + Number(defCount)} docs.`);
      } else if (mode === 'ps' && recordCount !== undefined) {
        setStatus(`Success! ${recordCount} position statements filled and zipped.`);
      } else {
        setStatus('Success! Download started.');
      }

      // Handle download
      const contentDisposition = response.headers['content-disposition'];
      let filename: string;
      if (mode === 'ps') {
        filename = 'position_statements_filled.zip';
      } else if (mode === 'multi') {
        filename = 'multi_template_bulk_filled.zip';
      } else if (mode === 'bulk') {
        filename = `${wordFiles![0].name.split('.')[0]}_bulk_filled.zip`;
      } else {
        filename = wordFiles!.length === 1
          ? `${wordFiles![0].name.split('.')[0]}_filled.docx`
          : 'filled_documents.zip';
      }

      if (contentDisposition) {
        const filenameMatch = contentDisposition.match(/filename="?(.+)"?/);
        if (filenameMatch && filenameMatch[1]) {
          filename = filenameMatch[1];
        }
      }

      const url = window.URL.createObjectURL(new Blob([response.data]));
      const link = document.createElement('a');
      link.href = url;
      link.setAttribute('download', filename);
      document.body.appendChild(link);
      link.click();
      link.remove();
      window.URL.revokeObjectURL(url);

      setTimeout(() => {
        setIsProcessing(false);
        setProgress(0);
        setStatus('');
      }, 4000);

    } catch (error: any) {
      console.error('Error processing files:', error);
      let msg = error.message;
      // Blob error responses need to be parsed
      if (error.response?.data instanceof Blob) {
        try {
          const text = await error.response.data.text();
          const json = JSON.parse(text);
          msg = json.error || msg;
        } catch {
          // ignore
        }
      } else if (error.response?.data?.error) {
        msg = error.response.data.error;
      }
      setStatus(`Error: ${msg}`);
      setIsProcessing(false);
    }
  };

  const modeHint =
    mode === 'single'
      ? 'Uses the "Fields to Enter" sheet (one placeholder per row).'
      : mode === 'bulk'
        ? 'Uses the "Export" sheet — placeholders in row 1, one filled doc per data row.'
        : mode === 'ps'
          ? 'Uses the "Field to Fill" sheet. Routes each record to the correct template based on "Number of Comps in Position Statement" and "Procedure Type" rows.'
          : 'Uses the "Fields to Replace" sheet (column-oriented). Each record auto-routes to SCS or Default template based on the [Procedure] row.';

  const submitLabel = isProcessing
    ? 'Processing...'
    : mode === 'ps'
      ? 'Run Position Statement Sweep'
      : mode === 'multi'
        ? 'Run Multi-Template Sweep'
        : mode === 'bulk'
          ? 'Run Bulk Sweep'
          : 'Generate Documents';

  return (
    <>
      <div className="background-blobs">
        <div className="blob blob-1"></div>
        <div className="blob blob-2"></div>
        <div className="blob blob-3"></div>
      </div>

      <main className="container">
        <header className="glass-header">
          <h1>Auto-Filler</h1>
          <p>Seamlessly populate Word templates with Excel data</p>
        </header>

        <section className="card glass">
          <form onSubmit={handleSubmit}>
            <div className="mode-toggle" role="tablist" aria-label="Fill mode">
              <button
                type="button"
                role="tab"
                aria-selected={mode === 'single'}
                className={`mode-button ${mode === 'single' ? 'active' : ''}`}
                onClick={() => setMode('single')}
              >
                Single Fill
              </button>
              <button
                type="button"
                role="tab"
                aria-selected={mode === 'bulk'}
                className={`mode-button ${mode === 'bulk' ? 'active' : ''}`}
                onClick={() => setMode('bulk')}
              >
                Bulk Sweep
              </button>
              <button
                type="button"
                role="tab"
                aria-selected={mode === 'multi'}
                className={`mode-button ${mode === 'multi' ? 'active' : ''}`}
                onClick={() => setMode('multi')}
              >
                Multi-Template
              </button>
              <button
                type="button"
                role="tab"
                aria-selected={mode === 'ps'}
                className={`mode-button ${mode === 'ps' ? 'active' : ''}`}
                onClick={() => setMode('ps')}
              >
                Position Statement
              </button>
            </div>
            <p className="mode-hint">{modeHint}</p>
            <div className="upload-section">
              <div className="input-group">
                <label className="section-label">1. Data Source (Excel)</label>
                <div
                  className="file-drop-zone"
                  onClick={() => excelInputRef.current?.click()}
                  onDragOver={(e) => e.preventDefault()}
                >
                  <input
                    type="file"
                    ref={excelInputRef}
                    onChange={handleExcelChange}
                    accept=".xlsx"
                    hidden
                  />
                  <span className="icon">📊</span>
                  <p className="file-name">
                    {excelFile ? excelFile.name : "Drop your Excel file here or click to browse"}
                  </p>
                </div>
              </div>

              {mode === 'ps' ? (
                <div className="input-group">
                  <label className="section-label">2. Word Templates (all carriers variants)</label>
                  <div
                    className="file-drop-zone"
                    onClick={() => psTemplateInputRef.current?.click()}
                    onDragOver={(e) => e.preventDefault()}
                  >
                    <input
                      type="file"
                      ref={psTemplateInputRef}
                      onChange={handlePsTemplateChange}
                      accept=".docx"
                      multiple
                      hidden
                    />
                    <span className="icon">📋</span>
                    <p className="file-name">
                      {psTemplateFiles
                        ? `${psTemplateFiles.length} template${psTemplateFiles.length !== 1 ? 's' : ''} selected`
                        : 'Drop 1–5 carrier templates here (auto-detected by filename)'}
                    </p>
                  </div>
                  {psTemplateFiles && psTemplateFiles.length > 0 && (
                    <ul className="file-list">
                      {Array.from(psTemplateFiles).map((f, i) => (
                        <li key={i}>{f.name}</li>
                      ))}
                    </ul>
                  )}
                </div>
              ) : mode === 'multi' ? (
                <>
                  <div className="input-group">
                    <label className="section-label">2. SCS Template (Word)</label>
                    <div
                      className="file-drop-zone"
                      onClick={() => scsInputRef.current?.click()}
                      onDragOver={(e) => e.preventDefault()}
                    >
                      <input
                        type="file"
                        ref={scsInputRef}
                        onChange={handleScsChange}
                        accept=".docx"
                        hidden
                      />
                      <span className="icon">🧠</span>
                      <p className="file-name">
                        {scsFile ? scsFile.name : "Drop the SCS Word template here"}
                      </p>
                    </div>
                  </div>

                  <div className="input-group">
                    <label className="section-label">3. Default Template (Word)</label>
                    <div
                      className="file-drop-zone"
                      onClick={() => defaultInputRef.current?.click()}
                      onDragOver={(e) => e.preventDefault()}
                    >
                      <input
                        type="file"
                        ref={defaultInputRef}
                        onChange={handleDefaultChange}
                        accept=".docx"
                        hidden
                      />
                      <span className="icon">📝</span>
                      <p className="file-name">
                        {defaultFile ? defaultFile.name : "Drop the Default (non-SCS) Word template here"}
                      </p>
                    </div>
                  </div>
                </>
              ) : (
                <div className="input-group">
                  <label className="section-label">
                    {mode === 'bulk' ? '2. Word Template (single)' : '2. Word Templates'}
                  </label>
                  <div
                    className="file-drop-zone"
                    onClick={() => wordInputRef.current?.click()}
                    onDragOver={(e) => e.preventDefault()}
                  >
                    <input
                      type="file"
                      ref={wordInputRef}
                      onChange={handleWordChange}
                      accept=".docx"
                      multiple={mode === 'single'}
                      hidden
                    />
                    <span className="icon">📝</span>
                    <p className="file-name">
                      {wordFiles
                        ? (wordFiles.length === 1 ? wordFiles[0].name : `${wordFiles.length} files selected`)
                        : (mode === 'bulk'
                            ? "Drop one Word template here"
                            : "Drop Word templates here (multiple allowed)")}
                    </p>
                  </div>
                </div>
              )}
            </div>

            <div className="actions">
              <button type="submit" className="premium-button" disabled={isProcessing}>
                {isProcessing && <div className="loader"></div>}
                <span>{submitLabel}</span>
              </button>
            </div>
          </form>

          {isProcessing && (
            <div className="status-area">
              <div className="progress-bar-container">
                <div className="progress-bar" style={{ width: `${progress}%` }}></div>
              </div>
              <p id="statusText">{status}</p>
            </div>
          )}
          {!isProcessing && status && (
             <div className="status-area">
                <p id="statusText" style={{ color: status.startsWith('Error') ? 'var(--error)' : 'var(--success)' }}>
                    {status}
                </p>
             </div>
          )}
        </section>

        <footer className="footer">
          <p>&copy; 2024 Auto-Filler Engine. All systems operational.</p>
        </footer>
      </main>
    </>
  );
}

export default App;
