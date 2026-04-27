import { useState, useRef } from 'react';
import type { ChangeEvent, FormEvent } from 'react';
import axios from 'axios';
import './index.css';

const API_BASE = 'http://127.0.0.1:5001/api';
const SINGLE_URL = `${API_BASE}/process`;
const BULK_URL = `${API_BASE}/bulk`;

type Mode = 'single' | 'bulk';

function App() {
  const [excelFile, setExcelFile] = useState<File | null>(null);
  const [wordFiles, setWordFiles] = useState<FileList | null>(null);
  const [mode, setMode] = useState<Mode>('single');
  const [isProcessing, setIsProcessing] = useState(false);
  const [status, setStatus] = useState('');
  const [progress, setProgress] = useState(0);

  const excelInputRef = useRef<HTMLInputElement>(null);
  const wordInputRef = useRef<HTMLInputElement>(null);

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

  const handleSubmit = async (e: FormEvent) => {
    e.preventDefault();
    if (!excelFile || !wordFiles) {
      alert("Please select both an Excel file and Word templates.");
      return;
    }

    if (mode === 'bulk' && wordFiles.length !== 1) {
      alert("Bulk mode expects exactly one Word template.");
      return;
    }

    setIsProcessing(true);
    setStatus('Uploading and processing...');
    setProgress(30);

    const formData = new FormData();
    formData.append('excel', excelFile);
    Array.from(wordFiles).forEach((file) => {
      formData.append('word', file);
    });

    const endpoint = mode === 'bulk' ? BULK_URL : SINGLE_URL;

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
      setStatus('Success! Download started.');

      // Handle download
      const contentDisposition = response.headers['content-disposition'];
      let filename: string;
      if (mode === 'bulk') {
        filename = `${wordFiles[0].name.split('.')[0]}_bulk_filled.zip`;
      } else {
        filename = wordFiles.length === 1
          ? `${wordFiles[0].name.split('.')[0]}_filled.docx`
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
      }, 3000);

    } catch (error: any) {
      console.error('Error processing files:', error);
      setStatus(`Error: ${error.response?.data?.error || error.message}`);
      setIsProcessing(false);
    }
  };

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
            </div>
            <p className="mode-hint">
              {mode === 'single'
                ? 'Uses the "Fields to Enter" sheet (one placeholder per row).'
                : 'Uses the "Export" sheet — placeholders in row 1, one filled doc per data row.'}
            </p>
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
            </div>

            <div className="actions">
              <button type="submit" className="premium-button" disabled={isProcessing}>
                {isProcessing && <div className="loader"></div>}
                <span>
                  {isProcessing
                    ? "Processing..."
                    : mode === 'bulk'
                      ? "Run Bulk Sweep"
                      : "Generate Documents"}
                </span>
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
