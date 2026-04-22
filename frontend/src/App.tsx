import { useState, useRef, ChangeEvent, FormEvent } from 'react';
import axios from 'axios';
import './index.css';

const API_URL = 'http://127.0.0.1:5000/api/process';

function App() {
  const [excelFile, setExcelFile] = useState<File | null>(null);
  const [wordFiles, setWordFiles] = useState<FileList | null>(null);
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

    setIsProcessing(true);
    setStatus('Uploading and processing...');
    setProgress(30);

    const formData = new FormData();
    formData.append('excel', excelFile);
    Array.from(wordFiles).forEach((file) => {
      formData.append('word', file);
    });

    try {
      const response = await axios.post(API_URL, formData, {
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
      let filename = wordFiles.length === 1 ? `${wordFiles[0].name.split('.')[0]}_filled.docx` : 'filled_documents.zip';
      
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
                <label className="section-label">2. Word Templates</label>
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
                    multiple 
                    hidden 
                  />
                  <span className="icon">📝</span>
                  <p className="file-name">
                    {wordFiles 
                      ? (wordFiles.length === 1 ? wordFiles[0].name : `${wordFiles.length} files selected`) 
                      : "Drop Word templates here (multiple allowed)"}
                  </p>
                </div>
              </div>
            </div>

            <div className="actions">
              <button type="submit" className="premium-button" disabled={isProcessing}>
                {isProcessing && <div className="loader"></div>}
                <span>{isProcessing ? "Processing..." : "Generate Documents"}</span>
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
