import React, { useState } from 'react';
import { UploadCloud, FileSpreadsheet, CheckCircle2, AlertCircle, Download, Loader2 } from 'lucide-react';
import { processExcelFiles } from './utils/excelProcessor';
import './index.css';

function App() {
  const [files, setFiles] = useState({
    gt: null,
    nlc: null,
    item: null
  });
  const [isProcessing, setIsProcessing] = useState(false);
  const [error, setError] = useState(null);
  const [downloadUrl, setDownloadUrl] = useState(null);

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
      const blob = await processExcelFiles(files.gt, files.nlc, files.item);
      const url = URL.createObjectURL(blob);
      setDownloadUrl(url);
    } catch (err) {
      setError(err.message || "An error occurred while processing the files.");
    } finally {
      setIsProcessing(false);
    }
  };

  const FileUploadBox = ({ title, type, file, id }) => (
    <div className="upload-box">
      <input
        type="file"
        id={id}
        accept=".xlsx, .xls, .csv"
        onChange={(e) => handleFileChange(e, type)}
        className="hidden-input"
      />
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
                {isProcessing ? (
                  <>
                    <Loader2 className="spinner" />
                    Processing Data...
                  </>
                ) : (
                  <>
                    <UploadCloud size={20} />
                    Process Excel Files
                  </>
                )}
              </button>
            ) : (
              <a href={downloadUrl} download={`NLC_Processed_${new Date().getTime()}.xlsx`} className="download-btn">
                <Download size={20} />
                Download Processed File
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
