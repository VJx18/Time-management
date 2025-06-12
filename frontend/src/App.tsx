import React, { useState } from 'react';
import './App.scss';

// TypeScript interfaces for type safety
interface ProcessingResponse {
  success: boolean;
  message: string;
  employees_processed?: number;
  download_url?: string;
  filename?: string;
}

interface UploadState {
  file: File | null;
  isProcessing: boolean;
  processedFileUrl: string | null;
  processedFileName: string | null;
  error: string | null;
  successMessage: string | null;
  employeesProcessed: number | null;
}

function App() {
  // State management for file upload and processing
  const [uploadState, setUploadState] = useState<UploadState>({
    file: null,
    isProcessing: false,
    processedFileUrl: null,
    processedFileName: null,
    error: null,
    successMessage: null,
    employeesProcessed: null
  });

  // Backend server URL
  const API_BASE_URL = 'http://localhost:8000';

  /**
   * Handle file selection from input
   */
  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const selectedFile = e.target.files?.[0] || null;
    
    // Reset previous results when new file is selected
    setUploadState({
      file: selectedFile,
      isProcessing: false,
      processedFileUrl: null,
      processedFileName: null,
      error: null,
      successMessage: null,
      employeesProcessed: null
    });
  };

  /**
   * Validate file before upload
   */
  const validateFile = (file: File): string | null => {
    // Check file type
    if (!file.name.toLowerCase().endsWith('.xlsx')) {
      return 'Please select an Excel file (.xlsx format only)';
    }

    // Check file size (16MB limit to match backend)
    const maxSize = 16 * 1024 * 1024; // 16MB
    if (file.size > maxSize) {
      return 'File size too large. Maximum size is 16MB';
    }

    return null; // No validation errors
  };

  /**
   * Upload and process the selected file
   */
  const handleUploadAndProcess = async () => {
    if (!uploadState.file) {
      setUploadState(prev => ({
        ...prev,
        error: 'Please select a file to upload'
      }));
      return;
    }

    // Validate file
    const validationError = validateFile(uploadState.file);
    if (validationError) {
      setUploadState(prev => ({
        ...prev,
        error: validationError
      }));
      return;
    }

    // Start processing
    setUploadState(prev => ({
      ...prev,
      isProcessing: true,
      error: null,
      successMessage: null
    }));

    try {
      // Create FormData for file upload
      const formData = new FormData();
      formData.append('file', uploadState.file);
      
      // Optional: Add max employees parameter
      // formData.append('max_employees', '30');

      // Send request to Flask backend
      const response = await fetch(`${API_BASE_URL}/process-holidays`, {
        method: 'POST',
        body: formData,
      });

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      const result: ProcessingResponse = await response.json();

      if (result.success) {
        // Success - update state with download info
        setUploadState(prev => ({
          ...prev,
          isProcessing: false,
          processedFileUrl: result.download_url || null,
          processedFileName: result.filename || null,
          successMessage: result.message || 'File processed successfully!',
          employeesProcessed: result.employees_processed || null,
          error: null
        }));
      } else {
        // Backend returned error
        setUploadState(prev => ({
          ...prev,
          isProcessing: false,
          error: result.message || 'Processing failed'
        }));
      }
    } catch (err) {
      // Network or other errors
      console.error('Upload error:', err);
      setUploadState(prev => ({
        ...prev,
        isProcessing: false,
        error: err instanceof Error ? err.message : 'An error occurred during processing'
      }));
    }
  };

  /**
   * Download the processed file
   */
  const handleDownload = async () => {
    if (!uploadState.processedFileUrl) return;

    try {
      const response = await fetch(`${API_BASE_URL}${uploadState.processedFileUrl}`);
      
      if (!response.ok) {
        throw new Error('Download failed');
      }

      // Create blob and download
      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = uploadState.processedFileName || 'processed_holidays.xlsx';
      document.body.appendChild(a);
      a.click();
      
      // Cleanup
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);
    
    // eslint-disable-next-line @typescript-eslint/no-unused-vars
    } catch (err) {
      setUploadState(prev => ({
        ...prev,
        error: 'Failed to download file'
      }));
    }
  };

  /**
   * Reset form to initial state
   */
  const handleReset = () => {
    setUploadState({
      file: null,
      isProcessing: false,
      processedFileUrl: null,
      processedFileName: null,
      error: null,
      successMessage: null,
      employeesProcessed: null
    });
  };

  return (
    <div className="app-container">
      <header className="app-header">
        <h1>Time Distribution System</h1>
      </header>

      <main className="main-content">
        <section className="upload-section">
          <h2>Upload Employee Data</h2>
          <p>
            Upload your Excel file with employee data. The file must contain both 
            <strong> "MA Übersicht"</strong> and <strong>"IST Stunden"</strong> sheets.
          </p>
          
          {/* File Input */}
          <div className="file-input-container">
            <input 
              type="file" 
              accept=".xlsx"
              onChange={handleFileChange}
              disabled={uploadState.isProcessing}
              className="file-input"
              id="file-upload"
            />
            <label htmlFor="file-upload" className="file-input-label">
              {uploadState.file ? uploadState.file.name : 'Choose Excel File (.xlsx)'}
            </label>
          </div>

          {/* File Info */}
          {uploadState.file && (
            <div className="file-info">
              <p><strong>Selected:</strong> {uploadState.file.name}</p>
              <p><strong>Size:</strong> {(uploadState.file.size / 1024 / 1024).toFixed(2)} MB</p>
            </div>
          )}
          
          {/* Process Button */}
          <button 
            onClick={handleUploadAndProcess}
            disabled={!uploadState.file || uploadState.isProcessing}
            className="process-button"
          >
            {uploadState.isProcessing ? (
              <>
                <span className="spinner"></span>
                Processing...
              </>
            ) : (
              'Process Time Distribution'
            )}
          </button>
        </section>

        {/* Error Message */}
        {uploadState.error && (
          <section className="error-section">
            <h3> Error</h3>
            <p>{uploadState.error}</p>
            <button onClick={handleReset} className="reset-button">
              Try Again
            </button>
          </section>
        )}

        {/* Success Section */}
        {uploadState.successMessage && uploadState.processedFileUrl && (
          <section className="success-section">
            <h3>Processing Complete!</h3>
            <div className="success-content">
              <p >{uploadState.successMessage}</p>
              
              {uploadState.employeesProcessed && (
                <p className="employees-count">
                  <strong>Employees Processed:</strong> {uploadState.employeesProcessed}
                </p>
              )}
              
              <div className="download-section">
                <button onClick={handleDownload} className="download-button">
                   Download Processed File
                </button>
                {uploadState.processedFileName && (
                  <p className="filename">File: {uploadState.processedFileName}</p>
                )}
              </div>
              
              <button onClick={handleReset} className="reset-button">
                Process Another File
              </button>
            </div>
          </section>
        )}

        {/* Instructions */}
        {/* <section className="instructions-section">
          <h3>📋 Instructions</h3>
          <div className="instructions-content">
            <h4>Required Excel File Structure:</h4>
            <ul>
              <li><strong>MA Übersicht</strong> - Employee overview with names, states, and dates</li>
              <li><strong>IST Stunden</strong> - Actual working hours data</li>
            </ul>
            
            <h4>What This Tool Does:</h4>
            <ul>
              <li>Reads German employee data from your Excel file</li>
              <li>Applies state-specific public holiday rules</li>
              <li>Marks holidays in the working hours sheet</li>
              <li>Generates a processed file with holiday adjustments</li>
            </ul>
            
            <h4>File Requirements:</h4>
            <ul>
              <li>Maximum 30 employees per file</li>
              <li>Excel format (.xlsx only)</li>
              <li>Maximum file size: 16MB</li>
              <li>German sheet names required</li>
            </ul>
          </div>
        </section> */}
      </main>

      <footer className="app-footer">
        <p>Holiday Distribution System - Built with React & Flask</p>
      </footer>
    </div>
  );
}

export default App;