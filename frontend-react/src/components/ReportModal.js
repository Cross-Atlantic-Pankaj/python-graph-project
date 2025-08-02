import React, { useState } from 'react';
import styles from './ReportModal.module.css';
import { CircularProgress } from '@mui/material';
import { 
  Description as DescriptionIcon,
  CloudUpload as CloudUploadIcon,
  Close as CloseIcon,
  CheckCircle as CheckCircleIcon,
  Error as ErrorIcon
} from '@mui/icons-material';

const ReportModal = ({ 
  isOpen, 
  onClose, 
  projectName, 
  onGenerateReport, 
  isGenerating = false,
  progress = { message: '', percentage: 0 },
  isBatchGenerating = false,
  batchProgress = { current: 0, total: 0, message: '', percentage: 0 }
}) => {
  const [uploadMode, setUploadMode] = useState('single');
  const [reportFile, setReportFile] = useState(null);
  const [zipFile, setZipFile] = useState(null);

  const handleFileChange = (e, type) => {
    const file = e.target.files[0];
    if (type === 'single') {
      setReportFile(file);
      setZipFile(null);
    } else {
      setZipFile(file);
      setReportFile(null);
    }
  };

  const handleSubmit = () => {
    if (uploadMode === 'single' && reportFile) {
      onGenerateReport(reportFile, 'single');
    } else if (uploadMode === 'batch' && zipFile) {
      onGenerateReport(zipFile, 'batch');
    }
  };

  const handleClose = () => {
    setReportFile(null);
    setZipFile(null);
    onClose();
  };

  if (!isOpen) return null;

  return (
    <div className={styles.modalOverlay}>
      <div className={styles.modalContainer}>
        {/* Header */}
        <div className={styles.modalHeader}>
          <div className={styles.headerContent}>
            <div className={styles.headerIcon}>
              <DescriptionIcon />
            </div>
            <h2 className={styles.modalTitle}>
              Generate Report - {projectName}
            </h2>
          </div>
          <button className={styles.closeButton} onClick={handleClose}>
            <CloseIcon />
          </button>
        </div>

        {/* Tabs */}
        <div className={styles.tabContainer}>
          <button 
            className={`${styles.tab} ${uploadMode === 'single' ? styles.activeTab : ''}`}
            onClick={() => setUploadMode('single')}
          >
            Single Report
          </button>
          <button 
            className={`${styles.tab} ${uploadMode === 'batch' ? styles.activeTab : ''}`}
            onClick={() => setUploadMode('batch')}
          >
            Batch Reports
          </button>
        </div>

        {/* Content */}
        <div className={styles.modalContent}>
          {uploadMode === 'single' ? (
            <div className={styles.uploadSection}>
              <div className={styles.uploadIcon}>
                <CloudUploadIcon />
              </div>
              <h3 className={styles.uploadTitle}>Single File Upload</h3>
              <p className={styles.uploadDescription}>
                Upload a single Excel (.xlsx) or CSV (.csv) file to generate a report.
              </p>
              
              <div className={styles.fileUpload}>
                <input
                  type="file"
                  id="single-file"
                  accept=".xlsx,.csv"
                  onChange={(e) => handleFileChange(e, 'single')}
                  className={styles.fileInput}
                />
                <label htmlFor="single-file" className={styles.fileLabel}>
                  <CloudUploadIcon />
                  <span>{reportFile ? reportFile.name : 'Choose File'}</span>
                </label>
              </div>

              {reportFile && (
                <div className={styles.fileInfo}>
                  <CheckCircleIcon className={styles.fileIcon} />
                  <span>{reportFile.name}</span>
                  <span className={styles.fileSize}>
                    ({(reportFile.size / 1024 / 1024).toFixed(2)} MB)
                  </span>
                </div>
              )}
            </div>
          ) : (
            <div className={styles.uploadSection}>
              <div className={styles.uploadIcon}>
                <CloudUploadIcon />
              </div>
              <h3 className={styles.uploadTitle}>Batch File Upload</h3>
              <p className={styles.uploadDescription}>
                Upload a ZIP file containing multiple Excel files for batch processing.
              </p>
              
              <div className={styles.fileUpload}>
                <input
                  type="file"
                  id="batch-file"
                  accept=".zip"
                  onChange={(e) => handleFileChange(e, 'batch')}
                  className={styles.fileInput}
                />
                <label htmlFor="batch-file" className={styles.fileLabel}>
                  <CloudUploadIcon />
                  <span>{zipFile ? zipFile.name : 'Choose ZIP File'}</span>
                </label>
              </div>

              {zipFile && (
                <div className={styles.fileInfo}>
                  <CheckCircleIcon className={styles.fileIcon} />
                  <span>{zipFile.name}</span>
                  <span className={styles.fileSize}>
                    ({(zipFile.size / 1024 / 1024).toFixed(2)} MB)
                  </span>
                </div>
              )}
            </div>
          )}

          {/* Progress Section */}
          {(isGenerating || isBatchGenerating) && (
            <div className={styles.progressSection}>
              <div className={styles.progressHeader}>
                <CircularProgress size={20} className={styles.progressIcon} />
                <span className={styles.progressText}>
                  {isBatchGenerating ? batchProgress.message : progress.message}
                </span>
                <span className={styles.progressPercentage}>
                  {isBatchGenerating ? batchProgress.percentage : progress.percentage}%
                </span>
              </div>
              <div className={styles.progressBar}>
                <div 
                  className={styles.progressFill}
                  style={{ 
                    width: `${isBatchGenerating ? batchProgress.percentage : progress.percentage}%` 
                  }}
                ></div>
              </div>
              {isBatchGenerating && batchProgress.total > 0 && (
                <div className={styles.batchInfo}>
                  Processing: {batchProgress.current} of {batchProgress.total} files
                </div>
              )}
            </div>
          )}
        </div>

        {/* Footer */}
        <div className={styles.modalFooter}>
          <button 
            className={styles.cancelButton} 
            onClick={handleClose}
            disabled={isGenerating || isBatchGenerating}
          >
            Cancel
          </button>
          <button 
            className={styles.generateButton}
            onClick={handleSubmit}
            disabled={isGenerating || isBatchGenerating || (!reportFile && !zipFile)}
          >
            {isGenerating || isBatchGenerating ? (
              <>
                <CircularProgress size={16} className={styles.buttonSpinner} />
                Generating...
              </>
            ) : (
              'Generate Report'
            )}
          </button>
        </div>
      </div>
    </div>
  );
};

export default ReportModal; 