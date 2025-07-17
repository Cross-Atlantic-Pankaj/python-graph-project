import React, { useState, useEffect } from 'react';
import { useNavigate } from 'react-router-dom';
import {
  Container,
  Typography,
  Button,
  Box,
  IconButton,
  Dialog,
  DialogTitle,
  DialogContent,
  DialogActions,
  TextField,
  Table,
  TableBody,
  TableCell,
  TableContainer,
  TableHead,
  TableRow,
  Paper,
  CircularProgress,
  Tabs,
  Tab,
  Alert,
  Accordion,
  AccordionSummary,
  AccordionDetails,
  Chip,
  Divider,
  List,
  ListItem,
  ListItemText,
  ListItemIcon,
  LinearProgress,
  Snackbar,
} from '@mui/material';
import { 
  Add as AddIcon, 
  Logout as LogoutIcon, 
  Description as DescriptionIcon,
  Error as ErrorIcon,
  Warning as WarningIcon,
  Info as InfoIcon,
  ExpandMore as ExpandMoreIcon,
  BugReport as BugReportIcon,
  Code as CodeIcon,
  Settings as SettingsIcon
} from '@mui/icons-material';
import axios from 'axios';

axios.defaults.withCredentials = true;

function Dashboard() {
  const navigate = useNavigate();
  const [projects, setProjects] = useState([]);
  const [openCreateProjectDialog, setOpenCreateProjectDialog] = useState(false);
  const [openReportUploadDialog, setOpenReportUploadDialog] = useState(false);
  const [newProject, setNewProject] = useState({ name: '', description: '', file: null });
  const [selectedProjectForReport, setSelectedProjectForReport] = useState(null);
  const [reportFile, setReportFile] = useState(null);
  const [zipFile, setZipFile] = useState(null);
  const [user, setUser] = useState(null);
  const [isGeneratingReport, setIsGeneratingReport] = useState(false);
  const [isBatchGenerating, setIsBatchGenerating] = useState(false);
  const [batchProgress, setBatchProgress] = useState({ current: 0, total: 0, message: '', percentage: 0 });
  const [singleProgress, setSingleProgress] = useState({ message: '', percentage: 0 });
  const [uploadMode, setUploadMode] = useState('single'); // 'single' or 'batch'
  const [chartErrors, setChartErrors] = useState({});
  const [showErrorDialog, setShowErrorDialog] = useState(false);
  const [selectedProjectForErrors, setSelectedProjectForErrors] = useState(null);
  const [customAlert, setCustomAlert] = useState({
    open: false,
    message: '',
    severity: 'success', // 'success', 'warning', 'error'
    title: ''
  });

  useEffect(() => {
    loadUser();
    loadProjects();
  }, []);

  const loadUser = async () => {
    try {
      const response = await axios.get(`${process.env.REACT_APP_API_URL}/api/user`);
      setUser(response.data.user);
    } catch (error) {
      navigate('/login');
    }
  };

  const loadProjects = async () => {
    try {
      const response = await axios.get(`${process.env.REACT_APP_API_URL}/api/projects`);
      setProjects(response.data.projects);
    } catch (error) {
      console.error('Error loading projects:', error);
    }
  };

  const handleCreateProject = async () => {
    try {
      const formData = new FormData();
      formData.append('name', newProject.name);
      formData.append('description', newProject.description);
      if (newProject.file) {
        formData.append('file', newProject.file);
      }
  
      await axios.post(`${process.env.REACT_APP_API_URL}/api/projects`, formData, {
        headers: {
          'Content-Type': 'multipart/form-data',
        },
      });
  
      setOpenCreateProjectDialog(false);
      setNewProject({ name: '', description: '', file: null });
      loadProjects();
    } catch (error) {
      console.error('Error creating project:', error.response?.data || error.message);
    }
  };

  const handleOpenReportUploadDialog = (project) => {
    setSelectedProjectForReport(project);
    setOpenReportUploadDialog(true);
  };

  const handleCloseReportUploadDialog = () => {
    setOpenReportUploadDialog(false);
    setSelectedProjectForReport(null);
    setReportFile(null);
    setZipFile(null);
    setBatchProgress({ current: 0, total: 0, message: '', percentage: 0 });
    setSingleProgress({ message: '', percentage: 0 });
    // Clear any existing chart errors when closing dialog
    setChartErrors({});
  };

  const handleShowChartErrors = async (project) => {
    setSelectedProjectForErrors(project);
    try {
      const response = await axios.get(`${process.env.REACT_APP_API_URL}/api/projects/${project.id}/chart_errors`);
      setChartErrors(response.data);
      setShowErrorDialog(true);
    } catch (error) {
      console.error('Error fetching chart errors:', error);
      setChartErrors({ error: 'Failed to fetch chart errors' });
      setShowErrorDialog(true);
    }
  };

  const clearProjectErrors = async (projectId) => {
    try {
      // Clear errors by making a request to reset them
      await axios.post(`${process.env.REACT_APP_API_URL}/api/projects/${projectId}/clear_errors`);
    } catch (error) {
      console.error('Error clearing project errors:', error);
    }
  };

  const handleCloseErrorDialog = () => {
    setShowErrorDialog(false);
    setSelectedProjectForErrors(null);
    setChartErrors({});
  };

  const showCustomAlert = (title, message, severity = 'success') => {
    setCustomAlert({
      open: true,
      title,
      message,
      severity
    });
  };

  const handleCloseCustomAlert = () => {
    setCustomAlert(prev => ({ ...prev, open: false }));
  };

  const handleReportFileUpload = async () => {
    if ((!reportFile && !zipFile) || !selectedProjectForReport) {
      alert('Please select a file to upload.');
      return;
    }

    if (zipFile) {
      // Batch processing
      setIsBatchGenerating(true);
      setBatchProgress({ current: 0, total: 0, message: 'Uploading ZIP file...', percentage: 10 });

      try {
        const formData = new FormData();
        formData.append('zip_file', zipFile);

        // Step 1: Upload ZIP and trigger batch report generation
        setBatchProgress({ current: 0, total: 0, message: 'Processing ZIP file...', percentage: 20 });
        
        const response = await axios.post(
          `${process.env.REACT_APP_API_URL}/api/projects/${selectedProjectForReport.id}/upload_zip`,
          formData,
          {
            headers: { 'Content-Type': 'multipart/form-data' },
          }
        );

        // Extract progress info from response
        const { total_files, processed_files } = response.data;
        const percentage = Math.round((processed_files / total_files) * 100);
        
        setBatchProgress({ 
          current: processed_files, 
          total: total_files, 
          message: `Generated ${processed_files} of ${total_files} reports. Downloading ZIP...`, 
          percentage: Math.min(percentage, 90) 
        });

        // Step 2: Download the resulting ZIP
        setBatchProgress({ 
          current: processed_files, 
          total: total_files, 
          message: 'Downloading generated reports...', 
          percentage: 95 
        });
        
        const downloadResponse = await axios.get(
          `${process.env.REACT_APP_API_URL}/api/reports/batch_reports_${selectedProjectForReport.id}.zip`,
          { responseType: 'blob' }
        );

        const blob = new Blob([downloadResponse.data]);
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = `batch_reports_${selectedProjectForReport.id}.zip`;
        link.click();
        URL.revokeObjectURL(link.href);

        setBatchProgress({ 
          current: processed_files, 
          total: total_files, 
          message: 'Batch reports downloaded successfully!', 
          percentage: 100 
        });
        
        setTimeout(() => {
          showCustomAlert(
            'Batch Processing Complete! 🎉',
            `Successfully generated and downloaded ${processed_files} out of ${total_files} reports.`,
            'success'
          );
          handleCloseReportUploadDialog();
        }, 1000);
        
      } catch (error) {
        console.error('Batch report error:', error.response?.data || error.message);
        showCustomAlert(
          'Batch Processing Failed! ❌',
          'Failed to process ZIP file. Please check your file and try again.',
          'error'
        );
      } finally {
        setIsBatchGenerating(false);
        setBatchProgress({ current: 0, total: 0, message: '', percentage: 0 });
      }
      return;
    }

    // Single file processing (existing logic)
    setIsGeneratingReport(true);
    setSingleProgress({ message: 'Uploading Excel file...', percentage: 20 });

    try {
      const formData = new FormData();
      formData.append('report_file', reportFile);

      setSingleProgress({ message: 'Processing Excel data...', percentage: 40 });
      
      await axios.post(
        `${process.env.REACT_APP_API_URL}/api/projects/${selectedProjectForReport.id}/upload_report`,
        formData,
        {
          headers: { 'Content-Type': 'multipart/form-data' },
        }
      );
      
      setSingleProgress({ message: 'Generating charts and report...', percentage: 70 });
      // Check for chart errors after generation
      try {
        const errorResponse = await axios.get(`${process.env.REACT_APP_API_URL}/api/projects/${selectedProjectForReport.id}/chart_errors`);
        const errors = errorResponse.data;
        
        const chartErrorCount = Object.keys(errors.chart_generation_errors || {}).length;
        const reportErrorCount = (errors.report_generation_errors || []).length;
        const totalErrors = chartErrorCount + reportErrorCount;
        
        if (totalErrors === 0) {
          // All charts generated successfully - Green alert
          showCustomAlert(
            'Report Generated Successfully! 🎉',
            'All charts were generated without any errors.',
            'success'
          );
        } else if (totalErrors < 5) { // Assuming reasonable threshold for "some" errors
          // Some charts failed - Yellow warning alert
          showCustomAlert(
            'Report Generated with Warnings ⚠️',
            `${totalErrors} chart(s) failed to generate. The report was created but some charts may be missing. Click "View Errors" for details.`,
            'warning'
          );
        } else {
          // Many charts failed - Red error alert
          showCustomAlert(
            'Report Generation Issues ❌',
            `${totalErrors} charts failed to generate. The report may be incomplete. Click "View Errors" to see what went wrong.`,
            'error'
          );
        }
      } catch (errorCheckError) {
        console.error('Error checking for chart errors:', errorCheckError);
        showCustomAlert(
          'Report Generated! 📄',
          'Report was created successfully, but error checking failed.',
          'warning'
        );
      }

      setSingleProgress({ message: 'Downloading generated report...', percentage: 90 });
      
      // Download the generated report
      try {
        const downloadResponse = await axios.get(
          `${process.env.REACT_APP_API_URL}/api/reports/${selectedProjectForReport.id}/download`,
          {
            responseType: 'blob',
            withCredentials: true,
          }
        );

        const blob = new Blob([downloadResponse.data]);
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);

        const contentDisposition = downloadResponse.headers['content-disposition'];
        let filename = `${selectedProjectForReport.name}_report.docx`;
        if (contentDisposition) {
          const filenameMatch = contentDisposition.match(/filename="(.+)"/);
          if (filenameMatch && filenameMatch[1]) {
            filename = decodeURIComponent(filenameMatch[1]);
          }
        }
        link.download = filename;
        link.click();
        URL.revokeObjectURL(link.href);
        
        setSingleProgress({ message: 'Report downloaded successfully!', percentage: 100 });
      } catch (downloadError) {
        console.error('Error downloading report:', downloadError.response?.data || downloadError.message);
        showCustomAlert(
          'Download Failed! ⚠️',
          'Report was generated but failed to download. Please try again.',
          'warning'
        );
      }

      setTimeout(() => {
        handleCloseReportUploadDialog();
      }, 1000);
    } catch (uploadError) {
      console.error('Error uploading report:', uploadError.response?.data || uploadError.message);
      showCustomAlert(
        'Upload Failed! ❌',
        'Failed to upload report to server. Please check your file and try again.',
        'error'
      );
    } finally {
      setIsGeneratingReport(false);
      setSingleProgress({ message: '', percentage: 0 });
    }
  };

  const handleLogout = async () => {
    try {
      await axios.get(`${process.env.REACT_APP_API_URL}/api/logout`);
      navigate('/login');
    } catch (error) {
      console.error('Error logging out:', error);
    }
  };

  return (
    <Container maxWidth="lg" sx={{ py: 4 }}>
      <Box sx={{ display: 'flex', justifyContent: 'space-between', mb: 4 }}>
        <Typography variant="h4" component="h1">
          My Projects
        </Typography>
        <Box sx={{ display: 'flex', alignItems: 'center', gap: 2 }}>
          <Typography variant="body1">
            Welcome, {user?.full_name}
          </Typography>
          <IconButton onClick={handleLogout} color="error">
            <LogoutIcon />
          </IconButton>
        </Box>
      </Box>

      <Button
        variant="contained"
        startIcon={<AddIcon />}
        onClick={() => setOpenCreateProjectDialog(true)}
        sx={{ mb: 4 }}
      >
        Create New Project
      </Button>

      <TableContainer component={Paper}>
        <Table sx={{ minWidth: 650 }} aria-label="projects table">
          <TableHead>
            <TableRow>
              <TableCell>Project Name</TableCell>
              <TableCell>Description</TableCell>
              <TableCell>Created On</TableCell>
              <TableCell align="right">Actions</TableCell>
            </TableRow>
          </TableHead>
          <TableBody>
            {projects.length === 0 ? (
              <TableRow>
                <TableCell colSpan={4} align="center">
                  No projects yet. Create your first project!
                </TableCell>
              </TableRow>
            ) : (
              projects.map((project) => (
                <TableRow key={project.id}>
                  <TableCell component="th" scope="row">
                    {project.name}
                  </TableCell>
                  <TableCell>{project.description}</TableCell>
                  <TableCell>{new Date(project.created_at).toLocaleDateString()}</TableCell>
                  <TableCell align="right">
                    <Box sx={{ display: 'flex', gap: 1 }}>
                      <Button
                        variant="outlined"
                        startIcon={<DescriptionIcon />}
                        onClick={() => handleOpenReportUploadDialog(project)}
                      >
                        Generate Report
                      </Button>
                      <Button
                        variant="outlined"
                        color="warning"
                        startIcon={<BugReportIcon />}
                        onClick={() => handleShowChartErrors(project)}
                      >
                        View Errors
                      </Button>
                    </Box>
                  </TableCell>
                </TableRow>
              ))
            )}
          </TableBody>
        </Table>
      </TableContainer>

      {/* Create Project Dialog */}
      <Dialog open={openCreateProjectDialog} onClose={() => setOpenCreateProjectDialog(false)}>
        <DialogTitle>Create New Project</DialogTitle>
        <DialogContent>
          <TextField
            autoFocus
            margin="dense"
            label="Project Name"
            fullWidth
            value={newProject.name}
            onChange={(e) => setNewProject({ ...newProject, name: e.target.value })}
          />
          <TextField
            margin="dense"
            label="Description"
            fullWidth
            multiline
            rows={4}
            value={newProject.description}
            onChange={(e) => setNewProject({ ...newProject, description: e.target.value })}
          />
          <input
            type="file"
            accept=".doc,.docx"
            onChange={(e) => setNewProject({ ...newProject, file: e.target.files[0] })}
            style={{ marginTop: '16px' }}
          />
        </DialogContent>
        <DialogActions>
          <Button onClick={() => setOpenCreateProjectDialog(false)}>Cancel</Button>
          <Button onClick={handleCreateProject} variant="contained">
            Create
          </Button>
        </DialogActions>
      </Dialog>

      {/* Report Upload Dialog */}
      <Dialog open={openReportUploadDialog} onClose={handleCloseReportUploadDialog}>
        <DialogTitle>Upload Report File for "{selectedProjectForReport?.name}"</DialogTitle>
        <DialogContent>
          <Tabs
            value={uploadMode}
            onChange={(e, val) => {
              setUploadMode(val);
              setReportFile(null);
              setZipFile(null);
            }}
            sx={{ mb: 2 }}
          >
            <Tab label="Single Report" value="single" />
            <Tab label="Batch Reports" value="batch" />
          </Tabs>
          {uploadMode === 'single' && (
            <>
              <Typography variant="body2" sx={{ mb: 1 }}>
                Upload a single Excel (.xlsx) or CSV (.csv) file.
          </Typography>
          <input
            type="file"
            accept=".xlsx,.csv"
            onChange={(e) => setReportFile(e.target.files[0])}
                style={{ marginTop: '8px' }}
              />
            </>
          )}
          {uploadMode === 'batch' && (
            <>
              <Typography variant="body2" sx={{ mb: 1 }}>
                Upload a ZIP file containing multiple Excel files for batch processing.
              </Typography>
              <input
                type="file"
                accept=".zip"
                onChange={(e) => setZipFile(e.target.files[0])}
                style={{ marginTop: '8px' }}
          />
            </>
          )}
          {isBatchGenerating && (
            <Box sx={{ mt: 2 }}>
              <Box sx={{ display: 'flex', alignItems: 'center', mb: 1 }}>
                <CircularProgress size={20} sx={{ mr: 1 }} />
                <Typography variant="body2" sx={{ flexGrow: 1 }}>
                  {batchProgress.message}
                </Typography>
                <Typography variant="body2" color="primary" fontWeight="bold">
                  {batchProgress.percentage}%
                </Typography>
              </Box>
              <LinearProgress 
                variant="determinate" 
                value={batchProgress.percentage} 
                sx={{ height: 8, borderRadius: 4 }}
              />
              {batchProgress.total > 0 && (
                <Typography variant="caption" color="text.secondary" sx={{ mt: 0.5, display: 'block' }}>
                  Progress: {batchProgress.current} of {batchProgress.total} files
                </Typography>
              )}
            </Box>
          )}
          
          {isGeneratingReport && uploadMode === 'single' && (
            <Box sx={{ mt: 2 }}>
              <Box sx={{ display: 'flex', alignItems: 'center', mb: 1 }}>
                <CircularProgress size={20} sx={{ mr: 1 }} />
                <Typography variant="body2" sx={{ flexGrow: 1 }}>
                  {singleProgress.message}
                </Typography>
                <Typography variant="body2" color="primary" fontWeight="bold">
                  {singleProgress.percentage}%
                </Typography>
              </Box>
              <LinearProgress 
                variant="determinate" 
                value={singleProgress.percentage} 
                sx={{ height: 8, borderRadius: 4 }}
              />
            </Box>
          )}
        </DialogContent>
        <DialogActions>
          <Button onClick={handleCloseReportUploadDialog} disabled={isGeneratingReport || isBatchGenerating}>Cancel</Button>
          <Button
            onClick={handleReportFileUpload}
            variant="contained"
            disabled={isGeneratingReport || isBatchGenerating}
            startIcon={(isGeneratingReport || isBatchGenerating) ? <CircularProgress size={20} color="inherit" /> : null}
          >
            {(isGeneratingReport || isBatchGenerating) ? 'Generating...' : 'Generate Report'}
          </Button>
        </DialogActions>
      </Dialog>

      {/* Chart Errors Dialog */}
      <Dialog 
        open={showErrorDialog} 
        onClose={handleCloseErrorDialog}
        maxWidth="sm"
        fullWidth
      >
        <DialogTitle sx={{ display: 'flex', alignItems: 'center', gap: 1 }}>
          <ErrorIcon color="error" />
          Chart Issues - {selectedProjectForErrors?.name}
        </DialogTitle>
        <DialogContent>
          {chartErrors.error ? (
            <Alert severity="error" sx={{ mb: 2 }}>
              {chartErrors.error}
            </Alert>
          ) : (
            <Box>
              {/* Chart Generation Errors */}
              {chartErrors.chart_generation_errors && Object.keys(chartErrors.chart_generation_errors).length > 0 && (
                <Box sx={{ mb: 3 }}>
                  <Typography variant="h6" sx={{ display: 'flex', alignItems: 'center', gap: 1, mb: 2 }}>
                    <ErrorIcon color="error" />
                    Chart Problems ({Object.keys(chartErrors.chart_generation_errors).length})
                  </Typography>
                  <List>
                    {Object.entries(chartErrors.chart_generation_errors).map(([tag, error], index) => (
                      <ListItem key={index} sx={{ flexDirection: 'column', alignItems: 'flex-start', p: 0, mb: 2 }}>
                        <Box sx={{ display: 'flex', alignItems: 'center', gap: 1, mb: 1 }}>
                          <ErrorIcon color="error" fontSize="small" />
                          <Typography variant="subtitle1" fontWeight="bold">
                            {tag}
                          </Typography>
                          <Chip 
                            label={error.chart_type} 
                            color="primary" 
                            size="small" 
                            variant="outlined"
                          />
                        </Box>
                        <Alert severity="error" sx={{ width: '100%', mb: 1 }}>
                          {error.user_message}
                        </Alert>
                        <Box sx={{ display: 'flex', gap: 1, fontSize: '0.8rem', color: 'text.secondary' }}>
                          <span>Type: {error.error_type}</span>
                          <span>•</span>
                          <span>Data points: {error.data_points}</span>
                        </Box>
                      </ListItem>
                    ))}
                  </List>
                </Box>
              )}

                            {/* Report Generation Errors */}
              {chartErrors.report_generation_errors && chartErrors.report_generation_errors.length > 0 && (
                <Box sx={{ mb: 3 }}>
                  <Typography variant="h6" sx={{ display: 'flex', alignItems: 'center', gap: 1, mb: 2 }}>
                    <WarningIcon color="warning" />
                    Report Issues ({chartErrors.report_generation_errors.length})
                  </Typography>
                  <List>
                    {chartErrors.report_generation_errors.map((error, index) => (
                      <ListItem key={index} sx={{ flexDirection: 'column', alignItems: 'flex-start', p: 0, mb: 2 }}>
                        <Box sx={{ display: 'flex', alignItems: 'center', gap: 1, mb: 1 }}>
                          <WarningIcon color="warning" fontSize="small" />
                          <Typography variant="subtitle1" fontWeight="bold">
                            {error.tag}
                          </Typography>
                          {chartErrors.report_generation_errors_detailed && 
                           chartErrors.report_generation_errors_detailed[error.tag] && (
                            <Chip 
                              label={chartErrors.report_generation_errors_detailed[error.tag].chart_type} 
                              color="primary" 
                              size="small" 
                              variant="outlined"
                            />
                          )}
                        </Box>
                        <Alert severity="warning" sx={{ width: '100%', mb: 1 }}>
                          {error.error}
                        </Alert>
                        <Typography variant="caption" color="text.secondary">
                          This chart could not be inserted into the report document.
                        </Typography>
                      </ListItem>
                    ))}
                  </List>
                </Box>
              )}

              {/* No Errors */}
              {(!chartErrors.chart_generation_errors || Object.keys(chartErrors.chart_generation_errors).length === 0) &&
               (!chartErrors.report_generation_errors || chartErrors.report_generation_errors.length === 0) && (
                <Alert severity="success" sx={{ mt: 2 }}>
                  <Box sx={{ display: 'flex', alignItems: 'center', gap: 1 }}>
                    <InfoIcon />
                    <Typography>All charts generated successfully! 🎉</Typography>
                  </Box>
                </Alert>
              )}

              {/* Report Generation Info */}
              {chartErrors.report_generated_at && (
                <Alert severity="info" sx={{ mt: 2 }}>
                  <Typography variant="body2">
                    Last report: {new Date(chartErrors.report_generated_at).toLocaleString()}
                  </Typography>
                </Alert>
              )}
            </Box>
          )}
        </DialogContent>
        <DialogActions>
          <Button 
            onClick={() => {
              clearProjectErrors(selectedProjectForErrors?.id);
              handleCloseErrorDialog();
            }}
            color="warning"
          >
            Clear Errors
          </Button>
          <Button onClick={handleCloseErrorDialog}>Close</Button>
        </DialogActions>
      </Dialog>

      {/* Custom Alert Snackbar */}
      <Snackbar
        open={customAlert.open}
        autoHideDuration={6000}
        onClose={handleCloseCustomAlert}
        anchorOrigin={{ vertical: 'top', horizontal: 'center' }}
      >
        <Alert
          onClose={handleCloseCustomAlert}
          severity={customAlert.severity}
          variant="filled"
          sx={{
            width: '100%',
            maxWidth: '500px',
            '& .MuiAlert-message': {
              width: '100%'
            }
          }}
        >
          <Box>
            <Typography variant="h6" sx={{ fontWeight: 'bold', mb: 0.5 }}>
              {customAlert.title}
            </Typography>
            <Typography variant="body2">
              {customAlert.message}
            </Typography>
          </Box>
        </Alert>
      </Snackbar>
    </Container>
  );
}

export default Dashboard; 