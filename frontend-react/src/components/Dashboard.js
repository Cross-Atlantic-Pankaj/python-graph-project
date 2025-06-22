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
} from '@mui/material';
import { Add as AddIcon, Logout as LogoutIcon, Description as DescriptionIcon } from '@mui/icons-material';
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
  const [batchProgress, setBatchProgress] = useState({ current: 0, total: 0, message: '' });
  const [uploadMode, setUploadMode] = useState('single'); // 'single' or 'batch'

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
  };

  const handleReportFileUpload = async () => {
    if ((!reportFile && !zipFile) || !selectedProjectForReport) {
      alert('Please select a file to upload.');
      return;
    }

    if (zipFile) {
      // Batch processing
      setIsBatchGenerating(true);
      setBatchProgress({ current: 0, total: 0, message: 'Uploading ZIP...' });

      try {
        const formData = new FormData();
        formData.append('zip_file', zipFile);

        // Step 1: Upload ZIP and trigger batch report generation
        const response = await axios.post(
          `${process.env.REACT_APP_API_URL}/api/projects/${selectedProjectForReport.id}/upload_zip`,
          formData,
          {
            headers: { 'Content-Type': 'multipart/form-data' },
          }
        );

        setBatchProgress({ current: 1, total: 1, message: 'Batch report generation complete. Downloading ZIP...' });

        // Step 2: Download the resulting ZIP
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

        alert('Batch reports downloaded successfully!');
        handleCloseReportUploadDialog();
      } catch (error) {
        console.error('Batch report error:', error.response?.data || error.message);
        alert('Batch report generation failed.');
      } finally {
        setIsBatchGenerating(false);
        setBatchProgress({ current: 0, total: 0, message: '' });
      }
      return;
    }

    // Single file processing (existing logic)
    setIsGeneratingReport(true);

    try {
      const formData = new FormData();
      formData.append('report_file', reportFile);

      await axios.post(
        `${process.env.REACT_APP_API_URL}/api/projects/${selectedProjectForReport.id}/upload_report`,
        formData,
        {
          headers: { 'Content-Type': 'multipart/form-data' },
        }
      );
      alert('Report generation initiated successfully!');

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
        alert('Report downloaded successfully!');
      } catch (downloadError) {
        console.error('Error downloading report:', downloadError.response?.data || downloadError.message);
        alert('Failed to download report after generation. Please check backend logs for details.');
      }

      handleCloseReportUploadDialog();
    } catch (uploadError) {
      console.error('Error uploading report:', uploadError.response?.data || uploadError.message);
      alert('Failed to upload report to server. Please try again.');
    } finally {
      setIsGeneratingReport(false);
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
                    <Button
                      variant="outlined"
                      startIcon={<DescriptionIcon />}
                      onClick={() => handleOpenReportUploadDialog(project)}
                    >
                      Generate Report
                    </Button>
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
              <CircularProgress size={24} />
              <Typography variant="body2" sx={{ ml: 2, display: 'inline' }}>
                {batchProgress.message}
              </Typography>
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
    </Container>
  );
}

export default Dashboard; 