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
  Card,
  CardContent,
  CardActions,
  Grid,
  InputAdornment,
  Menu,
  MenuItem,
  Tooltip,
  Avatar,
  Badge,
  Fade,
  Zoom,
  useTheme,
  useMediaQuery,
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
  Settings as SettingsIcon,
  Edit as EditIcon,
  Delete as DeleteIcon,
  Search as SearchIcon,
  FilterList as FilterIcon,
  Sort as SortIcon,
  Dashboard as DashboardIcon,
  Folder as FolderIcon,
  MoreVert as MoreVertIcon,
  Download as DownloadIcon,
  Visibility as VisibilityIcon,
  CalendarToday as CalendarIcon,
  Person as PersonIcon,
  TrendingUp as TrendingUpIcon,
  CheckCircle as CheckCircleIcon,
  Cancel as CancelIcon,
  Refresh as RefreshIcon,
} from '@mui/icons-material';
import axios from 'axios';

axios.defaults.withCredentials = true;

function Dashboard() {
  const navigate = useNavigate();
  const [projects, setProjects] = useState([]);
  const [openCreateProjectDialog, setOpenCreateProjectDialog] = useState(false);
  const [openEditProjectDialog, setOpenEditProjectDialog] = useState(false);
  const [openReportUploadDialog, setOpenReportUploadDialog] = useState(false);
  const [newProject, setNewProject] = useState({ name: '', description: '', file: null });
  const [editingProject, setEditingProject] = useState({ name: '', description: '', file: null });
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

  // Enhanced functionality states
  const [searchTerm, setSearchTerm] = useState('');
  const [sortBy, setSortBy] = useState('created_at');
  const [sortOrder, setSortOrder] = useState('desc');
  const [filterStatus, setFilterStatus] = useState('all');
  const [anchorEl, setAnchorEl] = useState(null);
  const [selectedProjectForMenu, setSelectedProjectForMenu] = useState(null);
  const [viewMode, setViewMode] = useState('table'); // 'grid' or 'table'
  const [isLoading, setIsLoading] = useState(false);
  const [stats, setStats] = useState({
    total: 0,
    recent: 0,
    withErrors: 0,
    successful: 0
  });

  const theme = useTheme();
  const isMobile = useMediaQuery(theme.breakpoints.down('md'));

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
    setIsLoading(true);
    try {
      const response = await axios.get(`${process.env.REACT_APP_API_URL}/api/projects`);
      setProjects(response.data.projects);
      
      // Calculate stats
      const total = response.data.projects.length;
      const recent = response.data.projects.filter(p => {
        const daysSince = (new Date() - new Date(p.created_at)) / (1000 * 60 * 60 * 24);
        return daysSince <= 7;
      }).length;
      
      setStats({
        total,
        recent,
        withErrors: 0, // Will be updated when we check for errors
        successful: total
      });
    } catch (error) {
      console.error('Error loading projects:', error);
    } finally {
      setIsLoading(false);
    }
  };

  // Filter and sort projects
  const filteredAndSortedProjects = projects
    .filter(project => {
      const matchesSearch = project.name.toLowerCase().includes(searchTerm.toLowerCase()) ||
                           project.description.toLowerCase().includes(searchTerm.toLowerCase());
      
      if (filterStatus === 'all') return matchesSearch;
      if (filterStatus === 'recent') {
        const daysSince = (new Date() - new Date(project.created_at)) / (1000 * 60 * 60 * 24);
        return matchesSearch && daysSince <= 7;
      }
      return matchesSearch;
    })
    .sort((a, b) => {
      let aValue, bValue;
      
      switch (sortBy) {
        case 'name':
          aValue = a.name.toLowerCase();
          bValue = b.name.toLowerCase();
          break;
        case 'created_at':
          aValue = new Date(a.created_at);
          bValue = new Date(b.created_at);
          break;
        default:
          aValue = a[sortBy];
          bValue = b[sortBy];
      }
      
      if (sortOrder === 'asc') {
        return aValue > bValue ? 1 : -1;
      } else {
        return aValue < bValue ? 1 : -1;
      }
    });

  // Handle menu actions
  const handleMenuOpen = (event, project) => {
    setAnchorEl(event.currentTarget);
    setSelectedProjectForMenu(project);
  };

  const handleMenuClose = () => {
    setAnchorEl(null);
    setSelectedProjectForMenu(null);
  };

  const handleMenuAction = (action) => {
    if (!selectedProjectForMenu) return;
    
    switch (action) {
      case 'edit':
        handleEditProject(selectedProjectForMenu);
        break;
      case 'generate':
        handleOpenReportUploadDialog(selectedProjectForMenu);
        break;
      case 'errors':
        handleShowChartErrors(selectedProjectForMenu);
        break;
      case 'delete':
        handleDeleteProject(selectedProjectForMenu.id);
        break;
    }
    handleMenuClose();
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
      showCustomAlert('Success!', 'Project created successfully.', 'success');
    } catch (error) {
      console.error('Error creating project:', error.response?.data || error.message);
      showCustomAlert('Error!', error.response?.data?.error || 'Failed to create project.', 'error');
    }
  };

  const handleEditProject = (project) => {
    setEditingProject({
      id: project.id,
      name: project.name,
      description: project.description,
      file: null
    });
    setOpenEditProjectDialog(true);
  };

  const handleUpdateProject = async () => {
    try {
      const formData = new FormData();
      formData.append('name', editingProject.name);
      formData.append('description', editingProject.description);
      if (editingProject.file) {
        formData.append('file', editingProject.file);
      }

      await axios.put(`${process.env.REACT_APP_API_URL}/api/projects/${editingProject.id}`, formData, {
        headers: {
          'Content-Type': 'multipart/form-data',
        },
      });

      setOpenEditProjectDialog(false);
      setEditingProject({ name: '', description: '', file: null });
      loadProjects();
      showCustomAlert('Success!', 'Project updated successfully.', 'success');
    } catch (error) {
      console.error('Error updating project:', error.response?.data || error.message);
      showCustomAlert('Error!', error.response?.data?.error || 'Failed to update project.', 'error');
    }
  };

  const handleDeleteProject = async (projectId) => {
    if (!window.confirm('Are you sure you want to delete this project? This action cannot be undone.')) {
      return;
    }

    try {
      await axios.delete(`${process.env.REACT_APP_API_URL}/api/projects/${projectId}`);
      loadProjects();
      showCustomAlert('Success!', 'Project deleted successfully.', 'success');
    } catch (error) {
      console.error('Error deleting project:', error.response?.data || error.message);
      showCustomAlert('Error!', error.response?.data?.error || 'Failed to delete project.', 'error');
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
          message: 'Preparing download...', 
          percentage: 92 
        });
        
        // Small delay to ensure ZIP file is fully created
        await new Promise(resolve => setTimeout(resolve, 1000));
        
        setBatchProgress({ 
          current: processed_files, 
          total: total_files, 
          message: 'Downloading generated reports...', 
          percentage: 95 
        });
        
        try {
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
        } catch (downloadError) {
          console.error('Error downloading batch reports:', downloadError);
          setBatchProgress({ 
            current: processed_files, 
            total: total_files, 
            message: 'Reports generated but download failed', 
            percentage: 90 
          });
          
          setTimeout(() => {
            showCustomAlert(
              'Reports Generated! ⚠️',
              `${processed_files} reports were generated successfully, but the download failed. Please try again or contact support.`,
              'warning'
            );
            handleCloseReportUploadDialog();
          }, 1000);
        }
        
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
    <Box sx={{ 
      minHeight: '100vh',
      background: 'linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%)',
      py: 4,
      px: 2
    }}>
      <Container maxWidth="xl">
      {/* Header Section */}
      <Box sx={{ 
        background: 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)',
          borderRadius: 6,
        p: 4,
        mb: 4,
        color: 'white',
        position: 'relative',
          overflow: 'hidden',
          boxShadow: '0 20px 40px rgba(102, 126, 234, 0.3)',
          '&::before': {
            content: '""',
            position: 'absolute',
            top: 0,
            left: 0,
            right: 0,
            bottom: 0,
            background: 'radial-gradient(circle at 20% 80%, rgba(255, 255, 255, 0.1) 0%, transparent 50%), radial-gradient(circle at 80% 20%, rgba(255, 255, 255, 0.1) 0%, transparent 50%)',
            zIndex: 1,
          }
      }}>
          <Box sx={{ position: 'absolute', top: -20, right: -20, opacity: 0.1 }}>
            <DashboardIcon sx={{ fontSize: 160 }} />
        </Box>
          <Box sx={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', position: 'relative', zIndex: 2 }}>
          <Box>
              <Typography variant="h3" component="h1" sx={{ fontWeight: 800, mb: 2, letterSpacing: '-0.02em' }}>
              Project Dashboard
            </Typography>
              <Typography variant="h5" sx={{ opacity: 0.95, fontWeight: 400 }}>
              Welcome back, {user?.full_name} 👋
            </Typography>
          </Box>
            <Box sx={{ display: 'flex', alignItems: 'center', gap: 3 }}>
              <Avatar sx={{ 
                bgcolor: 'rgba(255,255,255,0.2)', 
                width: 56, 
                height: 56,
                border: '2px solid rgba(255,255,255,0.3)',
                backdropFilter: 'blur(10px)',
                boxShadow: '0 8px 32px rgba(0, 0, 0, 0.1)'
              }}>
                <PersonIcon sx={{ fontSize: 28 }} />
            </Avatar>
            <Tooltip title="Logout">
                <IconButton 
                  onClick={handleLogout} 
                  sx={{ 
                    color: 'white',
                    bgcolor: 'rgba(255,255,255,0.1)',
                    border: '1px solid rgba(255,255,255,0.3)',
                    borderRadius: 3,
                    backdropFilter: 'blur(10px)',
                    width: 48,
                    height: 48,
                    '&:hover': {
                      bgcolor: 'rgba(255,255,255,0.2)',
                      transform: 'scale(1.05)',
                      boxShadow: '0 8px 25px rgba(0, 0, 0, 0.2)',
                    },
                    transition: 'all 0.3s ease'
                  }}
                >
                <LogoutIcon />
              </IconButton>
            </Tooltip>
          </Box>
        </Box>
      </Box>

      {/* Stats Cards */}
      {/* Stats Cards */}
      <Grid container spacing={4} sx={{ mb: 5 }}>
        <Grid item xs={12} sm={6} md={3}>
          <Fade in timeout={600}>
          <Card sx={{ 
            background: 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)',
            color: 'white',
            position: 'relative',
              overflow: 'hidden',
              borderRadius: 6,
              boxShadow: '0 12px 24px rgba(102, 126, 234, 0.25)',
              transition: 'all 0.3s ease',
              '&:hover': {
                transform: 'translateY(-8px)',
                boxShadow: '0 20px 40px rgba(102, 126, 234, 0.4)',
              }
          }}>
              <CardContent sx={{ p: 3 }}>
              <Box sx={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <Box>
                    <Typography variant="h3" sx={{ fontWeight: 800, mb: 1 }}>
                    {stats.total}
                  </Typography>
                    <Typography variant="body1" sx={{ opacity: 0.95, fontWeight: 500 }}>
                    Total Projects
                  </Typography>
                </Box>
                  <Box sx={{ 
                    bgcolor: 'rgba(255,255,255,0.2)', 
                    borderRadius: '50%', 
                    p: 2,
                    backdropFilter: 'blur(10px)',
                    boxShadow: '0 8px 32px rgba(0, 0, 0, 0.1)'
                  }}>
                    <FolderIcon sx={{ fontSize: 32, opacity: 0.9 }} />
                  </Box>
              </Box>
            </CardContent>
          </Card>
          </Fade>
        </Grid>
        <Grid item xs={12} sm={6} md={3}>
          <Fade in timeout={800}>
          <Card sx={{ 
            background: 'linear-gradient(135deg, #f093fb 0%, #f5576c 100%)',
              color: 'white',
              borderRadius: 6,
              boxShadow: '0 12px 24px rgba(240, 147, 251, 0.25)',
              transition: 'all 0.3s ease',
              '&:hover': {
                transform: 'translateY(-8px)',
                boxShadow: '0 20px 40px rgba(240, 147, 251, 0.4)',
              }
          }}>
              <CardContent sx={{ p: 3 }}>
              <Box sx={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <Box>
                    <Typography variant="h3" sx={{ fontWeight: 800, mb: 1 }}>
                    {stats.recent}
                  </Typography>
                    <Typography variant="body1" sx={{ opacity: 0.95, fontWeight: 500 }}>
                    Recent (7 days)
                  </Typography>
                </Box>
                  <Box sx={{ 
                    bgcolor: 'rgba(255,255,255,0.2)', 
                    borderRadius: '50%', 
                    p: 2,
                    backdropFilter: 'blur(10px)',
                    boxShadow: '0 8px 32px rgba(0, 0, 0, 0.1)'
                  }}>
                    <TrendingUpIcon sx={{ fontSize: 32, opacity: 0.9 }} />
                  </Box>
              </Box>
            </CardContent>
          </Card>
          </Fade>
        </Grid>
        <Grid item xs={12} sm={6} md={3}>
          <Fade in timeout={1000}>
          <Card sx={{ 
            background: 'linear-gradient(135deg, #4facfe 0%, #00f2fe 100%)',
              color: 'white',
              borderRadius: 6,
              boxShadow: '0 12px 24px rgba(79, 172, 254, 0.25)',
              transition: 'all 0.3s ease',
              '&:hover': {
                transform: 'translateY(-8px)',
                boxShadow: '0 20px 40px rgba(79, 172, 254, 0.4)',
              }
          }}>
              <CardContent sx={{ p: 3 }}>
              <Box sx={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <Box>
                    <Typography variant="h3" sx={{ fontWeight: 800, mb: 1 }}>
                    {stats.successful}
                  </Typography>
                    <Typography variant="body1" sx={{ opacity: 0.95, fontWeight: 500 }}>
                    Successful
                  </Typography>
                </Box>
                  <Box sx={{ 
                    bgcolor: 'rgba(255,255,255,0.2)', 
                    borderRadius: '50%', 
                    p: 2,
                    backdropFilter: 'blur(10px)',
                    boxShadow: '0 8px 32px rgba(0, 0, 0, 0.1)'
                  }}>
                    <CheckCircleIcon sx={{ fontSize: 32, opacity: 0.9 }} />
                  </Box>
              </Box>
            </CardContent>
          </Card>
          </Fade>
        </Grid>
        <Grid item xs={12} sm={6} md={3}>
          <Fade in timeout={1200}>
          <Card sx={{ 
            background: 'linear-gradient(135deg, #fa709a 0%, #fee140 100%)',
              color: 'white',
              borderRadius: 6,
              boxShadow: '0 12px 24px rgba(250, 112, 154, 0.25)',
              transition: 'all 0.3s ease',
              '&:hover': {
                transform: 'translateY(-8px)',
                boxShadow: '0 20px 40px rgba(250, 112, 154, 0.4)',
              }
          }}>
              <CardContent sx={{ p: 3 }}>
              <Box sx={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <Box>
                    <Typography variant="h3" sx={{ fontWeight: 800, mb: 1 }}>
                    {stats.withErrors}
                  </Typography>
                    <Typography variant="body1" sx={{ opacity: 0.95, fontWeight: 500 }}>
                    With Issues
                  </Typography>
                </Box>
                  <Box sx={{ 
                    bgcolor: 'rgba(255,255,255,0.2)', 
                    borderRadius: '50%', 
                    p: 2,
                    backdropFilter: 'blur(10px)',
                    boxShadow: '0 8px 32px rgba(0, 0, 0, 0.1)'
                  }}>
                    <CancelIcon sx={{ fontSize: 32, opacity: 0.9 }} />
                  </Box>
              </Box>
            </CardContent>
          </Card>
          </Fade>
        </Grid>
      </Grid>

      {/* Controls Section */}
      <Paper sx={{ 
        p: 3, 
        mb: 3, 
        borderRadius: 4,
        background: 'white',
        border: '2px solid #000',
        boxShadow: '0 4px 12px rgba(0, 0, 0, 0.1)',
      }}>
        <Box sx={{ display: 'flex', flexDirection: isMobile ? 'column' : 'row', gap: 2, alignItems: 'center', justifyContent: 'space-between' }}>
          <Box sx={{ display: 'flex', gap: 2, flexWrap: 'wrap', flex: 1 }}>
            {/* Search */}
            <TextField
              size="small"
              placeholder="Search projects..."
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              InputProps={{
                startAdornment: (
                  <InputAdornment position="start">
                    <SearchIcon sx={{ color: '#667eea' }} />
                  </InputAdornment>
                ),
                sx: {
                  borderRadius: 2,
                  '&:hover .MuiOutlinedInput-notchedOutline': {
                    borderColor: '#667eea',
                  },
                  '&.Mui-focused .MuiOutlinedInput-notchedOutline': {
                    borderColor: '#667eea',
                  },
                }
              }}
              sx={{ minWidth: 250 }}
            />
            
            {/* Filter */}
            <TextField
              select
              size="small"
              value={filterStatus}
              onChange={(e) => setFilterStatus(e.target.value)}
              InputProps={{
                startAdornment: (
                  <InputAdornment position="start">
                    <FilterIcon sx={{ color: '#667eea' }} />
                  </InputAdornment>
                ),
                sx: {
                  borderRadius: 2,
                  '&:hover .MuiOutlinedInput-notchedOutline': {
                    borderColor: '#667eea',
                  },
                  '&.Mui-focused .MuiOutlinedInput-notchedOutline': {
                    borderColor: '#667eea',
                  },
                }
              }}
              sx={{ minWidth: 120 }}
            >
              <MenuItem value="all">All Projects</MenuItem>
              <MenuItem value="recent">Recent (7 days)</MenuItem>
            </TextField>

            {/* Sort */}
            <TextField
              select
              size="small"
              value={sortBy}
              onChange={(e) => setSortBy(e.target.value)}
              InputProps={{
                startAdornment: (
                  <InputAdornment position="start">
                    <SortIcon sx={{ color: '#667eea' }} />
                  </InputAdornment>
                ),
                sx: {
                  borderRadius: 2,
                  '&:hover .MuiOutlinedInput-notchedOutline': {
                    borderColor: '#667eea',
                  },
                  '&.Mui-focused .MuiOutlinedInput-notchedOutline': {
                    borderColor: '#667eea',
                  },
                }
              }}
              sx={{ minWidth: 140 }}
            >
              <MenuItem value="created_at">Date Created</MenuItem>
              <MenuItem value="name">Project Name</MenuItem>
            </TextField>

            {/* Sort Order */}
            <IconButton 
              onClick={() => setSortOrder(sortOrder === 'asc' ? 'desc' : 'asc')}
              sx={{ 
                border: 2, 
                borderColor: '#667eea',
                borderRadius: 2,
                color: '#667eea',
                '&:hover': {
                  bgcolor: 'rgba(102, 126, 234, 0.1)',
                  transform: 'scale(1.05)',
                },
                transition: 'all 0.2s ease'
              }}
            >
              <SortIcon sx={{ transform: sortOrder === 'desc' ? 'scaleY(-1)' : 'none' }} />
            </IconButton>
          </Box>

          <Box sx={{ display: 'flex', gap: 1 }}>
            {/* View Mode Toggle */}
            <Button
              variant={viewMode === 'grid' ? 'contained' : 'outlined'}
              size="small"
              onClick={() => setViewMode('grid')}
              startIcon={<DashboardIcon />}
              sx={{
                borderRadius: 2,
                px: 2,
                border: '2px solid #000',
                ...(viewMode === 'grid' && {
                  background: '#000',
                  color: 'white',
                  '&:hover': {
                    background: '#333',
                  }
                })
              }}
            >
              Grid
            </Button>
            <Button
              variant={viewMode === 'table' ? 'contained' : 'outlined'}
              size="small"
              onClick={() => setViewMode('table')}
              startIcon={<DescriptionIcon />}
              sx={{
                borderRadius: 2,
                px: 2,
                border: '2px solid #000',
                ...(viewMode === 'table' && {
                  background: '#000',
                  color: 'white',
                  '&:hover': {
                    background: '#333',
                  }
                })
              }}
            >
              Table
            </Button>
            
            {/* Create Project Button */}
            <Button
              variant="contained"
              size="medium"
              startIcon={<AddIcon />}
              onClick={() => setOpenCreateProjectDialog(true)}
              sx={{ 
                background: 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)',
                borderRadius: 2,
                px: 3,
                py: 1,
                fontWeight: 600,
                fontSize: 14,
                border: '2px solid #000',
                transition: 'all 0.2s ease',
                '&:hover': {
                  background: 'linear-gradient(135deg, #5a6fd8 0%, #6a4190 100%)',
                  transform: 'translateY(-1px)',
                }
              }}
            >
              New Project
            </Button>
          </Box>
        </Box>
      </Paper>

      {/* Projects Display */}
      {isLoading ? (
        <Box sx={{ display: 'flex', justifyContent: 'center', py: 8 }}>
          <CircularProgress size={60} />
        </Box>
      ) : filteredAndSortedProjects.length === 0 ? (
        <Paper sx={{ p: 8, textAlign: 'center', borderRadius: 4 }}>
          <FolderIcon sx={{ fontSize: 80, color: 'text.secondary', mb: 2 }} />
          <Typography variant="h5" color="text.secondary" sx={{ mb: 1 }}>
            {searchTerm ? 'No projects found' : 'No projects yet'}
          </Typography>
          <Typography variant="body1" color="text.secondary" sx={{ mb: 3 }}>
            {searchTerm ? 'Try adjusting your search terms' : 'Create your first project to get started'}
          </Typography>
          {!searchTerm && (
            <Button
              variant="contained"
              startIcon={<AddIcon />}
              onClick={() => setOpenCreateProjectDialog(true)}
              sx={{ 
                background: 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)',
                '&:hover': {
                  background: 'linear-gradient(135deg, #5a6fd8 0%, #6a4190 100%)',
                }
              }}
            >
              Create First Project
            </Button>
          )}
        </Paper>
      ) : (
        <TableContainer component={Paper} sx={{ 
          borderRadius: 4, 
          boxShadow: '0 4px 12px rgba(0, 0, 0, 0.1)',
          overflow: 'hidden',
          background: 'white',
          border: '2px solid #000'
        }}>
          <Table>
            <TableHead>
              <TableRow sx={{ 
                background: 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)',
                '& .MuiTableCell-head': {
                  color: 'white',
                  fontWeight: 700,
                  fontSize: '0.9rem',
                  padding: '12px 16px',
                  borderBottom: 'none',
                },
                '& .MuiTableCell-head:first-of-type': {
                  borderTopLeftRadius: 16,
                },
                '& .MuiTableCell-head:last-of-type': {
                  borderTopRightRadius: 16,
                }
              }}>
                <TableCell>Project Name</TableCell>
                <TableCell>Description</TableCell>
                <TableCell>Created On</TableCell>
                <TableCell align="center">Status</TableCell>
                <TableCell align="right">Actions</TableCell>
              </TableRow>
            </TableHead>
            <TableBody>
              {filteredAndSortedProjects.map((project, index) => (
                <TableRow 
                  key={project.id} 
                  hover
                  sx={{ 
                    transition: 'all 0.2s ease',
                    '&:hover': {
                      backgroundColor: '#f5f5f5',
                      transform: 'scale(1.002)',
                    },
                    '&:nth-of-type(even)': {
                      backgroundColor: 'rgba(0, 0, 0, 0.02)',
                    },
                    '& .MuiTableCell-body': {
                      padding: '12px 16px',
                      borderBottom: '1px solid rgba(0, 0, 0, 0.1)',
                    }
                  }}
                >
                  <TableCell component="th" scope="row" sx={{ fontWeight: 600, fontSize: '1rem' }}>
                    <Box sx={{ display: 'flex', alignItems: 'center', gap: 1.5 }}>
                      <Box sx={{ 
                        bgcolor: 'rgba(102, 126, 234, 0.1)', 
                        borderRadius: '50%', 
                        p: 0.8,
                        display: 'flex',
                        alignItems: 'center',
                        justifyContent: 'center',
                        border: '1px solid #000'
                      }}>
                        <FolderIcon sx={{ color: '#667eea', fontSize: 16 }} />
                      </Box>
                      {project.name}
                    </Box>
                  </TableCell>
                  <TableCell sx={{ maxWidth: 300 }}>
                    <Typography variant="body2" color="text.secondary" sx={{ 
                      overflow: 'hidden',
                      textOverflow: 'ellipsis',
                      display: '-webkit-box',
                      WebkitLineClamp: 2,
                      WebkitBoxOrient: 'vertical',
                      lineHeight: 1.3,
                    }}>
                      {project.description || 'No description provided'}
                    </Typography>
                  </TableCell>
                  <TableCell>
                    <Box sx={{ display: 'flex', alignItems: 'center', gap: 1.5 }}>
                      <Box sx={{ 
                        bgcolor: 'rgba(102, 126, 234, 0.1)', 
                        borderRadius: '50%', 
                        p: 0.8,
                        display: 'flex',
                        alignItems: 'center',
                        justifyContent: 'center',
                        border: '1px solid #000'
                      }}>
                        <CalendarIcon sx={{ color: '#667eea', fontSize: 14 }} />
                      </Box>
                      <Typography variant="body2" fontWeight={500}>
                        {new Date(project.created_at).toLocaleDateString()}
                      </Typography>
                    </Box>
                  </TableCell>
                  <TableCell align="center">
                    <Chip 
                      label="Active" 
                      color="success" 
                      size="small"
                      icon={<CheckCircleIcon />}
                      sx={{ 
                        fontWeight: 600,
                        borderRadius: 2,
                        border: '1px solid #000',
                        '& .MuiChip-icon': {
                          color: 'inherit'
                        }
                      }}
                    />
                  </TableCell>
                  <TableCell align="right">
                    <Box sx={{ display: 'flex', gap: 0.5, justifyContent: 'flex-end', flexWrap: 'wrap' }}>
                      <Tooltip title="Generate Report">
                        <IconButton
                          size="small"
                          onClick={() => handleOpenReportUploadDialog(project)}
                          sx={{ 
                            backgroundColor: 'rgba(102, 126, 234, 0.1)',
                            borderRadius: 2,
                            border: '1px solid #000',
                            '&:hover': {
                              backgroundColor: 'rgba(102, 126, 234, 0.2)',
                              transform: 'scale(1.05)',
                            },
                            transition: 'all 0.2s ease'
                          }}
                        >
                          <DescriptionIcon sx={{ fontSize: 18, color: '#667eea' }} />
                        </IconButton>
                      </Tooltip>
                      <Tooltip title="View Errors">
                        <IconButton
                          size="small"
                          onClick={() => handleShowChartErrors(project)}
                          sx={{ 
                            backgroundColor: 'rgba(255, 152, 0, 0.1)',
                            borderRadius: 2,
                            border: '1px solid #000',
                            '&:hover': {
                              backgroundColor: 'rgba(255, 152, 0, 0.2)',
                              transform: 'scale(1.05)',
                            },
                            transition: 'all 0.2s ease'
                          }}
                        >
                          <BugReportIcon sx={{ fontSize: 18, color: '#ff9800' }} />
                        </IconButton>
                      </Tooltip>
                      <Tooltip title="Edit Project">
                        <IconButton
                          size="small"
                          onClick={() => handleEditProject(project)}
                          sx={{ 
                            backgroundColor: 'rgba(3, 169, 244, 0.1)',
                            borderRadius: 2,
                            border: '1px solid #000',
                            '&:hover': {
                              backgroundColor: 'rgba(3, 169, 244, 0.2)',
                              transform: 'scale(1.05)',
                            },
                            transition: 'all 0.2s ease'
                          }}
                        >
                          <EditIcon sx={{ fontSize: 18, color: '#03a9f4' }} />
                        </IconButton>
                      </Tooltip>
                      <Tooltip title="More Options">
                        <IconButton
                          size="small"
                          onClick={(e) => handleMenuOpen(e, project)}
                          sx={{ 
                            backgroundColor: 'rgba(0, 0, 0, 0.05)',
                            borderRadius: 2,
                            border: '1px solid #000',
                            '&:hover': {
                              backgroundColor: 'rgba(0, 0, 0, 0.1)',
                              transform: 'scale(1.05)',
                            },
                            transition: 'all 0.2s ease'
                          }}
                        >
                          <MoreVertIcon sx={{ fontSize: 18, color: '#000' }} />
                        </IconButton>
                      </Tooltip>
                    </Box>
                  </TableCell>
                </TableRow>
              ))}
            </TableBody>
          </Table>
        </TableContainer>
      )}

      {/* Grid View (Hidden by default, can be toggled) */}
      {viewMode === 'grid' && (
        <Grid container spacing={3}>
          {filteredAndSortedProjects.map((project) => (
            <Grid item xs={12} sm={6} md={4} key={project.id}>
              <Fade in timeout={300}>
                <Card sx={{ 
                  height: '100%', 
                  display: 'flex', 
                  flexDirection: 'column',
                  transition: 'transform 0.2s, box-shadow 0.2s',
                  borderRadius: 4,
                  '&:hover': {
                    transform: 'translateY(-4px)',
                    boxShadow: theme.shadows[8],
                  }
                }}>
                  <CardContent sx={{ flexGrow: 1 }}>
                    <Box sx={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', mb: 2 }}>
                      <Typography variant="h6" component="h2" sx={{ fontWeight: 'bold', flex: 1 }}>
                        {project.name}
                      </Typography>
                      <IconButton
                        size="small"
                        onClick={(e) => handleMenuOpen(e, project)}
                      >
                        <MoreVertIcon />
                      </IconButton>
                    </Box>
                    
                    <Typography variant="body2" color="text.secondary" sx={{ mb: 2, minHeight: 40 }}>
                      {project.description || 'No description provided'}
                    </Typography>
                    
                    <Box sx={{ display: 'flex', alignItems: 'center', gap: 1, mb: 2 }}>
                      <CalendarIcon fontSize="small" color="action" />
                      <Typography variant="caption" color="text.secondary">
                        Created {new Date(project.created_at).toLocaleDateString()}
                      </Typography>
                    </Box>
                  </CardContent>
                  
                  <CardActions sx={{ p: 2, pt: 0 }}>
                    <Button
                      size="small"
                      variant="outlined"
                      startIcon={<DescriptionIcon />}
                      onClick={() => handleOpenReportUploadDialog(project)}
                      fullWidth
                    >
                      Generate Report
                    </Button>
                  </CardActions>
                </Card>
              </Fade>
            </Grid>
          ))}
        </Grid>
      )}

      {/* Project Actions Menu */}
      <Menu
        anchorEl={anchorEl}
        open={Boolean(anchorEl)}
        onClose={handleMenuClose}
        PaperProps={{
          sx: {
            borderRadius: 3,
            minWidth: 180,
          }
        }}
      >
        <MenuItem onClick={() => handleMenuAction('edit')}>
          <ListItemIcon>
            <EditIcon fontSize="small" />
          </ListItemIcon>
          Edit Project
        </MenuItem>
        <MenuItem onClick={() => handleMenuAction('generate')}>
          <ListItemIcon>
            <DescriptionIcon fontSize="small" />
          </ListItemIcon>
          Generate Report
        </MenuItem>
        <MenuItem onClick={() => handleMenuAction('errors')}>
          <ListItemIcon>
            <BugReportIcon fontSize="small" />
          </ListItemIcon>
          View Errors
        </MenuItem>
        <Divider />
        <MenuItem onClick={() => handleMenuAction('delete')} sx={{ color: 'error.main' }}>
          <ListItemIcon>
            <DeleteIcon fontSize="small" color="error" />
          </ListItemIcon>
          Delete Project
        </MenuItem>
      </Menu>

      {/* Create Project Dialog */}
      <Dialog 
        open={openCreateProjectDialog} 
        onClose={() => setOpenCreateProjectDialog(false)}
        PaperProps={{
          sx: { borderRadius: 4, minWidth: 400 }
        }}
      >
        <DialogTitle sx={{ 
          background: 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)',
          color: 'white',
          display: 'flex',
          alignItems: 'center',
          gap: 1
        }}>
          <AddIcon />
          Create New Project
        </DialogTitle>
        <DialogContent sx={{ pt: 3 }}>
          <TextField
            autoFocus
            margin="dense"
            label="Project Name"
            fullWidth
            value={newProject.name}
            onChange={(e) => setNewProject({ ...newProject, name: e.target.value })}
            sx={{ mb: 2 }}
          />
          <TextField
            margin="dense"
            label="Description"
            fullWidth
            multiline
            rows={4}
            value={newProject.description}
            onChange={(e) => setNewProject({ ...newProject, description: e.target.value })}
            sx={{ mb: 2 }}
          />
          <Box sx={{ mt: 2 }}>
            <Typography variant="body2" color="text.secondary" sx={{ mb: 1 }}>
              Word Template (Optional)
            </Typography>
            <input
              type="file"
              accept=".doc,.docx"
              onChange={(e) => setNewProject({ ...newProject, file: e.target.files[0] })}
              style={{ 
                width: '100%',
                padding: '8px',
                border: '1px dashed #ccc',
                borderRadius: '4px',
                backgroundColor: '#fafafa'
              }}
            />
          </Box>
        </DialogContent>
        <DialogActions sx={{ p: 3, pt: 0 }}>
          <Button onClick={() => setOpenCreateProjectDialog(false)}>
            Cancel
          </Button>
          <Button 
            onClick={handleCreateProject} 
            variant="contained"
            sx={{ 
              background: 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)',
              '&:hover': {
                background: 'linear-gradient(135deg, #5a6fd8 0%, #6a4190 100%)',
              }
            }}
          >
            Create Project
          </Button>
        </DialogActions>
      </Dialog>

      {/* Edit Project Dialog */}
      <Dialog 
        open={openEditProjectDialog} 
        onClose={() => setOpenEditProjectDialog(false)}
        PaperProps={{
          sx: { borderRadius: 4, minWidth: 400 }
        }}
      >
        <DialogTitle sx={{ 
          background: 'linear-gradient(135deg, #4facfe 0%, #00f2fe 100%)',
          color: 'white',
          display: 'flex',
          alignItems: 'center',
          gap: 1
        }}>
          <EditIcon />
          Edit Project
        </DialogTitle>
        <DialogContent sx={{ pt: 3 }}>
          <TextField
            autoFocus
            margin="dense"
            label="Project Name"
            fullWidth
            value={editingProject.name}
            onChange={(e) => setEditingProject({ ...editingProject, name: e.target.value })}
            sx={{ mb: 2 }}
          />
          <TextField
            margin="dense"
            label="Description"
            fullWidth
            multiline
            rows={4}
            value={editingProject.description}
            onChange={(e) => setEditingProject({ ...editingProject, description: e.target.value })}
            sx={{ mb: 2 }}
          />
          <Box sx={{ mt: 2 }}>
            <Typography variant="body2" color="text.secondary" sx={{ mb: 1 }}>
              Upload new Word template (optional - leave empty to keep current template)
            </Typography>
            <input
              type="file"
              accept=".doc,.docx"
              onChange={(e) => setEditingProject({ ...editingProject, file: e.target.files[0] })}
              style={{ 
                width: '100%',
                padding: '8px',
                border: '1px dashed #ccc',
                borderRadius: '4px',
                backgroundColor: '#fafafa'
              }}
            />
          </Box>
        </DialogContent>
        <DialogActions sx={{ p: 3, pt: 0 }}>
          <Button onClick={() => setOpenEditProjectDialog(false)}>
            Cancel
          </Button>
          <Button 
            onClick={handleUpdateProject} 
            variant="contained"
            sx={{ 
              background: 'linear-gradient(135deg, #4facfe 0%, #00f2fe 100%)',
              '&:hover': {
                background: 'linear-gradient(135deg, #3e9bed 0%, #00d8e0 100%)',
              }
            }}
          >
            Save Changes
          </Button>
        </DialogActions>
      </Dialog>

      

      {/* Report Upload Dialog */}
      <Dialog 
        open={openReportUploadDialog} 
        onClose={handleCloseReportUploadDialog}
        PaperProps={{
          sx: { borderRadius: 4, minWidth: 500 }
        }}
      >
        <DialogTitle sx={{ 
          background: 'linear-gradient(135deg, #f093fb 0%, #f5576c 100%)',
          color: 'white',
          display: 'flex',
          alignItems: 'center',
          gap: 1
        }}>
          <DescriptionIcon />
          Generate Report - {selectedProjectForReport?.name}
        </DialogTitle>
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
        maxWidth="md"
        fullWidth
        PaperProps={{
          sx: { borderRadius: 4 }
        }}
      >
        <DialogTitle sx={{ 
          background: 'linear-gradient(135deg, #fa709a 0%, #fee140 100%)',
          color: 'white',
          display: 'flex',
          alignItems: 'center',
          gap: 1
        }}>
          <ErrorIcon />
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
        TransitionComponent={Zoom}
      >
        <Alert
          onClose={handleCloseCustomAlert}
          severity={customAlert.severity}
          variant="filled"
          elevation={6}
          sx={{
            width: '100%',
            maxWidth: '500px',
              borderRadius: 4,
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
      </Box>
  );
}

export default Dashboard; 