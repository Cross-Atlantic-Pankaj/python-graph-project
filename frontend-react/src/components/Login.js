import React, { useState } from 'react';
import { useNavigate, Link } from 'react-router-dom';
import {
  Container,
  Paper,
  TextField,
  Button,
  Typography,
  Box,
  Fade,
  Zoom,
  InputAdornment,
  IconButton,
  Divider,
  CircularProgress,
} from '@mui/material';
import Alert from './Alert';
import api from '../api';
import LockOutlinedIcon from '@mui/icons-material/LockOutlined';
import VisibilityIcon from '@mui/icons-material/Visibility';
import VisibilityOffIcon from '@mui/icons-material/VisibilityOff';
import EmailIcon from '@mui/icons-material/Email';
import PersonIcon from '@mui/icons-material/Person';

// Ensure cookies are included for session-based auth
api.defaults.withCredentials = true;

function Login() {
  const navigate = useNavigate();
  const [formData, setFormData] = useState({
    username: '',
    password: '',
  });
  const [error, setError] = useState('');
  const [showPassword, setShowPassword] = useState(false);
  const [isLoading, setIsLoading] = useState(false);
  const [alert, setAlert] = useState({
    open: false,
    message: '',
    severity: 'error',
    title: ''
  });

  const handleChange = (e) => {
    setFormData({
      ...formData,
      [e.target.name]: e.target.value,
    });
  };

  const handleSubmit = async (e) => {
    e.preventDefault();
    setIsLoading(true);
    setError('');
    
    try {
      const response = await api.post('/login', formData);
      if (response.data.message === 'Login successful') {
        navigate('/dashboard');
      }
    } catch (err) {
      const errorMessage = err.response?.data?.error || 'Login failed. Please check your credentials.';
      setError(errorMessage);
      setAlert({
        open: true,
        title: 'Login Failed',
        message: errorMessage,
        severity: 'error'
      });
    } finally {
      setIsLoading(false);
    }
  };

  const handleCloseAlert = () => {
    setAlert(prev => ({ ...prev, open: false }));
  };

    return (
    <>
      <Box
        sx={{
          height: '100vh',
          width: '100vw',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          background: '#f8f9fa',
          p: 2,
          overflow: 'hidden',
          position: 'fixed',
          top: 0,
          left: 0,
          right: 0,
          bottom: 0,
        }}
      >
        <Fade in timeout={800}>
          <Paper
            elevation={8}
            sx={{
              p: { xs: 2.5, sm: 4 },
              width: '100%',
              maxWidth: 450,
              height: 'auto',
              maxHeight: '85vh',
              borderRadius: 4,
              boxShadow: '0 8px 32px rgba(0, 0, 0, 0.1)',
              display: 'flex',
              flexDirection: 'column',
              alignItems: 'center',
              background: 'white',
              border: '1px solid #e0e0e0',
              overflow: 'hidden',
              position: 'relative',
            }}
          >
              <Zoom in timeout={1000}>
                <Box sx={{ mb: 2, display: 'flex', flexDirection: 'column', alignItems: 'center' }}>
                  <Box
                    sx={{
                      background: 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)',
                      borderRadius: '50%',
                      width: 80,
                      height: 80,
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'center',
                      mb: 2,
                      boxShadow: '0 4px 16px rgba(102, 126, 234, 0.2)',
                    }}
                  >
                    <LockOutlinedIcon sx={{ color: '#fff', fontSize: 40 }} />
                  </Box>
                  <Typography 
                    variant="h4" 
                    component="h1" 
                    fontWeight={700} 
                    gutterBottom 
                    align="center" 
                    sx={{
                      color: '#667eea',
                      mb: 1,
                    }}
                  >
                    Welcome Back
                  </Typography>
                  <Typography variant="body1" color="text.secondary" align="center" sx={{ opacity: 0.8 }}>
                    Sign in to access your dashboard
                  </Typography>
                </Box>
              </Zoom>

              <form onSubmit={handleSubmit} style={{ width: '100%', flex: 1, display: 'flex', flexDirection: 'column', justifyContent: 'center' }}>
                <TextField
                  fullWidth
                  label="Username"
                  name="username"
                  value={formData.username}
                  onChange={handleChange}
                  margin="normal"
                  required
                  autoFocus
                  InputProps={{
                    sx: { 
                      borderRadius: 2,
                      '&:hover .MuiOutlinedInput-notchedOutline': {
                        borderColor: '#667eea',
                      },
                      '&.Mui-focused .MuiOutlinedInput-notchedOutline': {
                        borderColor: '#667eea',
                      },
                    },
                    startAdornment: (
                      <InputAdornment position="start">
                        <PersonIcon sx={{ color: '#667eea', opacity: 0.7 }} />
                      </InputAdornment>
                    ),
                  }}
                  InputLabelProps={{
                    sx: {
                      '&.Mui-focused': {
                        color: '#667eea',
                      },
                    },
                  }}
                />
                <TextField
                  fullWidth
                  label="Password"
                  name="password"
                  type={showPassword ? 'text' : 'password'}
                  value={formData.password}
                  onChange={handleChange}
                  margin="normal"
                  required
                  InputProps={{
                    sx: { 
                      borderRadius: 2,
                      '&:hover .MuiOutlinedInput-notchedOutline': {
                        borderColor: '#667eea',
                      },
                      '&.Mui-focused .MuiOutlinedInput-notchedOutline': {
                        borderColor: '#667eea',
                      },
                    },
                    startAdornment: (
                      <InputAdornment position="start">
                        <LockOutlinedIcon sx={{ color: '#667eea', opacity: 0.7 }} />
                      </InputAdornment>
                    ),
                    endAdornment: (
                      <InputAdornment position="end">
                        <IconButton
                          onClick={() => setShowPassword(!showPassword)}
                          edge="end"
                          sx={{ color: '#667eea' }}
                        >
                          {showPassword ? <VisibilityOffIcon /> : <VisibilityIcon />}
                        </IconButton>
                      </InputAdornment>
                    ),
                  }}
                  InputLabelProps={{
                    sx: {
                      '&.Mui-focused': {
                        color: '#667eea',
                      },
                    },
                  }}
                />
                
                <Button
                  type="submit"
                  fullWidth
                  variant="contained"
                  size="large"
                  disabled={isLoading}
                  sx={{ 
                    mt: 3, 
                    mb: 2,
                    borderRadius: 2, 
                    fontWeight: 600, 
                    fontSize: 16, 
                    py: 1.5, 
                    background: 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)',
                    boxShadow: '0 4px 12px rgba(102, 126, 234, 0.2)',
                    transition: 'all 0.3s ease',
                    '&:hover': {
                      background: 'linear-gradient(135deg, #5a6fd8 0%, #6a4190 100%)',
                      boxShadow: '0 6px 20px rgba(102, 126, 234, 0.3)',
                      transform: 'translateY(-1px)',
                    },
                    '&:disabled': {
                      background: 'linear-gradient(135deg, #b0b0b0 0%, #909090 100%)',
                      boxShadow: 'none',
                      transform: 'none',
                    },
                  }}
                >
                  {isLoading ? (
                    <CircularProgress size={20} sx={{ color: '#fff' }} />
                  ) : (
                    'Sign In'
                  )}
                </Button>
              </form>

              <Divider sx={{ width: '100%', my: 2, opacity: 0.3 }} />

              <Box sx={{ mt: 2, textAlign: 'center', width: '100%' }}>
                <Typography variant="body2" color="text.secondary">
                  New to this platform?{' '}
                  <Link 
                    to="/register" 
                    style={{ 
                      textDecoration: 'none', 
                      color: '#667eea', 
                      fontWeight: 600,
                      transition: 'color 0.3s ease',
                    }}
                    onMouseEnter={(e) => e.target.style.color = '#5a6fd8'}
                    onMouseLeave={(e) => e.target.style.color = '#667eea'}
                  >
                    Create an account
                  </Link>
                </Typography>
              </Box>
            </Paper>
          </Fade>
        </Box>

        {/* Custom Alert Component */}
        <Alert
          open={alert.open}
          onClose={handleCloseAlert}
          title={alert.title}
          message={alert.message}
          severity={alert.severity}
        />
      </>
    );
}

export default Login; 