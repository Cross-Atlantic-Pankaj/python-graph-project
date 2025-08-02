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
  Stepper,
  Step,
  StepLabel,
  CircularProgress,
} from '@mui/material';
import Alert from './Alert';
import axios from 'axios';
import PersonAddAlt1Icon from '@mui/icons-material/PersonAddAlt1';
import VisibilityIcon from '@mui/icons-material/Visibility';
import VisibilityOffIcon from '@mui/icons-material/VisibilityOff';
import EmailIcon from '@mui/icons-material/Email';
import PersonIcon from '@mui/icons-material/Person';
import BadgeIcon from '@mui/icons-material/Badge';

function Register() {
  const navigate = useNavigate();
  const [formData, setFormData] = useState({
    full_name: '',
    username: '',
    email: '',
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
      const response = await axios.post(`${process.env.REACT_APP_API_URL}/api/register`, formData);
      if (response.data.message === 'Registration successful') {
        setAlert({
          open: true,
          title: 'Registration Successful!',
          message: 'Account created successfully. Please sign in.',
          severity: 'success'
        });
        setTimeout(() => {
          navigate('/login');
        }, 2000);
      }
    } catch (err) {
      const errorMessage = err.response?.data?.error || 'Registration failed. Please try again.';
      setError(errorMessage);
      setAlert({
        open: true,
        title: 'Registration Failed',
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

  const steps = ['Personal Info', 'Account Details', 'Security'];

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
            maxWidth: 500,
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
                    width: 70,
                    height: 70,
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'center',
                    mb: 2,
                    boxShadow: '0 4px 16px rgba(102, 126, 234, 0.2)',
                  }}
                >
                  <PersonAddAlt1Icon sx={{ color: '#fff', fontSize: 35 }} />
                </Box>
                <Typography 
                  variant="h5" 
                  component="h1" 
                  fontWeight={700} 
                  gutterBottom 
                  align="center" 
                  sx={{
                    color: '#667eea',
                    mb: 1,
                  }}
                >
                  Create Account
                </Typography>
                <Typography variant="body1" color="text.secondary" align="center" sx={{ opacity: 0.8, mb: 2 }}>
                  Join us and start creating amazing reports
                </Typography>
                
                {/* Progress Stepper */}
                <Stepper activeStep={1} alternativeLabel sx={{ width: '100%', mb: 2 }}>
                  {steps.map((label) => (
                    <Step key={label}>
                      <StepLabel 
                        sx={{
                          '& .MuiStepLabel-label': {
                            fontSize: '0.8rem',
                            fontWeight: 500,
                          },
                          '& .MuiStepIcon-root': {
                            color: '#667eea',
                            fontSize: '1.3rem',
                          },
                        }}
                      >
                        {label}
                      </StepLabel>
                    </Step>
                  ))}
                </Stepper>
              </Box>
            </Zoom>



            <form onSubmit={handleSubmit} style={{ width: '100%', flex: 1, display: 'flex', flexDirection: 'column', justifyContent: 'center' }}>
              <TextField
                fullWidth
                label="Full Name"
                name="full_name"
                value={formData.full_name}
                onChange={handleChange}
                margin="dense"
                required
                size="small"
                InputProps={{
                  sx: { 
                    borderRadius: 2,
                    fontSize: '0.95rem',
                    '&:hover .MuiOutlinedInput-notchedOutline': {
                      borderColor: '#667eea',
                    },
                    '&.Mui-focused .MuiOutlinedInput-notchedOutline': {
                      borderColor: '#667eea',
                    },
                  },
                  startAdornment: (
                    <InputAdornment position="start">
                      <BadgeIcon sx={{ color: '#667eea', opacity: 0.7, fontSize: 22 }} />
                    </InputAdornment>
                  ),
                }}
                InputLabelProps={{
                  sx: {
                    fontSize: '1.05rem',
                    '&.Mui-focused': {
                      color: '#667eea',
                    },
                  },
                }}
              />
              <TextField
                fullWidth
                label="Username"
                name="username"
                value={formData.username}
                onChange={handleChange}
                margin="dense"
                required
                size="small"
                InputProps={{
                  sx: { 
                    borderRadius: 2,
                    fontSize: '0.95rem',
                    '&:hover .MuiOutlinedInput-notchedOutline': {
                      borderColor: '#667eea',
                    },
                    '&.Mui-focused .MuiOutlinedInput-notchedOutline': {
                      borderColor: '#667eea',
                    },
                  },
                  startAdornment: (
                    <InputAdornment position="start">
                      <PersonIcon sx={{ color: '#667eea', opacity: 0.7, fontSize: 22 }} />
                    </InputAdornment>
                  ),
                }}
                InputLabelProps={{
                  sx: {
                    fontSize: '1.05rem',
                    '&.Mui-focused': {
                      color: '#667eea',
                    },
                  },
                }}
              />
              <TextField
                fullWidth
                label="Email"
                name="email"
                type="email"
                value={formData.email}
                onChange={handleChange}
                margin="dense"
                required
                size="small"
                InputProps={{
                  sx: { 
                    borderRadius: 2,
                    fontSize: '0.95rem',
                    '&:hover .MuiOutlinedInput-notchedOutline': {
                      borderColor: '#667eea',
                    },
                    '&.Mui-focused .MuiOutlinedInput-notchedOutline': {
                      borderColor: '#667eea',
                    },
                  },
                  startAdornment: (
                    <InputAdornment position="start">
                      <EmailIcon sx={{ color: '#667eea', opacity: 0.7, fontSize: 22 }} />
                    </InputAdornment>
                  ),
                }}
                InputLabelProps={{
                  sx: {
                    fontSize: '1.05rem',
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
                margin="dense"
                required
                size="small"
                InputProps={{
                  sx: { 
                    borderRadius: 2,
                    fontSize: '0.95rem',
                    '&:hover .MuiOutlinedInput-notchedOutline': {
                      borderColor: '#667eea',
                    },
                    '&.Mui-focused .MuiOutlinedInput-notchedOutline': {
                      borderColor: '#667eea',
                    },
                  },
                  startAdornment: (
                    <InputAdornment position="start">
                      <PersonAddAlt1Icon sx={{ color: '#667eea', opacity: 0.7, fontSize: 22 }} />
                    </InputAdornment>
                  ),
                  endAdornment: (
                    <InputAdornment position="end">
                      <IconButton
                        onClick={() => setShowPassword(!showPassword)}
                        edge="end"
                        size="small"
                        sx={{ color: '#667eea' }}
                      >
                        {showPassword ? <VisibilityOffIcon sx={{ fontSize: 20 }} /> : <VisibilityIcon sx={{ fontSize: 20 }} />}
                      </IconButton>
                    </InputAdornment>
                  ),
                }}
                InputLabelProps={{
                  sx: {
                    fontSize: '1.05rem',
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
                size="medium"
                disabled={isLoading}
                sx={{ 
                  mt: 2, 
                  mb: 2,
                  borderRadius: 2, 
                  fontWeight: 600, 
                  fontSize: 15, 
                  py: 1.2, 
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
                  'Create Account'
                )}
              </Button>
            </form>

            <Divider sx={{ width: '100%', my: 2, opacity: 0.3 }} />

            <Box sx={{ mt: 1, textAlign: 'center', width: '100%' }}>
              <Typography variant="body2" color="text.secondary" sx={{ fontSize: '0.9rem' }}>
                Already have an account?{' '}
                <Link 
                  to="/login" 
                  style={{ 
                    textDecoration: 'none', 
                    color: '#667eea', 
                    fontWeight: 600,
                    transition: 'color 0.3s ease',
                  }}
                  onMouseEnter={(e) => e.target.style.color = '#5a6fd8'}
                  onMouseLeave={(e) => e.target.style.color = '#667eea'}
                >
                  Sign in here
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

export default Register; 