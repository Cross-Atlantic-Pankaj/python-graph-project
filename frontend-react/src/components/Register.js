import React, { useState } from 'react';
import { useNavigate, Link } from 'react-router-dom';
import {
  Container,
  Paper,
  TextField,
  Button,
  Typography,
  Box,
  Alert,
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
        navigate('/login');
      }
    } catch (err) {
      setError(err.response?.data?.error || 'Registration failed. Please try again.');
    } finally {
      setIsLoading(false);
    }
  };

  const steps = ['Personal Info', 'Account Details', 'Security'];

  return (
    <Container maxWidth="sm" disableGutters>
      <Box
        sx={{
          minHeight: '100vh',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          background: 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)',
          position: 'relative',
          overflow: 'hidden',
          '&::before': {
            content: '""',
            position: 'absolute',
            top: 0,
            left: 0,
            right: 0,
            bottom: 0,
            background: 'radial-gradient(circle at 20% 80%, rgba(120, 119, 198, 0.3) 0%, transparent 50%), radial-gradient(circle at 80% 20%, rgba(255, 119, 198, 0.3) 0%, transparent 50%)',
            zIndex: 1,
          }
        }}
      >
        <Fade in timeout={800}>
          <Paper
            elevation={24}
            sx={{
              p: { xs: 3, sm: 5 },
              width: '100%',
              maxWidth: 500,
              borderRadius: 4,
              boxShadow: '0 20px 60px rgba(0, 0, 0, 0.3)',
              display: 'flex',
              flexDirection: 'column',
              alignItems: 'center',
              position: 'relative',
              zIndex: 2,
              background: 'rgba(255, 255, 255, 0.95)',
              backdropFilter: 'blur(10px)',
              border: '1px solid rgba(255, 255, 255, 0.2)',
            }}
          >
            <Zoom in timeout={1000}>
              <Box sx={{ mb: 3, display: 'flex', flexDirection: 'column', alignItems: 'center' }}>
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
                    boxShadow: '0 8px 32px rgba(102, 126, 234, 0.4)',
                    animation: 'pulse 2s infinite',
                    '@keyframes pulse': {
                      '0%': {
                        boxShadow: '0 8px 32px rgba(102, 126, 234, 0.4)',
                      },
                      '50%': {
                        boxShadow: '0 8px 32px rgba(102, 126, 234, 0.6)',
                      },
                      '100%': {
                        boxShadow: '0 8px 32px rgba(102, 126, 234, 0.4)',
                      },
                    },
                  }}
                >
                  <PersonAddAlt1Icon sx={{ color: '#fff', fontSize: 40 }} />
                </Box>
                <Typography 
                  variant="h4" 
                  component="h1" 
                  fontWeight={700} 
                  gutterBottom 
                  align="center" 
                  sx={{
                    background: 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)',
                    backgroundClip: 'text',
                    WebkitBackgroundClip: 'text',
                    WebkitTextFillColor: 'transparent',
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
                            fontSize: '0.75rem',
                            fontWeight: 500,
                          },
                          '& .MuiStepIcon-root': {
                            color: '#667eea',
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

            {error && (
              <Zoom in timeout={300}>
                <Alert 
                  severity="error" 
                  sx={{ 
                    mb: 3, 
                    width: '100%', 
                    borderRadius: 2,
                    boxShadow: '0 4px 12px rgba(244, 67, 54, 0.15)',
                  }}
                >
                  {error}
                </Alert>
              </Zoom>
            )}

            <form onSubmit={handleSubmit} style={{ width: '100%' }}>
              <TextField
                fullWidth
                label="Full Name"
                name="full_name"
                value={formData.full_name}
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
                      <BadgeIcon sx={{ color: '#667eea', opacity: 0.7 }} />
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
                label="Username"
                name="username"
                value={formData.username}
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
                label="Email"
                name="email"
                type="email"
                value={formData.email}
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
                      <EmailIcon sx={{ color: '#667eea', opacity: 0.7 }} />
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
                      <PersonAddAlt1Icon sx={{ color: '#667eea', opacity: 0.7 }} />
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
                  mt: 4, 
                  mb: 2,
                  borderRadius: 2, 
                  fontWeight: 600, 
                  fontSize: 16, 
                  py: 1.5, 
                  background: 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)',
                  boxShadow: '0 8px 24px rgba(102, 126, 234, 0.3)',
                  transition: 'all 0.3s ease',
                  '&:hover': {
                    background: 'linear-gradient(135deg, #5a6fd8 0%, #6a4190 100%)',
                    boxShadow: '0 12px 32px rgba(102, 126, 234, 0.4)',
                    transform: 'translateY(-2px)',
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

            <Box sx={{ mt: 2, textAlign: 'center', width: '100%' }}>
              <Typography variant="body2" color="text.secondary">
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
    </Container>
  );
}

export default Register; 