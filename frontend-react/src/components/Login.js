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
} from '@mui/material';
import axios from 'axios';
import LockOutlinedIcon from '@mui/icons-material/LockOutlined';

function Login() {
  const navigate = useNavigate();
  const [formData, setFormData] = useState({
    username: '',
    password: '',
  });
  const [error, setError] = useState('');

  const handleChange = (e) => {
    setFormData({
      ...formData,
      [e.target.name]: e.target.value,
    });
  };

  const handleSubmit = async (e) => {
    e.preventDefault();
    try {
      const response = await axios.post(`${process.env.REACT_APP_API_URL}/api/login`, formData);
      if (response.data.message === 'Login successful') {
        navigate('/dashboard');
      }
    } catch (err) {
      setError(err.response?.data?.error || 'Login failed');
    }
  };

  return (
    <Container maxWidth="sm" disableGutters>
      <Box
        sx={{
          minHeight: '100vh',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          background: 'linear-gradient(135deg, #e0e7ff 0%, #f8fafc 100%)',
        }}
      >
        <Paper
          elevation={6}
          sx={{
            p: 5,
            width: '100%',
            borderRadius: 4,
            boxShadow: '0 8px 32px 0 rgba(31, 38, 135, 0.15)',
            display: 'flex',
            flexDirection: 'column',
            alignItems: 'center',
          }}
        >
          <Box sx={{ mb: 2, display: 'flex', flexDirection: 'column', alignItems: 'center' }}>
            <Box
              sx={{
                background: 'linear-gradient(135deg, #6366f1 30%, #818cf8 90%)',
                borderRadius: '50%',
                width: 64,
                height: 64,
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                mb: 1,
                boxShadow: '0 2px 8px rgba(99,102,241,0.15)'
              }}
            >
              <LockOutlinedIcon sx={{ color: '#fff', fontSize: 36 }} />
            </Box>
            <Typography variant="h4" component="h1" fontWeight={700} gutterBottom align="center" color="primary">
              Welcome Back
            </Typography>
          </Box>
          {error && (
            <Alert severity="error" sx={{ mb: 2, width: '100%' }}>
              {error}
            </Alert>
          )}
          <form onSubmit={handleSubmit} style={{ width: '100%' }}>
            <TextField
              fullWidth
              label="Username"
              name="username"
              value={formData.username}
              onChange={handleChange}
              margin="normal"
              required
              autoFocus
              InputProps={{ sx: { borderRadius: 2 } }}
            />
            <TextField
              fullWidth
              label="Password"
              name="password"
              type="password"
              value={formData.password}
              onChange={handleChange}
              margin="normal"
              required
              InputProps={{ sx: { borderRadius: 2 } }}
            />
            <Button
              type="submit"
              fullWidth
              variant="contained"
              size="large"
              sx={{ mt: 3, borderRadius: 2, fontWeight: 600, fontSize: 18, py: 1.5, background: 'linear-gradient(90deg, #6366f1 0%, #818cf8 100%)' }}
            >
              Sign In
            </Button>
          </form>
          <Box sx={{ mt: 3, textAlign: 'center', width: '100%' }}>
            <Typography variant="body2">
              New to this website?{' '}
              <Link to="/register" style={{ textDecoration: 'none', color: '#6366f1', fontWeight: 500 }}>
                Register here
              </Link>
            </Typography>
          </Box>
        </Paper>
      </Box>
    </Container>
  );
}

export default Login; 