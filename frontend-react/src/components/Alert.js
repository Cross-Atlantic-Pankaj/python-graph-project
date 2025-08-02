import React from 'react';
import {
  Alert as MuiAlert,
  Snackbar,
  Box,
  Typography,
  IconButton,
  Fade,
  Zoom,
} from '@mui/material';
import {
  CheckCircle as CheckCircleIcon,
  Error as ErrorIcon,
  Warning as WarningIcon,
  Info as InfoIcon,
  Close as CloseIcon,
} from '@mui/icons-material';

const Alert = ({ 
  open, 
  onClose, 
  title, 
  message, 
  severity = 'success',
  autoHideDuration = 4000,
  position = 'top-right'
}) => {
  const getIcon = () => {
    switch (severity) {
      case 'success':
        return <CheckCircleIcon sx={{ fontSize: 20 }} />;
      case 'error':
        return <ErrorIcon sx={{ fontSize: 20 }} />;
      case 'warning':
        return <WarningIcon sx={{ fontSize: 20 }} />;
      case 'info':
        return <InfoIcon sx={{ fontSize: 20 }} />;
      default:
        return <CheckCircleIcon sx={{ fontSize: 20 }} />;
    }
  };

  const getSeverityColor = () => {
    switch (severity) {
      case 'success':
        return '#4caf50';
      case 'error':
        return '#f44336';
      case 'warning':
        return '#ff9800';
      case 'info':
        return '#2196f3';
      default:
        return '#4caf50';
    }
  };

  const getBackgroundColor = () => {
    switch (severity) {
      case 'success':
        return 'linear-gradient(135deg, #4caf50 0%, #45a049 100%)';
      case 'error':
        return 'linear-gradient(135deg, #f44336 0%, #d32f2f 100%)';
      case 'warning':
        return 'linear-gradient(135deg, #ff9800 0%, #f57c00 100%)';
      case 'info':
        return 'linear-gradient(135deg, #2196f3 0%, #1976d2 100%)';
      default:
        return 'linear-gradient(135deg, #4caf50 0%, #45a049 100%)';
    }
  };

  return (
    <Snackbar
      open={open}
      autoHideDuration={autoHideDuration}
      onClose={onClose}
      anchorOrigin={{ 
        vertical: position.includes('top') ? 'top' : 'bottom',
        horizontal: position.includes('right') ? 'right' : 'left'
      }}
      TransitionComponent={Zoom}
      sx={{
        '& .MuiSnackbar-root': {
          top: 24,
          right: 24,
        }
      }}
    >
      <Fade in={open} timeout={300}>
        <Box
          sx={{
            background: getBackgroundColor(),
            color: 'white',
            borderRadius: 3,
            boxShadow: '0 8px 24px rgba(0, 0, 0, 0.15)',
            minWidth: 320,
            maxWidth: 400,
            p: 2.5,
            position: 'relative',
            overflow: 'hidden',
            '&::before': {
              content: '""',
              position: 'absolute',
              top: 0,
              left: 0,
              right: 0,
              bottom: 0,
              background: 'radial-gradient(circle at 20% 80%, rgba(255, 255, 255, 0.1) 0%, transparent 50%)',
              zIndex: 1,
            }
          }}
        >
          {/* Close Button */}
          <IconButton
            onClick={onClose}
            sx={{
              position: 'absolute',
              top: 8,
              right: 8,
              color: 'white',
              opacity: 0.8,
              zIndex: 3,
              '&:hover': {
                opacity: 1,
                backgroundColor: 'rgba(255, 255, 255, 0.1)',
              },
              transition: 'all 0.2s ease',
            }}
            size="small"
          >
            <CloseIcon sx={{ fontSize: 16 }} />
          </IconButton>

          {/* Content */}
          <Box sx={{ position: 'relative', zIndex: 2 }}>
            <Box sx={{ display: 'flex', alignItems: 'flex-start', gap: 2, pr: 3 }}>
              {/* Icon */}
              <Box
                sx={{
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'center',
                  width: 32,
                  height: 32,
                  borderRadius: '50%',
                  backgroundColor: 'rgba(255, 255, 255, 0.2)',
                  backdropFilter: 'blur(10px)',
                  flexShrink: 0,
                }}
              >
                {getIcon()}
              </Box>

              {/* Text Content */}
              <Box sx={{ flex: 1, minWidth: 0 }}>
                {title && (
                  <Typography
                    variant="subtitle2"
                    sx={{
                      fontWeight: 600,
                      mb: 0.5,
                      fontSize: '0.9rem',
                      lineHeight: 1.2,
                    }}
                  >
                    {title}
                  </Typography>
                )}
                {message && (
                  <Typography
                    variant="body2"
                    sx={{
                      opacity: 0.95,
                      fontSize: '0.85rem',
                      lineHeight: 1.4,
                      wordBreak: 'break-word',
                    }}
                  >
                    {message}
                  </Typography>
                )}
              </Box>
            </Box>
          </Box>

          {/* Progress Bar */}
          <Box
            sx={{
              position: 'absolute',
              bottom: 0,
              left: 0,
              right: 0,
              height: 2,
              backgroundColor: 'rgba(255, 255, 255, 0.2)',
              zIndex: 2,
            }}
          >
            <Box
              sx={{
                height: '100%',
                backgroundColor: 'rgba(255, 255, 255, 0.8)',
                animation: 'progress 4s linear',
                '@keyframes progress': {
                  '0%': { width: '100%' },
                  '100%': { width: '0%' },
                },
              }}
            />
          </Box>
        </Box>
      </Fade>
    </Snackbar>
  );
};

export default Alert; 