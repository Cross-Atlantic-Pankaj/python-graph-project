import axios from 'axios';

// Use backend URL directly
const api = axios.create({
  baseURL: 'http://13.235.141.55:5001/api',
  headers: {
    'Content-Type': 'application/json',
  },
  withCredentials: true,
});

export default api;
