import React from 'react';
import { Navigate } from 'react-router-dom';
import axios from 'axios';

function PrivateRoute({ children }) {
  const [isAuthenticated, setIsAuthenticated] = React.useState(null);

  React.useEffect(() => {
    const checkAuth = async () => {
      try {
        await axios.get(`${process.env.REACT_APP_API_URL}/api/user`);
        setIsAuthenticated(true);
      } catch (error) {
        setIsAuthenticated(false);
      }
    };
    checkAuth();
  }, []);

  if (isAuthenticated === null) {
    return null; // or a loading spinner
  }

  return isAuthenticated ? children : <Navigate to="/login" />;
}

export default PrivateRoute; 