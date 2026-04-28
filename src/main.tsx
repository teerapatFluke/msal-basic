import React from 'react';
import ReactDOM from 'react-dom/client';
import { PublicClientApplication } from '@azure/msal-browser';
import { MsalProvider } from '@azure/msal-react';
import App from './App';
import './index.css';

const tenantId = import.meta.env.VITE_TENANT_ID;
const clientId = import.meta.env.VITE_CLIENT_ID;

if (!tenantId || !clientId) {
  throw new Error('Missing VITE_TENANT_ID or VITE_CLIENT_ID in .env');
}

const msalInstance = new PublicClientApplication({
  auth: {
    clientId,
    authority: `https://login.microsoftonline.com/${tenantId}`,
    redirectUri: window.location.origin
  },
  cache: {
    cacheLocation: 'sessionStorage'
  }
});

ReactDOM.createRoot(document.getElementById('root')!).render(
  <React.StrictMode>
    <MsalProvider instance={msalInstance}>
      <App />
    </MsalProvider>
  </React.StrictMode>
);
