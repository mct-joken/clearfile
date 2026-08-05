import React from 'react';
import { createRoot } from 'react-dom/client';

const rootElement = document.createElement('div');
rootElement.id = 'onedrive-extension-root';
document.body.appendChild(rootElement);

const root = createRoot(rootElement);

root.render(
  <div style={{
    position: 'fixed',
    bottom: '20px',
    right: '20px',
    padding: '20px',
    backgroundColor: '#ffffff',
    border: '2px solid #0078d4',
    borderRadius: '8px',
    boxShadow: '0 4px 6px rgba(0,0,0,0.1)',
    zIndex: 99999,
    color: '#333333',
    fontFamily: 'Segoe UI, sans-serif'
  }}>
    <h2 style={{ margin: '0 0 10px 0', fontSize: '16px', color: '#0078d4' }}>
      OneDrive Extension
    </h2>
    <p style={{ margin: 0, fontSize: '14px' }}>
      React UI is successfully injected.
    </p>
  </div>
);