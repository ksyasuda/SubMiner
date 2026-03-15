import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import { App } from './App';
import './styles/globals.css';

const isOverlay = new URLSearchParams(window.location.search).has('overlay');
if (isOverlay) {
  document.body.classList.add('overlay-mode');
}

const root = document.getElementById('root');
if (root) {
  createRoot(root).render(
    <StrictMode>
      <App />
    </StrictMode>
  );
}
