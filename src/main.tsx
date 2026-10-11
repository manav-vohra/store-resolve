import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import './styles.css'
import App from './App'

function registerServiceWorker() {
  if (!('serviceWorker' in navigator)) return
  // Defer registration until the page is interactive; a failure must never
  // break the app shell.
  const register = () =>
    navigator.serviceWorker.register('/sw.js').catch((error) => {
      console.warn('[pwa] service worker registration failed', error)
    })
  if (document.readyState === 'complete') register()
  else window.addEventListener('load', register, { once: true })
}

registerServiceWorker()

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <App />
  </StrictMode>,
)
