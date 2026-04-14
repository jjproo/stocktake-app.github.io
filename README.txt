INVENTORY UI PERFORMANCE FIXES

Why it was slow:
1. The app was opened with file://, so the browser blocked the manifest and service worker.
2. Every tab click ran renderAll(), which rebuilt every table, even hidden ones.
3. Row-level listeners were reattached after every render.
4. Search inputs rendered immediately on every keypress.
5. Service worker cached an external CDN file, which can fail during offline/local opening.

What was changed:
- Service worker only runs on http/https/localhost.
- Manifest is injected only on supported contexts.
- Tab switching renders only the active tab.
- Search inputs are debounced.
- Event delegation is used for delete/count actions.
- Table rendering uses document fragments and row limits.
- Service worker cache list was simplified.

Recommended way to run:
- VS Code Live Server, XAMPP, Laragon, Python local server, or any localhost web server.
- Do NOT double-click index.html and open as file:// if you want PWA/offline features.

Quick run options:
Python:
  python -m http.server 8000
Then open:
  http://localhost:8000
