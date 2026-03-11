# Ports and how to start

Use these ports so the Syndesi iframe can load the Excel Import app.

| Service | Port | Start command |
|---------|------|----------------|
| **Syndesi backend** | 3000 | From `Syndesi\backend`: `run-backend.bat` or `node server.js` |
| **Syndesi frontend** | 4201 | From `Syndesi\frontend`: `run-frontend.bat` or `npx ng serve --port 4201` |
| **Excel Import backend** | 3001 | From `ExcelImport-project`: `run-backend.bat` |
| **Excel Import frontend** | 4200 | From `ExcelImport-project`: `run-frontend.bat` or `cd frontend` then `npm start` |

- Open **Syndesi** at: **http://localhost:4201**
- The Excel Import iframe loads from: **http://localhost:4200**

In PowerShell, run batch files with `.\run-backend.bat` (include `.\`).
