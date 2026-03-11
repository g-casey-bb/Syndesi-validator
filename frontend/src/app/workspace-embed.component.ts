import { Component } from '@angular/core';

/**
 * Minimal view for the /excel-google-workspace route when the app is loaded inside an iframe.
 * Without this route, the iframe would load the catch-all and show the full app (header, nav, content),
 * which looks like "an embedded copy of the whole page".
 * The actual workspace uploader UI is intended to be served by the separate Excel Import app on port 4200;
 * when only Syndesi runs (e.g. on 4200), this placeholder is shown in the iframe instead of the full app.
 */
@Component({
  selector: 'app-workspace-embed',
  standalone: true,
  template: `
    <div class="workspace-embed">
      <p class="workspace-embed-message">
        This iframe loads the uploader from <strong>http://localhost:4200/excel-google-workspace</strong>.
      </p>
      <p class="workspace-embed-message">
        You are seeing this because nothing is serving that URL. Start the <strong>Excel Import frontend</strong> on port 4200 (run <code>run-frontend.bat</code> in ExcelImport-project, or <code>npm start</code> in its <code>frontend</code> folder). Then open Syndesi at <strong>http://localhost:4201</strong> so the iframe can load the uploader from 4200.
      </p>
    </div>
  `,
  styles: [`
    .workspace-embed {
      padding: 1rem;
      font-size: 0.9rem;
      color: #555;
    }
    .workspace-embed-message {
      margin: 0 0 0.5rem 0;
    }
    .workspace-embed code {
      background: #eee;
      padding: 0.1rem 0.3rem;
      border-radius: 3px;
    }
  `],
})
export class WorkspaceEmbedComponent {}
