/* ==========================================================================
   DocBrisk Looker Studio Bridge
   Looker Studio has no public API to create a fully-populated report from
   arbitrary code — what it DOES have is a documented "Report Linking" URL
   scheme that opens a NEW report, copied from an existing template report,
   with a chosen data source already attached:
   https://developers.google.com/looker-studio/integrate/linking-api

   That still needs a real Google Sheet as the data source, and creating a
   Google Sheet from inside Excel requires Google OAuth + the Sheets API —
   which needs a backend (this is exactly what your existing docbrisk-api
   Cloudflare Worker is for; see README "Full automation (v2)"). Until
   that's wired up, this module gets the person to one Google Sheet paste
   away from a working Looker Studio report — everything else is generated
   for them, including the CSV formatted specifically for that import.
   ========================================================================== */

const LookerStudio = (() => {

  /** CSV export of the exact rows the dashboard was built from, so pasting
   *  straight into a new Google Sheet reproduces the same KPIs there. */
  function buildCsv(headers, rows) {
    const esc = (v) => {
      const s = String(v ?? '');
      return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
    };
    const lines = [headers.map(esc).join(',')];
    rows.forEach((r) => lines.push(r.map(esc).join(',')));
    return lines.join('\r\n');
  }

  /** Pull the spreadsheet ID and sheet (gid) out of a Google Sheets URL the
   *  user pastes in after importing the CSV — the one manual step. */
  function parseSheetUrl(url) {
    const idMatch = String(url || '').match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    if (!idMatch) return null;
    const gidMatch = String(url).match(/[?#&]gid=(\d+)/);
    return { spreadsheetId: idMatch[1], gid: gidMatch ? gidMatch[1] : '0' };
  }

  /**
   * Build the Report Linking URL. `templateReportId` is a one-time setup
   * value — see README "Looker Studio setup" — left blank in templates.js
   * until you've created and published a template report of your own.
   */
  function buildReportUrl({ templateReportId, reportName, spreadsheetId, gid, placeholderDsId }) {
    if (!templateReportId) return null;
    const base = 'https://lookerstudio.google.com/reporting/create';
    const params = new URLSearchParams();
    params.set('c.reportId', templateReportId);
    if (reportName) params.set('r.reportName', reportName);
    const ds = placeholderDsId || 'ds0';
    params.set(`ds.${ds}.connector`, 'googleSheets');
    params.set(`ds.${ds}.spreadsheetId`, spreadsheetId);
    params.set(`ds.${ds}.worksheetId`, gid || '0');
    return `${base}?${params.toString()}`;
  }

  function downloadCsv(filename, csvContent) {
    const blob = new Blob(['\uFEFF' + csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = filename;
    document.body.appendChild(a); a.click(); a.remove();
    URL.revokeObjectURL(url);
  }

  return { buildCsv, parseSheetUrl, buildReportUrl, downloadCsv };
})();

if (typeof module !== 'undefined' && module.exports) module.exports = LookerStudio;
