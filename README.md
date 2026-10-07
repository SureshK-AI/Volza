# Volza shipment export

Export the shipment results from any Volza workspace URL to an Excel workbook
with the TypeScript exporter, including every results page available to the
logged-in account.

## Setup

1. Install the project dependencies with `npm install`.
2. Start Chrome with remote debugging enabled, sign in to Volza, and keep that
   browser open. For example, on Windows:

   ```powershell
   & "$env:ProgramFiles\Google\Chrome\Application\chrome.exe" `
     --remote-debugging-port=9222 `
     --user-data-dir="$env:LOCALAPPDATA\VolzaExportProfile"
   ```

   Sign in to Volza - https://www.volza.com/signin-wizard-step-1/ in that Chrome profile if needed. To use another Chrome
   executable or debugging port, set `VOLZA_CDP_ENDPOINT` to its CDP address.
3. Run either command from the project folder:

   ```powershell
   npx tsx .\tests\export-volza.ts
   # or
   npm run export:excel
   ```

   Both commands run the same TypeScript script. `npm run export:excel` is an
   optional shortcut defined in `package.json`; use `npx tsx
   .\tests\export-volza.ts` to launch the script directly. Enter the logged-in
   search-results URL when prompted, or press Enter to use the included Volza
   URL.

The script opens the URL in a temporary tab in the authenticated browser,
sets the results page size to 50, and collects every named shipment-table
column and cell value from each page until the site disables its Next button.
The workbook includes the visible search criteria above the shipment table and
is saved in the current directory with a filename based on those criteria.
Unnamed utility columns (such as row controls) are excluded. The export is
limited to the results pages accessible to the account in Volza.

## Note for later

Run this exporter with `npx tsx .\tests\export-volza.ts`. You can also use
`npm run export:excel`; it is a shortcut for the same command.

Do not run it with `npx playwright test ... --headed`. That command is for
running Playwright test files. The exporter is a regular TypeScript script, not
a test file, so `tsx` is the tool that starts it. It connects to the Chrome
window that is already open and signed in to Volza, so you do not need the
`--headed` option.
