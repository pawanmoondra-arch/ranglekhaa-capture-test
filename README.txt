RANGLEKHAA FRESH V1

START CLEAN:
1. Apps Script: replace all Code.gs with this fresh Code.gs.
2. Deploy as a NEW web app version. Execute as Me. Keep access settings used by the current test.
3. GitHub: delete old frontend files and upload:
   index.html
   manifest.json
   sw.js
   icon-192.png
   icon-512.png
4. Open GitHub Pages in laptop browser first.
5. Confirm the page says FRESH-V1 via the backend health endpoint.
6. Then test RKL26-001.

WORKFLOW:
Naveen: capture -> replace/fix -> SUBMIT & LOCK.
Pawan: only unlock after lock.
No approval/pending state exists in this fresh workflow.

IMPORTANT:
A WORKFLOW sheet is created automatically. It is intentionally separate from the old CAPTURE_LOG, so old PENDING_APPROVAL records cannot affect the new lock state.
