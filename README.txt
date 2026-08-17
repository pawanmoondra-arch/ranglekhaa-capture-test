RANGLEKHAA V2.2 — DIRECT LOCK
FROZEN PERMISSION MODEL

Partner (Naveen):
- Capture
- Replace/repair
- Complete
- Finalize & lock
- View locked articles

Partner & Admin (Pawan):
- Everything Partner can do
- Unlock locked stages
- Edit/replace after unlock
- Finalize & lock again

NO PENDING APPROVAL STATE.

The old PENDING_APPROVAL rows in CAPTURE_LOG are historical test records only.
V2.2 calculates state only from the latest LOCKED / UNLOCKED workflow event.

INSTALL BACKEND:
1. Replace Code.gs in the existing Ranglekhaa Apps Script backend.
2. Save.
3. Deploy -> Manage deployments -> Edit -> New version -> Deploy.
4. Keep the existing Web App URL.

BACKEND TEST:
Open the Web App URL. It must return version V2.2.

FRONTEND:
Replace index.html, manifest.json, sw.js, icon-192.png, icon-512.png in GitHub Pages.
Commit and wait for GitHub Pages to publish.

TEST ORDER:
1. RKL26-001 -> Development should show 2/2 COMPLETE.
2. Partner (Naveen) -> SUBMIT & LOCK.
3. Confirm status endpoint says state LOCKED and locked true.
4. Confirm Drive photos remain unchanged.
5. Partner cannot replace while locked.
6. Switch test user to Partner & Admin (Pawan) only for the unlock test.
7. Pawan -> UNLOCK FOR CORRECTION.
8. Confirm state DRAFT.
9. Naveen/Partner can replace again.
10. Complete and lock again.

TEST USER SWITCH:
The PWA uses localStorage key rkl_partner_user. For the test build, the UI defaults to nsoni9068@gmail.com. To switch to the admin test account, use the browser console:
localStorage.setItem('rkl_partner_user','pawanmoondra@gmail.com'); location.reload();
This is TEST MODE only. Before production, replace client-supplied email with real Google authentication and backend identity enforcement.
