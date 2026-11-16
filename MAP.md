# Priority Work List — Codebase Map

## Files (Root GitHub Pages Bundle)
- `index.html`: App shell, tabs, tables, 10 modal overlays, script load order (`XLSX` -> Firebase -> `firebase-sync.js` -> `app.js`)
- `earnings.html`: Public earnings viewer plus an admin-PIN-protected RC CSV importer with user/department setup
- `style.css`: Dark/light CSS variables, layout, tables, tick/BER counters, modals, print stylesheet, role visibility
- `firebase-sync.js`: Firebase init, seed tables (`DCD_PARTS`, `GRADE_SEED`, `DEPT_SEED`), auth/PIN, requests, live sync, part CRUD, admin upload
- `app.js`: Theme, Priority List loader/filter/sort, stock math, tick/BER counters, Completed Today, History, KPI CSV importer, Commission Tracker
- `logo.png`: Vista Technology Support PNG logo shared by login screen & header

## HTML Modal IDs (`index.html`)
- `login-screen` (public earnings viewer, PIN-gated sign-in & dept notice) | `modal-overlay` (Add Part to Completed)
- `admin-overlay` (Admin PowerBI/Backorder file upload & download) | `adjust-overlay` (Stock adjustment)
- `newpart-overlay` (Admin add custom part) | `manage-overlay` (Manage parts, `?` ungraded, `!` new, `~` none, Excel export/import)
- `hidden-overlay` (Parts hidden by keyword/AT DEPLOYMENT filter) | `pin-overlay` (Enter/Create/Change PIN)
- `request-overlay` (User change/new part request) | `requests-overlay` (Admin approvals & changelog)
- `links-overlay` (Admin personal `?user=` links)

## Key Element IDs (`index.html`)
- **Header & Auth**: `active-user-badge`, `admin-btn`, `pin-btn`, `theme-toggle`, `date-badge`, `last-updated`, `live-status`, `dept-select`
- **Public Earnings (`earnings.html`)**: `earnings-user-select`, `public-earnings-month`, `public-earnings-status`, `public-earnings-content`, `public-earnings-days`, `public-earnings-bonus`, `public-earnings-total-bonus`, `public-earnings-target-bars`, `public-earnings-body`; admin RC import uses `rc-earnings-admin`, `rc-earnings-file`, `rc-earnings-people-body`
- **Stats & Tabs**: `loading-bar`, `error-msg`, `stats-bar` (`stat-r1`, `stat-r2`, `stat-total`, `stat-rep-total`, `stat-stock-total`, `stat-done`, `stat-ber-today`), `tab-bar` (`tab-count-main`, `tab-count-done`, `tab-count-history`)
- **Priority List**: `actions-bar` (`search-input`, `show-done-btn`, `new-btn`, `hidden-btn`, `requests-btn`, `print-menu`, `row-count`), `table-wrap`, `table-body`, `zero-section`, `zero-body`
- **Completed & History**: `completed-wrap`, `earnings-bar`, `completed-body`, `history-wrap`, `month-filter-bar`, `month-summary-card`, `ms-grade-breakdown`, `target-tracker` (`bar-18`, `bar-21`, `bar-24`), `history-body`, `ungraded-section`, `hist-import-section` (`import-drop-zone`, `import-preview-body`)

## Firebase Realtime Database Paths (`firebase-sync.js`)
- `users/{uid}/ticks/{TODAY}`, `users/{uid}/ber/{TODAY}`, `users/{uid}/history` (`_meta/adj/{month}`)
- `rcEarningsUsers/{uid}` stores RC-imported user display names and department assignments
- `allTicks/{TODAY}`, `userPins/{uid}`, `files/{powerbi|backorder}`, `filesUploadedAt`, `lastUpdated`
- `dcdParts`, `deptParts/{PC|PCB|Printers}`, `partGrades`, `noDept`, `customParts`, `stockAdjustments`
- `changeRequests`, `changeLog`, `autoAssigned`, `migrations/{dcdMoveToOwnDept1|dcdRemoveTemporary1|gradeFixes1}`

## Main Functions by File
- **`firebase-sync.js`**: `safeKey`, `activeParts`, `hashPin`, `startPinFlow`, `submitPin`, `openChangePin`, `initLoginFromLink`, `loginAs`, `switchUser`, `openLinksModal`, `isAdmin`, `applyRoleUI`, `logChange`, `openRequestModal`, `submitRequest`, `openRequestsModal`, `approveRequest`, `rejectRequest`, `startFirebaseListeners`, `scheduleReload`, `applyDcdMoveOnce`, `fetchFile`, `loadXLSX`, `loadCSV`, `setDCD`, `setPartGrade`, `setPartNoDept`, `setPartDept`, `exportPartsSheet`, `importPartsSheet`, `saveNewPart`, `downloadBackup`, `openAdmin`, `downloadCurrentFile`, `uploadAdminFiles`
- **`app.js`**: `escAttr`, `showToast`, `toggleTheme`, `loadAll`, `loadAllNow`, `buildPartIndex`, `applyKeywordHiding`, `autoAssignFromTeams`, `computeNewParts`, `updatePriorityTotals`, `renderTable`, `setDept`, `sortBy`, `toggleDone`, `resetTicks`, `printList`, `openAdjustModal`, `confirmAdjust`, `addTick`, `removeTick`, `addBer`, `removeBer`, `entryPostCount`, `nextEntry`, `getEffectiveRow`, `updateDoneCount`, `switchTab`, `renderCompleted`, `openModal`, `searchParts`, `confirmAddPart`, `saveToHistory`, `handleImportFile`, `parseImportCSV`, `confirmImport`, `setMonthFilter`, `updateMonthSummary`, `calcCommission`, `updateTargetTracker`, `filterByGrade`, `renderHistory`, `initEarningsViewer`, `parseRCEarningsCSV`, `importRCEarnings`, `loadPublicEarnings`, `adjustPublicDaysWorked`, `adjustDaysWorked`, `recalculateHistory`, `openManageModal`, `renderManageList`, `openHiddenModal`, `openNewPartModal`
