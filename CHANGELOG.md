# Changelog

All notable changes to Smart Org Chart are documented here.

---

## [1.3.1] — 2026-09-23

### Added
- **Custom chart theme** — a fifth **Chart Theme** option, **Custom**, lets an admin pick any accent color for the Directory and Org Chart instead of choosing from the four built-in palettes. The property pane shows a color swatch, a hex input, and R/G/B number inputs, plus a live contrast check that warns when text would be hard to read on the chosen color; text drawn on or from the accent color automatically switches to a readable shade.
- **Custom Attributes** — a new **Custom Attributes** group in the property pane lets an admin surface up to 10 additional Entra ID fields (e.g. `employeeId`, `companyName`, `city`, or `extensionAttribute1`–`15` for on-premises AD attributes like cost center or building) in the Employee Directory and/or the Org Chart's profile card, each with its own display label. Requires the Graph API data source — SharePoint Search has no generic way to read these fields.
- **Total headcount** — in addition to the existing direct-report count, cards in the Org Chart, the profile popup, and Employee Directory rows now also show the total number of people below that person at every level (e.g. "8 direct reports · 33 total"), computed instantly with no extra loading.
- **Refresh data button** — a new header-bar button manually reloads data from Microsoft Graph or SharePoint Search, bypassing the local cache (see below). Its tooltip shows when the data was last loaded.
- **Local data caching** — downloaded user and org data is now cached locally for up to 4 hours (in addition to the existing in-memory cache), so reopening or reloading the page no longer re-downloads the whole directory every time. Admin filter changes (excluded accounts, hidden guests/disabled accounts, etc.) now re-apply to the cached data instantly instead of triggering a fresh download.
- **Keyboard navigation in the org chart tree** — Arrow Up/Down/Left/Right and Home/End now move between cards in the Top Down and Left to Right layouts, with a visible focus ring; the search results dropdown supports Arrow Up/Down and Enter.

### Changed
- **Employee Directory cards resize to fit their content** — a card with several optional fields enabled (custom attributes, the new headcount line, etc.) now grows to show everything instead of being clipped at a fixed height.
- **Admin filter changes no longer re-download the whole directory** — changing a User Filter now re-applies it to already-loaded data in memory; only changing the data source, demo mode, the dotted-line attribute, or which custom attributes are configured triggers a fresh download (needed because those change what's actually requested from Graph).
- **Default Font Size now applies live** — changing the admin's default text scale now updates immediately for any user who hasn't chosen their own font size in Preferences.
- Large internal refactor of the Org Chart component for maintainability (split into focused modules) and performance (memoized rendering, debounced search, batched tree updates) — no user-facing behavior change beyond the fixes below.

### Fixed
- **Keyboard users couldn't expand org chart cards** — pressing Enter or Space on a card's Expand or Focus button opened the profile popup instead of activating that button.
- **Large organizations loaded and rendered slowly** — Expand All and photo loading no longer lock up the page while working through a large tree; the SharePoint Search fallback path now batches its per-manager lookups instead of issuing one request per manager; a throttled Microsoft Graph request is now retried with backoff instead of silently falling back to a partial dataset.
- **Two people configured as each other's manager could duplicate nodes** — this cycle is now detected and broken automatically instead of producing repeated cards or infinite loops.
- **PDF/CSV export and match counts could disagree with what was on screen** when a department filter was active.
- **Saved preferences could crash the org chart** if browser storage held an unexpected value (e.g. after a browser profile sync); saved settings are now validated before use.
- **Broken profile photos left a blank tile** instead of falling back to initials.
- A trailing space in the directory search box prevented an otherwise-matching result from appearing.
- Presence lookups that failed due to a missing permission were retried every 60 seconds instead of backing off.
- The property pane's color picker and "max employees per page" slider could lose focus or stop responding while being dragged, because every change was rebuilding the field's inputs from scratch.
- Various accessibility fixes: the org chart tree now exposes proper `tree`/`treeitem` roles, the profile popup traps focus and returns it to the card that opened it, and several controls that were invisible or unreachable by keyboard now have visible focus states.

---

## [1.3.0] — 2026-07-14

### Added
- **Demo Data auto-fills the top-level user** — turning on **Use Demo Data** in the property pane now fills in **Top-Level User** with the sample CEO automatically when the field is empty, so the demo works immediately with no typing required. An admin-configured value is never overwritten.

### Changed
- **Admin property pane changes now apply immediately** — switching the default org chart layout, default zoom, or any "Org Chart Features" toggle (Find Me, layout toggle, stats bar, department filter, user type filter) in the property pane now takes effect live. Previously some of these required a page reload, and turning a filter or layout toggle off could leave a user stuck with a persisted filter and no control to clear it — turning a toggle off now also resets its state.
- **Faster loading for large organizations on the SharePoint Search data source** — the fallback lookup that supplements unlicensed users' direct reports now issues requests to Microsoft Graph in parallel instead of one at a time, cutting initial load time significantly for tenants with many managers.
- **Faster Expand All and Employee Directory photo loading** — profile photos in the Employee Directory now batch into the UI instead of triggering a re-render per photo, and Org Chart's Expand All injects each newly loaded batch of the tree in a single pass instead of re-cloning the whole tree per node.
- **Keyboard accessibility** — Escape now closes the org chart's layout, department, and user-type filter popups (previously only the person profile card responded to Escape). The profile card is also now exposed as an accessible dialog, with focus moved to its close button when it opens.

### Fixed
- **Org chart could hang on self-managed or circular manager relationships** — an Azure AD account that is its own manager (common for CEOs) or a short manager cycle (A → B → A) could make the org chart render a user as their own descendant, and "Expand All" would loop indefinitely instead of completing. Both cases are now detected and handled safely.
- **Hiding a manager silently deleted their entire reporting branch** — when a User Filter (excluded pattern, hidden job title/department, etc.) hid a manager, everyone who reported to them disappeared from the org chart instead of just the directory. Their reports are now re-linked to the next visible manager above them, so the chart's structure is preserved.
- **Zoom "Reset" ignored the admin-configured default zoom** — the reset button always returned to 100% even when the property pane specified a different default (e.g. 75%). It now returns to whichever default the admin has configured.
- **Unnecessary `?socFocus` URL parameter on every page view** — simply opening the org chart at its default position wrote a focus parameter into the page URL and browser storage, which could also collide between two web part instances on the same page. The parameter now only appears once a user actually navigates away from the default view, and clears again when they return to it.
- **Presence lookups retried a failing permission every 60 seconds** — if the `Presence.Read.All` permission had not been approved, the web part kept retrying the failing Microsoft Graph call every minute indefinitely. It now backs off for 15 minutes after a failure before retrying.

### Removed
- **Unused PNG and vCard export code** — the underlying PNG image export and vCard download functions (already inaccessible from the UI since earlier releases) have been removed, along with the `html2canvas` dependency this dropped from the shipped bundle.

---

## [1.2.1] — 2026-06-16

### Changed
- **Export PDF moved to org chart toolbar** — the PDF export button is now in the org chart toolbar (where the PNG button was), keeping export controls alongside the chart rather than in the header bar.
- **Employee Directory photo sizes increased** — profile photos in the card view are now 80 px (up from 64 px) and in the list view 52 px (up from 40 px). Initials avatars and presence dots scale proportionally.

### Fixed
- **Department filter blanked the entire chart** — when a department filter was applied, the root node (e.g. CEO) was hidden because they were not in the selected department, causing the whole tree to disappear. Ancestor nodes are now kept visible when any of their descendants match the active filter.
- **Connector lines not updated after filtering** — horizontal connector bars between levels retained stale positions after a department, member, or guest filter was changed, or after the zoom level or layout was switched. `_fixConnectorLines` now runs whenever any of these values change.
- **Full-width horizontal connector flash** — the horizontal connector bar between tree levels briefly appeared full-width before JavaScript recalculated its span. The CSS default is now zero-width so there is no visible flash on initial render.

### Removed
- **Header-level PDF download button** — the Download (↓) button that appeared in the header bar next to the Settings gear has been removed. PDF export for the org chart is now available directly in the org chart toolbar.
- **Profile card action buttons simplified** — the Meet (schedule meeting), vCard (download contact), Copy Link, and Call buttons have been removed from the person profile card popup. The remaining actions are Chat in Teams, Email, and Focus.

---

## [1.2.0] — 2026-06-07

### Added
- **Export Employee Directory to CSV** — new Export button in the directory toolbar downloads the current filtered list as a UTF-8 CSV file (Excel-compatible). Exported columns match the user's visible field preferences (name, title, department, office, email, phone).
- **"View from person…" root picker** — a new search field in the org chart toolbar lets any user temporarily re-root the chart at any person in the organization. A cancel button resets back to the admin-configured root.
- **Configurable default org chart zoom** — new admin property pane setting (Default Org Chart Zoom). Options: Auto-fit (default), 50%, 75%, 100%, 125%, 150%. When a fixed zoom is set, the auto-fit behavior is suppressed.
- **Configurable font size** — text scale can now be set both by the admin (Default Font Size, in the property pane Visual Style group) and by each user (Font Size, in the Settings panel). Options range from 75% to 175%.
- **Demo dataset picker in Settings panel** — when Demo Data is enabled, the 150 / 500 / 1,000 person dataset size buttons are now accessible from the Settings panel (gear icon) instead of a banner below the header.

### Changed
- **Demo Data moved to its own property pane group** — the Use Demo Data toggle has moved from the General group to a dedicated Demo group at the bottom of the property pane.
- **Layout names standardized** — the three org chart layouts are now named *Drill-Down*, *Top Down*, and *Left to Right* consistently across the toolbar, layout picker, and property pane.
- **Page size control improved** — the Max employees per page slider in the property pane now includes a companion number input for precise values, in addition to the drag slider.
- **Statistics based on all users** — the stats bar (Total, Members, Guests, Depts) now reflects all visible users in the directory rather than being derived by walking the currently-loaded tree nodes.
- **Zoom level not persisted** — the org chart zoom level is no longer saved to the browser session. Each visit starts at the admin-configured zoom or auto-fit.

### Removed
- **"Avg span" removed from stats bar** — the average reporting span figure has been removed. The stats bar now shows Total, Members, Guests, and Depts.
- **Demo data banner removed** — the banner that appeared below the header in demo mode has been removed; dataset size controls are now in the Settings panel.

---

## [1.1.0] — 2026-06-02

### Changed
- **Zoom minimum lowered to 25%** — the org chart can now be zoomed out to 25% (previously 40%), giving more room to view large hierarchies.
- **Expand All now auto-fits zoom** — after expanding the full organization tree, the view automatically zooms to fit the newly loaded content.
- **Connector lines end at card edges** — the horizontal connector bar in the vertical tree now ends at the center of the leftmost and rightmost cards in each row, rather than extending to the full sub-tree width.
- **Large tree horizontal scrolling fixed** — the full tree is now accessible by scrolling left and right after Expand All, including nodes at the far left and right of wide org charts.

### Removed
- **Export to PNG** — the PNG image export button has been removed from the header bar. PDF export remains available.

---

## [1.0.1] — 2026-05-30

### Added
- **Room account exclusion by pattern** — new *Exclude room accounts by pattern* field in the property pane. Comma-separated patterns (e.g. `conf-, room-, mrm@`) hide matching mailboxes from all views regardless of their account-enabled status, covering room mailboxes that have sign-in enabled as well as disabled ones.
- **Office location on org chart cards** — the office/location field now appears on node cards in the tree and drill-down views when enabled in User Preferences.
- **Configurable data source** — choose between Graph API (live, no indexing delay), SharePoint Search (legacy), or Auto (Graph with SP Search fallback).
- **List view in Employee Directory** — toggle between card grid and compact single-row list view.
- **Admin user filters** — property pane options to hide Azure AD guest accounts, hide disabled accounts, restrict to tenant domain, and exclude accounts by name pattern.

### Changed
- **Consistent accent color** — card borders, avatar backgrounds, job title text, and org chart connector lines now all follow the SharePoint site's color theme instead of using per-department hashed colors.
- **Expand All loads the complete org** — clicking Expand All in tree modes now performs a full BFS load from Microsoft Graph, fetching every level of the hierarchy rather than only expanding the pre-loaded nodes.
- **Horizontal tree layout** — the left-right tree now uses the same recursive layout as the top-down tree: each manager's reports appear in a column to their right, expanding outward as branches are opened.
- **Direct report count** — the badge on the expand button always shows a person's own direct report count; it no longer changes as deeper nodes are expanded.
- **Hide disabled accounts and hide guest accounts** now default to **On** for new web part instances.
- **Profile photo sizes increased** for clearer display at larger card sizes.
- Long job titles now wrap to two lines on org chart cards instead of truncating.

### Fixed
- `hideDisabledAccounts` and `hideGuestUsers` filters were not being applied when the Graph API data source was active.
- Horizontal and vertical tree layouts now maintain strict level alignment — all nodes at the same org depth appear in the same row/column.

### Removed
- Multi-column wrapping for managers with 8+ direct reports has been removed. All direct reports now appear in a single row (vertical) or column (horizontal), consistent with a traditional org chart layout.

---

## [1.0.0] — 2025-12-01

### Added
- Initial release.
- Employee Directory with A–Z filter, real-time search, paginated card grid, three card sizes, and per-user field visibility toggles.
- Org Chart with drill-down, vertical tree, and horizontal tree layouts.
- Person profile card with presence status, manager chain, and action buttons (Chat, Email, Call, Focus).
- Org chart toolbar: search, Find Me, layout picker, stats bar, department filter, user type filter, zoom controls.
- Four themes: Modern, Minimal, Corporate, Dark.
- PDF and PNG export.
- Demo mode with 150 / 500 / 1,000 person mock datasets.
- User preferences panel persisted to `localStorage`.
- Graph API and SharePoint Search data sources with automatic fallback.
- CI release workflow producing `smart-org-chart.sppkg`.
