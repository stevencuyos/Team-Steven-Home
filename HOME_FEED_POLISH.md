# Home Feed Polish

## Scope
Presentation-only improvements for the Google Apps Script UI.

## What changed
- Refined task-card elevation, hover, focus, and active states.
- Improved horizontal task rails and scroll behavior.
- Improved category tiles and touch targets.
- Added small-screen sizing for Home task cards and tabs.
- Added reduced-motion support.
- Preserved existing GAS runtime, server calls, caching, and rendering logic.

## Validation
This change intentionally modifies CSS only. No `.gs` files, `google.script.run` calls, or server-side behavior were changed.

## GAS deployment note
Copy the updated `StyleResponsive.html` contents into the matching Apps Script HTML file. No new library or build step is required.
