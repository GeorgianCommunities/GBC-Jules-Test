# Palette's UX and Accessibility Learning Journal

This journal tracks critical UX and accessibility learnings, patterns, and decisions made in the Georgian Build Connect ecosystem.

## 2024-08-06 - Initial Discovery
**Learning:** Google Apps Script web application portals often suffer from missing native keyboard focus indicators and non-semantic interactive components (like spans and divs) used as buttons (e.g., modals, AI widgets, and sortable headers). This breaks keyboard navigation completely for assistive technology users.
**Action:** Always ensure interactive elements (even custom close buttons and tab headers) use semantic tags, correct ARIA roles, have explicit focus-visible styles, and support keyboard keypress events (Enter/Space).
