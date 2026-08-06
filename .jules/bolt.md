# Bolt's Journal - Critical Performance Learnings

## 2025-02-18 - [Targeted Spreadsheet Opens via Master Service Item Index]
**Learning:** Google Apps Script's `DriveApp.searchFiles` and `SpreadsheetApp.openById` are heavy O(1) synchronous IO operations that take ~1-2 seconds each. Crawling ALL Drive spreadsheets containing 'Warranty' (O(N) search) and opening them to search for a contractor's assigned items causes severe login latency (and risks exceeding the 6-minute Google Apps Script timeout limit once N grows). By indexing all active warranty items via `SERVICE_MASTER_ID`, we can instantly identify the exact small subset of lots assigned to the contractor and target only their specific files.
**Action:** Always cross-reference a master index/tracking sheet to prune O(N) Drive search space to only the required files before performing expensive synchronous open/read actions.
