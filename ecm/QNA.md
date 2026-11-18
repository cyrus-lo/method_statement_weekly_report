# ECM Workflow Q&A

## Why does Step 1 use `subprocess`?

`subprocess` starts PowerShell, which automates the installed desktop Excel application through its COM interface.

`openpyxl` can save an Excel filter definition such as `Count (Y/N) = Y`, but it does not execute that filter. Without Excel automation, the workbook can open with both Y and N rows visible until the user reapplies the filter manually.

The PowerShell/Excel step applies the filter after the workbook is saved. Excel then hides only the N rows as a normal filter result. No rows are deleted; clearing the filter in Excel shows every row again.

For Step 1b, `openpyxl` first validates, saves, and closes the finalized workbook. The filter helper then starts one hidden Excel session, reopens `ECM_Record.xlsx` once, applies the `Count (Y/N) = Y` filter, saves, closes the workbook once, and quits Excel once. Only after that complete Excel operation succeeds is `_1b_after.xlsx` created. The later manual `Use SMM` / `Use KMS` decisions are a separate post-Step-1b Excel operation.

Without `subprocess`, the program can preserve the filter dropdown criterion, but cannot reliably open the workbook with the Y filter already active using `openpyxl` alone.
