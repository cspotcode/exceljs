# excel-validator

Uses Excel Desktop on Windows as a validator, via its COM API.
Opens an `.xlsx` file and reports whether Excel needed to repair or recover it.

This directory holds only the thin native shim: the bare minimum necessary to
call Excel's COM APIs.

## Requirements

- **Windows only.** Uses `Microsoft.Office.Interop.Excel` COM automation.
- **Microsoft Excel Desktop must be installed** on the machine running this tool. There is
  no headless/server mode — this launches (and quits) a real, invisible Excel process.

## Build

```shell
npm run build:excel-validator
```
