# VBA API for the SAC add-in (`SAPOFXLA.xlam`)

The SAC add-in is a JavaScript (Office web) add-in, so VBA cannot drive it natively. SAP ships a protected bridge add-in, `SAPOFXLA.xlam`, as an attachment to **SAP Note 3594453** (component `LOD-ANA-OF-XLA`; note version 3, 2026-07-08). The note's `VbaApiDocumentation.pdf` is the reference, and an S-user is needed to download it. **[N 3594453]**

Key facts from the note and its PDF:

- It works only in Excel desktop on Windows and Mac. There is no VBA in Excel for the Web. The workbook must be macro-enabled (`.xlsm`). **[N]**
- Deploy the `.xlam` on every client and add it through *VBA editor ▸ Tools ▸ References ▸ Browse*. Do not modify the `.xlam`. **[N]**
- SAP supports only the `.xlam` functions and never supports customer VBA code. Behaviour can differ between Windows and Mac. **[N]**
- Calls are messages to a browser process. If VBA is busy or paused (for example at a debugger breakpoint), messages can be lost and the call fails. **[N PDF]**
- On Windows, calls are synchronous by default. Mac supports only async calls (`SetAsync True` plus a callback sub or `WithEvents` on `APIExecutor`). For cross-platform solutions, write async code. **[N PDF]**
- The method signature pattern is `SAPOFXLA.VBAAPI.<Method> workbookName, …, [callbackSubName], [Tag]`. The result is in `SAPOFXLA.GetApiExecutor()` with `.Result` (`"Ok"`/`"Ko"`), `.ApiError`, `.Tag`, and `.IsIdle()`. **[N PDF]**

## Methods (as of the note's PDF, retrieved 2026-10-29)

| Area | Methods |
|---|---|
| Session | `SignIn` (the workbook must already hold a tenant), `SignOut`, `ShowPane`, `SwitchOfflineMode` (`OFFLINE`/`ONLINE`, for formulas) |
| Refresh | `RefreshAll`, `RefreshSheet`, `RefreshSheetFunctions`, `RefreshSelectedFunctions`, `RefreshTable`, `PauseRefresh` |
| Prompts and filters (dialogs only) | `ShowPromptsForWorkbook`, `ShowPromptsForDataSources`, `SelectDimensionFilter` (opens the dialog; no timeout) |
| Planning | `SetPlanningMode`, `ProcessData`, `PublishData`, `RevertData` (optional `AskUserConfirmation`), `ExecutePlanningTrigger`, `ClearData` |
| Comments | `EnableDataPointCommenting` |
| Runtime control | `SetAsync`, `SetTimeout` (affects VBA only; the SAC operation keeps running), `GetTimeout`, `SetDisplayProcessingDialog` (Windows only), `SetLogLevel`/`GetLogLevel`, `isWindowsOS`, `GetApiExecutor` |

**Not available:** setting filter members or variable values from code, moving dimensions, reading cell metadata, and running BW-IP planning functions. Compare AO's `SAPSetFilter`, `SAPSetVariable`, `SAPMoveDimension`, `SAPGetCellInfo` and `SAPExecutePlanningFunction`. **[N PDF vs H AO "Working with macros"]** For filters, the workaround is a cell-driven `SAP.OVERWRITEDIMENSIONFILTER` formula followed by `RefreshTable`/`RefreshAll`.

## Safety rules for the agent

- `PublishData`, `ProcessData`, `ExecutePlanningTrigger`, `RevertData` with `AskUserConfirmation:=False`, and `ClearData` change or discard planning data. Write them into a macro only when the user explicitly asks for that behaviour. State what the macro will change, and keep a confirmation step for the end user (for example a `MsgBox vbYesNo`) before publish or data-action calls.
- Never put credentials in VBA. `SignIn` uses the tenant already stored in the workbook together with the user's interactive login.

## Known field issue

Community reports describe macros freezing at `SignIn`/`SignOut` when the user is already signed in, in the same or another open workbook (2025-09, unresolved in the thread). **[Q]** Guard the call with `GetApiExecutor().IsIdle()` and avoid repeated sign-ins.

## Minimal async pattern (cross-platform)

```vba
Sub RefreshAllAsync()
    If Not SAPOFXLA.GetApiExecutor().IsIdle() Then Exit Sub
    SAPOFXLA.VBAAPI.SetAsync True
    SAPOFXLA.VBAAPI.SetTimeout 60
    SAPOFXLA.VBAAPI.RefreshAll ThisWorkbook.Name, "'" & ThisWorkbook.Name & "'!Module1.AfterRefresh"
End Sub

Public Sub AfterRefresh(res As APIExecutor)
    If res.Result = "Ko" Then MsgBox "Refresh failed: " & res.ApiError
End Sub
```
