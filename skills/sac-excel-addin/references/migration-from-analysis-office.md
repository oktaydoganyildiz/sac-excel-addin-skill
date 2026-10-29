# Migrating from Analysis for Office (AO) to the SAC add-in

Use this when a user has existing SAP Analysis for Microsoft Office workbooks (Standard or SAC edition) and asks whether, or how, to move them to SAP Analytics Cloud, add-in for Microsoft Excel.

Source tags: **[H]** official SAP Help, **[N]** SAP Note, **[B]** SAP-employee blog, **[Q]** SAP Community Q&A. Full links are in `sources.md`. Anything without a tag, or marked *(unverified)*, is practitioner guidance.

## 1. Product positioning (decide first)

- Both add-ins are included in the SAC license. SAP calls the SAC add-in the strategic one. AO (SAC edition) gets quarterly service packs with bug fixes only and no new features. Mainstream maintenance is planned to end in late 2029 per the AO PAM. **[B, July 2025 update]**
- The SAC add-in is a Microsoft 365 web add-in. It installs from the Microsoft Store or through central deployment and updates itself. New features ship roughly monthly. **[B]** **[H What's New]**
- AO is a COM add-in that you install on each machine and that needs Excel desktop. **[B]**

## 2. Data source fit, the first go/no-go check

| Data source | SAC add-in | AO (SAC edition) |
|---|---|---|
| SAC models (new model type) | Yes | No |
| Datasphere analytic models | Yes (direct or tunnel connection) | No |
| S/4HANA Cloud queries | Yes | Yes (not recommended) |
| HANA Cloud views | Yes | Yes (needs the HANA Analytics Adapter, not recommended) |
| BW 7.50 SP14+ and BW/4HANA queries | Yes (tunnel connection only) | Yes |
| S/4HANA (on-premise) queries | Yes (tunnel connection only) | Yes |
| BPC | **No** | Yes |
| SAP HANA (on-premise) | **No** | Yes |

Sources: **[B]** for the matrix; **[H System Requirements]** for the connection types and BW versions. All SAC add-in connections are live connections defined in SAC, and the add-in never connects to a source system directly. **[H]**

**Stop rule:** if the workbook depends on BPC, on on-premise HANA, or on BW-IP planning functions and sequences, keep it on AO. The SAC add-in has no equivalent. **[B]**

## 3. Feature gaps to check per workbook

| AO capability | SAC add-in status |
|---|---|
| VBA automation (`Application.Run "SAPSetFilter"`, `SAPSetVariable`, `SAPExecuteCommand`, callbacks…) | A separate VBA API now exists (`SAPOFXLA.xlam`, SAP Note 3594453; Windows and Mac desktop only). **It has no method that sets filter values or variable values programmatically.** It can only open the filter or prompt dialogs. See `vba-api.md`. **[N 3594453 PDF]** |
| Planning functions and sequences (BW-IP), BPC | Not available. SAC data actions and multi actions run through planning triggers instead. **[B]** **[H What's New 2026.19]** |
| Custom calculations on the data source, style sets | AO only. **[B]** In the SAC add-in you can use restricted measures (since 2025.19), currency conversion (2026.21, SAC models), a second calculation structure (2026.03), and Excel formulas in added rows and columns. **[H What's New]** |
| Very large exports | SAC add-in standard tables hit the about 4 GB WebView2/WebKit2 memory limit at roughly 0.5–1 M cells. High-volume mode reaches up to about 10 M cells but has no planning, styling or comments. **[H System Requirements]** |
| Pre-calculation and broadcasting (APOS, BIExcellence, BW broadcaster) | No add-in equivalent. **[B]** |
| Workbook storage on BW or the BI Platform | SAC add-in workbooks are saved locally, on SharePoint/OneDrive, or in SAC as an *Add-in Workbook* object (`.xlsx`; `.xlsm` when macros are present). **[H Save and open workbooks]** |
| "Convert crosstab to formula" | No one-click equivalent found in SAC add-in docs *(unverified)*. Rebuild with a *data source for formulas* and `SAP.GETDATA`. |

## 4. Formula mapping (`SAPGetData` and related → `SAP.*`)

AO formulas live in the Excel category *Analysis*. SAC add-in formulas live in the *SAP Analytics Cloud* category, under the `SAP.` namespace. **[H]**

| AO function | Closest SAC add-in function | Behaviour difference |
|---|---|---|
| `SAPGetData` | `SAP.GETDATA` | See §5. Different argument shape and different context rules. |
| `SAPSetData` | `SAP.SETDATA` | SAC writes to the SAC planning model's edit version (public edit or private). After entry you run *Process Data*, then *Publish*. |
| `SAPGetDimensionStaticFilter`, `…DynamicFilter`, `…EffectiveFilter`, `SAPListOf…Filters` | `SAP.GETDIMENSIONFILTER` | A single function returns the current filter, plus a pencil icon that opens the filter dialog. |
| `SAPSetFilter` (VBA) / `SAPSetFilterComponent` | `SAP.OVERWRITEDIMENSIONFILTER` (formula) | The SAC version is a worksheet formula, not VBA. It takes effect when the target tables refresh. |
| `SAPGetVariable`, `SAPListOfVariables` | `SAP.GETVARIABLES` | Read only. Setting a value happens in the prompt dialog (VBA: `ShowPromptsForWorkbook` / `ShowPromptsForDataSources`). |
| `SAPListOfDimensions`, `SAPGetDimensionInfo` | `SAP.GETDIMENSIONS` | |
| `SAPGetMember` (member attribute/key/text) | `SAP.MEMBERPROPERTY` | SAC returns property values by member ID. *(mapping is closest-fit, unverified as 1:1)* |
| `SAPGetSourceInfo`, `SAPGetInfoLabel`, `SAPGetWorkbookInfo` | Partly `SAP.GETTABLESTATUS`, `SAP.GETTABLENAME` | No info-label function. Last refresh time can be shown in the panel per table (2026.23). **[H]** |
| `SAPListOfMembers`, `SAPListOfMessages`, `SAPGetMeasureFilter`, `SAPGetDisplayedMeasures`, `SAPGetUniformScaling`, `SAPGetMeasureInfo`, `SAPSelectMember`, `SAPListOfDesignRules` | none | These functions are absent from the SAC add-in function list. **[H Using formulas]** |
| — | `SAP.TIMEOFFSET`, `SAP.ASYMMETRICFILTER` + `SAP.BLOCK`, `SAP.GETDATACOMMENT`, `SAP.GETPLANNINGTRIGGER*` | SAC-only functions. |

## 5. `SAPGetData` → `SAP.GETDATA`, the conversion rules

```
AO : =SAPGetData("DS_1"; "Incoming Orders"; "Region=France;Product=Services")
AO : =SAPGetData("DS_1"; "Incoming Orders"; "Region"; B10; "Product"; "Services")
SAC: =SAP.GETDATA("DS_1", "Account", "Incoming_Orders", "Region", "FR", "Product", "SERVICES")
```

1. **Measure handling:** in AO the measure is its own argument. In SAC the measure is a regular dimension/member pair, using the account or measure dimension ID followed by the member ID. **[H both]**
2. **Pairs only:** SAC accepts only `"Dim","Member"` pairs. The AO forms `"Dim=Member;…"` and `"Dim=Member"` do not exist in SAC. Split them. **[H]**
3. **IDs, not texts:** AO accepts texts or keys. SAC formulas expect IDs. If an ID is not unique within a dimension, SAC picks one of the matching members at random. Use technical IDs. **[H Using formulas]**
4. **Context rule flips:** AO can only return combinations that are part of the current navigation state (on rows or columns, inside the filter). SAC needs no table layout. Dimensions you leave out fall back to the panel filter, then to the backend default. Virtual members such as "All members" are not recognised. **[H both]** Practical rule: specify every dimension that matters.
5. **Refresh:** SAC formulas do not recalculate when a filter changes. Use *Refresh ▸ Refresh Sheet Functions* or *Refresh Selected Functions* (VBA: `RefreshSheetFunctions`, `RefreshSelectedFunctions`). **[H]**
6. **Data source alias:** SAC creates `DS_1`, `DS_2`… when you *Add Data Source for Formulas*. You can rename it (2026.11), and a table name such as `"Table1"` also works. **[H What's New]**
7. **List separator:** AO help examples use `;` and SAC help examples use `,`. Both follow the machine's list separator. **[H both]**
8. **Volume:** keep at most about 5000 `SAP.GETDATA`/`SAP.SETDATA` per workbook. Turn on *Optimize Data Retrieval*, or *Enable advanced data retrieval* (2026.21) if you still hit server limits. **[H]**
9. **Empty cells:** AO has a setting `SAPGetDataReturnsZeroIfNoValue`. For SAC, check the result with `IFERROR`/`N()` in Excel *(unverified: no documented SAC equivalent found)*.

## 6. Migration procedure (practitioner checklist)

1. **Inventory.** For each workbook record the data sources, AO formulas used, VBA calls used, planning use (BW-IP/BPC vs none) and size (cells).
2. **Triage** each workbook with §2 and §3: *move*, *move with redesign*, or *stay on AO*. Anything VBA-heavy that sets filters or variables in code needs a redesign, for example `SAP.OVERWRITEDIMENSIONFILTER` driven by cells, or a prompt dialog.
3. **Recreate the connections in SAC.** BW and S/4 need tunnel connections, and with user/password authentication the "Save this credential for all users" option must be ticked. Datasphere uses a direct or tunnel connection. **[H System Requirements]**
4. **Rebuild.** Tables: *Add Table*. Formula reports: *Add Data Source for Formulas*, then convert formulas with §5. A fast start for an existing SAC story or Data Analyzer view is *Add a table from Story* or *from Data Analyzer*. **[H]**
5. **Planning:** map BW-IP functions to SAC data actions or multi actions, then add planning triggers. **[H Planning]**
6. **Reconcile** each cell between the AO and SAC versions before switching users over. Keep AO workbooks read-only during parallel running.
7. **Permissions, least privilege:** report consumers need only *Read* on the *Add-in Workbook* object. Give Create + Update only to workbook authors; saving needs both. **[H Permissions]**
