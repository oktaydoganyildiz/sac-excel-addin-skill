---
name: sac-excel-addin
description: >-
  Use for SAP Analytics Cloud, add-in for Microsoft Excel (SAC Excel add-in):
  writing, converting or debugging SAP.* formulas (SAP.GETDATA, SAP.SETDATA,
  SAP.GETDIMENSIONFILTER, SAP.OVERWRITEDIMENSIONFILTER, SAP.MEMBERPROPERTY,
  SAP.GETVARIABLES, SAP.TIMEOFFSET, SAP.ASYMMETRICFILTER); migrating Analysis for
  Office (AO, AfO) workbooks, e.g. SAPGetData to SAP.GETDATA, SAPSetFilter or
  SAPSetVariable VBA; planning in Excel (SAP.SETDATA, mass data entry, Process
  Data, publish, data action and multi action triggers); VBA macros with
  SAPOFXLA.xlam (SAP Note 3594453); tables on SAC models, Datasphere analytic
  models, BW, S/4HANA or HANA Cloud live connections; and add-in setup and
  errors: Go! button disabled, empty builder panel, trusted origins,
  non-standard tenant domains, missing named ranges, high-volume mode, memory
  crashes. Not for SAC stories, Analytics Designer scripting or SAC modeling.
license: Apache-2.0
metadata:
  author: Oktay Doganyildiz
  version: "1.0.0"
  verified-against: "SAC add-in 2026.23 docs, 2026-10-29"
---

# SAP Analytics Cloud, add-in for Microsoft Excel

The add-in brings SAC and its live data sources into Excel. It is a Microsoft 365 web add-in that runs in Excel for the Web and Excel desktop on Windows and Mac. It adds an **SAP Analytics Cloud** ribbon tab and requires an SAC license.

There are two ways to consume data:
- **Tables** are crosstabs bound to one data source. You can filter, sort, rank, use hierarchies, style and plan in them.
- **Data sources for formulas** insert no table. You build free-form reports with `SAP.*` formulas against an alias (`DS_1`, `DS_2`, …).

Supported data sources: SAC analytic and planning models; Datasphere analytic models (direct or tunnel); S/4HANA Cloud queries; HANA Cloud views; S/4HANA and BW 7.50 SP14+ / BW/4HANA queries over a tunnel connection. All connections are defined in SAC. **No BPC and no on-premise HANA.**

## Which reference to load

| Task | Read |
|---|---|
| Write, debug or explain any `SAP.*` formula | `references/formulas.md` |
| Move workbooks from Analysis for Office, convert `SAPGetData`, assess VBA | `references/migration-from-analysis-office.md` |
| VBA automation of the add-in | `references/vba-api.md` |
| Plan data, versions, publish, data and multi action triggers | `references/planning.md` |
| Tables, filters, hierarchies, variables, large tables, source limits | `references/analysis-tables.md` |
| Install, tenant domains, trusted origins, permissions, errors, shortcuts | `references/setup-and-troubleshooting.md` |
| Where a fact comes from, and when it was checked | `references/sources.md` |

Load only the file(s) the task needs.

## Rules that prevent most wrong answers

1. **`SAP.GETDATA` / `SAP.SETDATA` take dimension–member *pairs* of IDs**, one pair per dimension. The measure is passed as its own pair (`"Account","Gross_Margin"`). The AO string syntax `"Dim=Member;…"` does not work. Dimensions you leave out fall back to the panel filter, then to the backend default. "All members" is not accepted.
2. **Use IDs, not descriptions.** If an ID is not unique within a dimension, the add-in picks one of the matching members at random.
3. **Formulas do not recalculate when a filter changes.** Refresh them with *Refresh Sheet Functions* (`Shift+Ctrl+Alt+F10`) or *Refresh Selected Functions* (`…+F8`).
4. **Use the user's list separator.** The documentation examples use `,`; locales such as German need `;`.
5. **Volume:** keep ≤ 5000 GETDATA/SETDATA per workbook and switch on *Optimize Data Retrieval*. Standard tables reach their memory limit at about 0.5–1 M cells. Beyond that, use high-volume mode (about 10 M cells), which has no planning, styling or comments.
6. **VBA exists, but is narrow.** `SAPOFXLA.xlam` (SAP Note 3594453) works only in Excel desktop. It can refresh, sign in or out, process, publish, revert and run planning triggers. **It cannot set filter values or variable values programmatically.** Never invent methods; check `references/vba-api.md`.
7. **AO features with no equivalent:** BPC, BW-IP planning functions and sequences, AO's VBA filter and variable setters, custom calculations on the data source, broadcasting. Say so plainly and recommend keeping those workbooks on AO.
8. **Versions change monthly.** When a fact is tied to a version, name the version (for example "since 2026.19") and point to What's New. Mark anything you could not verify as *(unverified)*.

## Safety: confirm before anything that writes data

These actions change data on the SAC tenant or discard it:
- Process Data;
- Publish Data;
- Revert Data;
- Clear Data;
- running data action or multi action triggers;
- `SAP.SETDATA` with a value;
- the VBA calls `ProcessData`, `PublishData`, `RevertData`, `ExecutePlanningTrigger`, `ClearData`.

- Before you tell a user to run one of these, or generate a macro or formula that does, state the model, the version and what will change. Proceed only if the user asked for it.
- Generated publish/trigger macros must keep an explicit end-user confirmation step.
- Admin changes, such as `Set-OrganizationAddInOverrides` in Microsoft 365 or trusted origins in SAC, are explained for an administrator to perform. Do not run them.
- Never put credentials or tokens in workbooks, formulas or VBA. Sign-in is always interactive.

## Quick formula map

| Function | Returns |
|---|---|
| `SAP.GETDATA(src, dim, mem, …)` | Value for a member combination |
| `SAP.SETDATA(src, value, dim, mem, …)` | Writes a planning value; without a value it behaves like GETDATA |
| `SAP.GETDATACOMMENT(src, display, dim, mem, …)` | Last data point comment (optionally with the value) |
| `SAP.GETDIMENSIONFILTER(src, dim, …)` | Current filter, plus a pencil icon that opens the filter dialog |
| `SAP.OVERWRITEDIMENSIONFILTER(srcs, dim, members, …)` | Replaces a dimension's filter on tables or formula sources |
| `SAP.GETDIMENSIONS(src, …)` | Lists the dimensions |
| `SAP.MEMBERPROPERTY(src, dim, member(s), property, …)` | Property values |
| `SAP.GETVARIABLES(src, …)` | Variable values |
| `SAP.GETTABLENAME(cell)` / `SAP.GETTABLESTATUS(tables, …)` | Table name / refresh status (0 refreshed … 4 pending filters) |
| `SAP.TIMEOFFSET(src, refDate, offsets, …)` | Shifted time members (hierarchies only; not for BW/S4) |
| `SAP.ASYMMETRICFILTER(tables, blocks…)` + `SAP.BLOCK(dim, mem, …)` | Asymmetric axes (not for BW/S4 Cloud) |
| `SAP.GETPLANNINGTRIGGERNAME` / `…PARAMETERS` / `…PARAMETERVALUE` | Planning trigger metadata, for data and multi actions |

Example, AO to SAC:

```
AO : =SAPGetData("DS_1";"Net Sales";"0COUNTRY=DE;0MATERIAL=M-01")
SAC: =SAP.GETDATA("DS_1","Measures","NET_SALES","0COUNTRY","DE","0MATERIAL","M-01")
```

The measure dimension ID differs by source. For example, SAC account models use `Account`. Confirm the ID with `SAP.GETDIMENSIONS("DS_1",0)`.

## Answer style

- Give the formula or the steps first, then the reason.
- Cite the help page or note when a limit, version or restriction drives the answer.
- When the question really belongs to a neighbouring product (SAC story scripting, Datasphere modeling, BW query design), say so and stop at the add-in boundary.
