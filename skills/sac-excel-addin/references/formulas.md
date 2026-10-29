# SAP Analytics Cloud add-in: full formula reference

Verified against SAP Help, "Analyze and Plan with Formulas" (deliverable `29624d6020b74bf8817bb7df398e5cb6`), on 2026-10-29. The official list has **16 functions**: the 15 on the "Using formulas" page plus `SAP.GETDATACOMMENT`, which has its own page and was added in version 2025.17.

All functions are in the Excel function category **SAP Analytics Cloud**. Arguments use the machine's list separator. These examples use a comma; locales such as German use `;`. Member and dimension IDs do not depend on the user's language, so use them instead of descriptions. If a member ID is **not unique** in a dimension, the add-in picks one of the matching members at random.

**Golden rule for data formulas:** `SAP.GETDATA`, `SAP.SETDATA` and `SAP.GETDATACOMMENT` expect a dimension-ID + member-ID pair for **every dimension** of the source. A missing dimension falls back to the panel filter, then to the backend default. Virtual members such as "All members" are not recognised. After changing a filter, refresh the functions with `Refresh ▸ Refresh Sheet Functions` or `Refresh Selected Functions`. Since 2026.17 the formula panel shows a "Pending filter(s)" warning when a refresh is due.

**Insert a formula:** in Excel Web use `Insert Function ▸ Pick a category: SAP Analytics Cloud`. In Excel Desktop choose the same category under *Or select a category*, then fill in the Function Arguments dialog.

**Data source for formulas:** *Add Data Source for Formulas* adds a source without inserting a table. It gets an alias `DS_1`, `DS_2`, … You can rename the alias (2026.11), duplicate the source (2025.15), and add restricted measures to it (2026.19). A table name such as `"Table1"` is also a valid source.

---

## SAP.GETDATA
Returns the data value for a set of dimension/member combinations.

**Arguments:** `SAP.GETDATA(Table, [Dimension, Member], …)`
- **Table** — table name or data source alias, in quotes.
- **Dimension and member** — dimension ID followed by member ID, both in quotes; repeat per dimension (including the measure/account dimension).

**Example:**
```
=SAP.GETDATA("Table1", "Account", "Gross_Margin", "Product", "PD19")
```
Returns Gross_Margin for product PD19 (Coconut Water).

**Notes:** ≤ 5000 per workbook; calculated members cost more; sparsity affects performance. Works offline (shows last value). Enable **Optimize Data Retrieval** preference. See SAP Note 3323835.

---

## SAP.SETDATA
Writes a planning value. Without a value it behaves like `SAP.GETDATA` and returns the current value.

**Arguments:** `SAP.SETDATA(Table, Value, [Dimension, Member], …)`
- **Value** — number to send; must match the cell's number format. Value can be literal or a cell reference.
- **Dimension/member** — as GETDATA, one pair per dimension.

**Example:**
```
=SAP.SETDATA("Table1", F5, "Version", "public.Strategic Growth Plan", "Measure", "sqm per Workplace")
```
or literal: `=SAP.SETDATA("Table1", 650, "Version", "public.Strategic Growth Plan", "Measure", "sqm per Workplace")`

**Workflow:**
1. Entering a new value enables **Process Data** in the ribbon.
2. Process Data first submits manual table entries. A second Process Data then submits the `SAP.SETDATA` values. It covers every sheet in the workbook.
3. Run **Refresh Table** to see the values in the crosstab.
4. Choose **Publish** or **Revert**.

**Hierarchies:** a formula that names a hierarchy node is not submitted at node level. To plan on a node, leave that dimension out of the formula and set a node-level filter in the panel. A node formula can return a partial result such as "4 out of 5 changes submitted".

**Optimized planning area:** for dimensions used in the formula, the formula members define the area. For other dimensions the table filter applies. Input outside the area extends it.

**Gotchas:**
- With high-precision inputs, wrap the formula in Excel `ROUND()`. The server rounds the stored value, which leaves a permanent delta and reopens mass data entry endlessly.
- Keep at most 5000 formulas per workbook.
- Values entered offline can be processed after you reconnect.
- With a formula data source, a fixed number format applies. Excel formatting overrides it; set the format back to *General* to restore it.

---

## SAP.GETDATACOMMENT
Returns the data point comment (last comment of the thread) for a dimension/member combination; optionally the value too. Requires data point commenting enabled and an SAC model. Not usable on tables with `SAP.ASYMMETRICFILTER`.

**Arguments:** `SAP.GETDATACOMMENT(Table, Display, [Dimension, Member], …)`
- **Display:** `0` last comment only (default) · `1` value left / comment right · `2` comment left / value right · `3` value above comment · `4` comment above value.

**Example:**
```
=SAP.GETDATACOMMENT("Table1", 3, "Account", "Gross_Margin", "Product", "PD29")
```
Value in the cell, last comment in the cell below.

---

## SAP.GETDIMENSIONFILTER
Returns a dimension's current filter plus an icon (pencil) that opens the filter dialog from the sheet.

**Arguments:** `SAP.GETDIMENSIONFILTER(Table, Dimension, [Member_display], [Selection_display], [Orientation], [Separator])`
- **Member_display:** `0` ID · `1` description · `2` ID+desc (default) · `3` desc+ID · `4` full ID · `5` ID+desc in two cells · `6` desc+ID in two cells.
- **Selection_display:** `0` ID · `1` description · `2` ID+desc · `3` desc+ID · `4` display nothing (default).
- **Orientation:** `0` vertical (default) · `1` horizontal.
- **Separator:** string between ID and description.

**Example:**
```
=SAP.GETDIMENSIONFILTER("Table2","Product",3,1,0)
```
Members as desc+ID, selection with description, vertical. Pencil icon opens the dialog. Shortcut for the icon: `Shift+Ctrl+Alt+Space`.

---

## SAP.OVERWRITEDIMENSIONFILTER
Overwrites (replaces) the selected members of a dimension on one or more tables. Always returns the new filter displayed with ID. Apply by refreshing the table, then recalculating a `SAP.GETDIMENSIONFILTER` to see it.

**Arguments:** `SAP.OVERWRITEDIMENSIONFILTER(Tables, Dimension, Members, [Selection_types], [Member_display], [Separator])`
- **Tables** — names or aliases; comma-separated string, range, single alias (`DS_1`), or a mix (`"Table1, DS_3"`).
- **Dimension** — name in quotes; may append hierarchy/version after a slash.
- **Members** — single member directly; a range must be referenced from the sheet.
- **Selection_types:** `0` Self & Descendants (default) · `1` Descendants only · `2` Self & Children · `3` Children only · `4` Self only · `5` Leaf members only. (Note: `2` returns only the member and its children, ignoring extra members.)
- **Member_display:** `0` ID · `1` desc (default) · `2` ID+desc · `3` desc+ID · `4` full ID. The Members argument must use this same display.
- **Separator** — optional, between ID and description.

**Example:**
```
=SAP.OVERWRITEDIMENSIONFILTER("Table2";"Product";"Juices";2;1)
```
New Product filter = Juices (PC2) with children, entered by description.

---

## SAP.GETDIMENSIONS
Lists all dimensions of a data source.

**Arguments:** `SAP.GETDIMENSIONS(Table, [Display], [Orientation], [Sort_order])`
- **Display:** `0` ID · `1` desc · `2` ID+desc (default) · `3` desc+ID · `4` full ID.
- **Orientation:** `0` vertical (default) · `1` horizontal (ID+desc → 2 rows/columns).
- **Sort_order:** `0` none (default) · `1` ascending · `2` descending.

**Example:**
```
=SAP.GETDIMENSIONS("Table2";1;0;1)
```
Descriptions, vertical, ascending.

---

## SAP.MEMBERPROPERTY
Returns property value(s) for given dimension member(s).

**Arguments:** `SAP.MEMBERPROPERTY(Table, Dimension, DimensionMember, Property, [Orientation])`
- **Dimension** — name or ID.
- **DimensionMember** — one ID directly, one ID as a cell reference, or several IDs as a range reference.
- **Property** — one or more property descriptions or IDs (prefer ID for multi-language). Blank shown if property not found.
- **Orientation:** `0` vertical (default) · `1` horizontal.

**Example:**
```
=SAP.MEMBERPROPERTY("Table1","Customer", E10, "Description, Country (Short_Text)", 0)
```
Description + country short text for the customer whose ID is in E10. Tip: use one formula per dimension with a range selection for performance. Works offline.

---

## SAP.GETVARIABLES
Returns the variable values of a data source.

**Arguments:** `SAP.GETVARIABLES(Table, [Variables], [Variable_Display], [Value_Display], [Layout], [Orientation])`
- **Variables** — comma-separated variable IDs, or a cell/range. Optional; omit → all variables.
- **Variable_Display:** `0` ID · `1` description (default).
- **Value_Display:** `0` ID (default) · `1` description.
- **Layout:** `0` variable in a cell + all values in one cell · `1` variable + values in separate cells (default) · `2` all values in one cell, no variable · `3` values in separate cells, no variable.
- **Orientation:** `0` vertical · `1` horizontal (default).

**Example:**
```
=SAP.GETVARIABLES("Table1", "City", 1, 1, 3, 0)
```
City's values in separate cells, description, vertical, without the variable name.

**Notes:** error if no variable found; falls back to ID if value has no description; BW/S4 live variable-value descriptions can't be read in view mode.

---

## SAP.GETTABLENAME
Returns the table name of a selected cell.

**Arguments:** `SAP.GETTABLENAME(Cell)` — a cell that is part of the table.

**Example:**
```
=SAP.GETTABLENAME(A34)
```
For performance, **copy the result to another cell** and reference that in GETDATA/SETDATA rather than nesting the GETTABLENAME cell directly.

---

## SAP.GETTABLESTATUS
Returns the refresh status of tables.

**Arguments:** `SAP.GETTABLESTATUS(Tables, [Display])`
- **Tables** — quoted comma list `"Table1, Table2"` or a range `K13:K15`.
- **Display:** `0` name+status (default) · `1` status · `2` status+message · `3` name+message · `4` name+status+message.

**Status codes:** `0` Table refreshed · `1` Table not refreshed · `2` Remote connection logon error · `3` Table data source not available · `4` Table with pending filters.

**Example:**
```
=SAP.GETTABLESTATUS("Table1, Table2"; 4)
```

---

## SAP.TIMEOFFSET
Returns a time dimension member offset from a reference date. **Hierarchies only** (not flat lists); not for SAP BW / S/4HANA sources.

**Arguments:** `SAP.TIMEOFFSET(Table, ReferenceDate, OffsetValue, [Display], [Orientation], [DimensionSetting])`
- **ReferenceDate** — literal or cell reference.
- **OffsetValue** — comma-separated offsets, e.g. `"-1, 0, 1, 2"`.
- **Display:** `0` ID (default) · `1` desc · `2` ID+desc (two cells) · `3` desc+ID (two cells).
- **Orientation:** `0` vertical · `1` horizontal (default).
- **DimensionSetting** — optional; apply to a non-default dimension (e.g. `Product`).

**Example:**
```
=SAP.TIMEOFFSET("Table2","2022", "-1, 0, 1, 2")
```
→ 2021, 2022, 2023, 2024 across cells.

---

## SAP.ASYMMETRICFILTER
Applies asymmetric filters (blocks) to tables, overwriting complex filters to build asymmetric axes. **Not** for SAP BW / S/4HANA Cloud sources, and not with time-related calculated measures (YTD/QTD → empty). Disables data point commenting on that table.

**Arguments:** `SAP.ASYMMETRICFILTER(Table, Block, …)`
- **Table** — one or more names in quotes (comma list or range).
- **Block** — one or more blocks: Excel ranges (`F6:G9`) or a reference to a `SAP.BLOCK` cell (`#A10`).

Refresh the table(s) to apply. Overlapping asymmetric filters on the same roundtrip (Refresh All) error out — run each manually. Blocks appear as filters in the panel (deletable, not editable). Multiple ASYMMETRICFILTER formulas can target the same table.

**Example:**
```
=SAP.ASYMMETRICFILTER("Table1", F6:G9, #A10)
```

**Interconnected hierarchy tip:** if two blocks use connected levels (e.g. 2024 and 2024Q4), put that dimension first on the axis and expand its upper node after applying.

---

## SAP.BLOCK
Builds one or more blocks (dimension/member combinations) for `SAP.ASYMMETRICFILTER`.

**Arguments:** `SAP.BLOCK(Dimension, Member, …)` — up to five dimension/member pairs; at least one pair mandatory.
- **Dimension** — name or ID in quotes.
- **Member** — single member, list, single cell, or range. Comma is the separator, so member descriptions containing a comma (e.g. `France,Bordeaux`) can't be used.

**Example:**
```
=SAP.BLOCK("Version", "Actual", "Time", G4:G8)
```
Entered in A9 → block starts A9 (Version=Actual in B9), continues A10 (Time from G4, then G5–G8). Reference it as `#A10` in an ASYMMETRICFILTER.

---

## SAP.GETPLANNINGTRIGGERNAME
Returns the planning trigger name from a cell.

**Arguments:** `SAP.GETPLANNINGTRIGGERNAME(Cell)` — a cell that is part of the planning trigger.

---

## SAP.GETPLANNINGTRIGGERPARAMETERS
Lists all parameters of a planning trigger.

**Arguments:** `SAP.GETPLANNINGTRIGGERPARAMETERS(PlanningTrigger, [Display], [Orientation], [SortOrder])`
- **PlanningTrigger** — trigger name in quotes.
- **Display:** `0` ID · `1` desc · `2` ID+desc (default) · `3` desc+ID.
- **Orientation:** `0` vertical (default) · `1` horizontal.
- **SortOrder:** `0` none (default) · `1` ascending · `2` descending.

---

## SAP.GETPLANNINGTRIGGERPARAMETERVALUE
Lists a parameter's values plus an icon to open a dialog to change them from the sheet.

**Arguments:** `SAP.GETPLANNINGTRIGGERPARAMETERVALUE(PlanningTrigger, Parameter, [Display], [Orientation])`
- **PlanningTrigger** / **Parameter** — names in quotes.
- **Display:** `0` ID · `1` desc · `2` ID+desc (default) · `3` desc+ID.
- **Orientation:** `0` vertical (default) · `1` horizontal.

Pencil-icon shortcut: `Shift+Ctrl+Alt+Space`.
