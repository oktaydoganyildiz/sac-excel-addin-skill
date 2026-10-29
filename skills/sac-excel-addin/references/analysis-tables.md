# Analysing data with tables

Verified against SAP Help, "Analyzing Data" (deliverable `c637c9ff5d61457eb415ce161e38e57b`), and What's New up to version 2026.23 on 2026-10-29.

## Adding data

- **Add Table** opens *Select Data Source*:
  - My Models: SAC models.
  - Datasphere: direct or tunnel connection (tunnel since 2026.11).
  - HANA Cloud, S/4HANA Cloud, S/4HANA via tunnel, BW via tunnel.
  - SAC Content Performance sources (2026.05).
  - Since 2026.15 the dialog loads lazily and searches server-side (max 500 hits). Since 2026.23 it also browses **workspaces**.
- **From SAC:** use *Add a table from Story* (story table menu ▸ *Open Excel Add-in*) or Data Analyzer `Export ▸ Open in Excel Add-In`. The navigation state, filters and variables carry over. Not carried over:
  - user-managed time dimensions;
  - asymmetric layouts;
  - calculated dimensions;
  - story styling and added rows/columns.
  Dynamic dates become fixed values.
- **Copy / paste a table:** within a workbook, or into another workbook on the same tenant.

## Building the layout

- Use `+ Add Measures / Dimensions` for rows and columns. Drag items to reorder them. The *permanent pane* (widen the panel) lists every object and supports drag and drop.
- You can remove `Account` and `Version` only after filtering them to a single member. Since 2026.17, Datasphere analytic models get an automatic single-member filter on Version.
- **Rename Description** works for dimensions and measures; measures only in the permanent pane. Resetting the data source does not reset the renamed descriptions.
- **Table names:**
  - letters, digits, `_` and `|` only;
  - must be unique;
  - must not start with `SAC`.
  Formulas that contain the old name as text must be fixed by hand.
- **Manual edits inside a table** are overwritten on refresh. Put your own calculations in **added rows or columns** instead. They can hold Excel formulas such as XLOOKUP or IF, but you cannot plan on them. Clearing one formula cell clears the whole line.

## Filtering

- **Filter types:** single values, ranges (*Between* / *Not Between*), and *exclude* vs *deselect*. Exclude also removes the value from node totals.
- **Context menu:** `Filter Members` / `Exclude Members` works across dimensions (2026.05). Not available on header, property or data cells.
- **Flat presentation** loads at most **2000 members**. Search to reach the rest.
- **Paste Members** adds them as *Self Only*. You can tick *Paste as Excluded Members*.
- **Removing a filtered dimension** from rows/columns keeps its filter as a background filter.
- **Filter by Measure** (ribbon):
  - *Rank* (Top/Bottom N).
  - *Filter by Values* (Between, >, ≥, <, ≤). Conditions combine with OR. SAC and Datasphere sources only. Not available in mass data entry mode.

## Hierarchies

- Expand and collapse with the arrows, the ribbon, or `Shift+Ctrl+Alt+→/←`. *Compact Display* shows the universal display hierarchy (BW, S/4).
- **Node filter options:** Self & Descendants, Descendants Only, Self & Children, Children Only, Self Only, Leaf Members Only, Exclude.
  **BW queries** only offer *Self & Descendants* for nodes and *Self Only* for leaves. Selecting a link node does not select the other link nodes.
  **S/4HANA Cloud queries:** you cannot select single hierarchy nodes in the filter dialog.
- Custom member order works in flat mode and, since 2026.15, in hierarchy view. Switching back to flat mode drops the custom order.

## Sorting, totals, formats

- **Sorting:** member sort per dimension, plus *one* dominant value sort per table. Custom order for BW flat dimensions since 2026.23.
- **Totals:**
  - *Show Totals*, placed above or below.
  - Zero/null suppression cannot be combined with Unbooked Data.
- **Number format:** scale (k/m/bn), decimals, units in cells/rows/columns, *Show Sign As*. These apply only while **Apply Default Format** (Styling tab) is on. Turn it off to use native Excel formats.
- **Styling:**
  - conditional formats: Good/Warning/Critical thresholds as `SACThreshold…` cell styles;
  - table formats by member or header scope;
  - hierarchy styles: Uniform, Level-Based, Alternating;
  - *Keep Cells Height/Width*.
- **Restricted measures** (2025.19):
  - built by excluding members; *Constant Selection* since 2026.03;
  - work in tables and formula data sources, and you can plan on them.
  A restricted measure that serves as the base of another one cannot be deleted.
- **Calculations:**
  - a second calculation structure, *Structure* (2026.03);
  - currency conversion for SAC models whose model supports conversion (2026.21).

## Large tables

- **Standard mode** handles about 0.5–1 M cells. That is where the ~4 GB WebView2/WebKit2 memory limit is reached (System Requirements page).
- **High-volume mode** (table context menu) handles up to about 10 M cells. Tables are always added in standard mode and switched afterwards. You can use Pause Refresh to do both in one step.
  **Not supported:**
  - planning, including data actions and optimized planning area;
  - the Styling and Number Format tabs, conditional formatting;
  - data point comments, added rows/columns, backend custom texts.
  **Partly supported:**
  - totals;
  - units shown only in cells;
  - values shown unscaled;
  - *Show Repetitive Members* is off by default.
- The backend uses the HANA MDS default of **1,000,000 result records**. This is not the same number as the visible cells (SAP Note 2770570).
- If a table's start position plus its layout exceeds 1,048,576 × 16,384 cells, a dialog offers to crop the table (2026.09).

## Refresh and pause

- **Refresh options:**
  - Refresh All;
  - Refresh Sheet (2026.15);
  - Refresh Sheet Functions;
  - Refresh Selected Functions;
  - Clear Data, which strips data before sharing.
- **Pause refresh** (`Shift+Ctrl+Alt+F7`) batches navigation steps. If you navigate while a step is still running, the new step is dropped with a message.
- Refresh status icons: blue means refreshed, yellow means not refreshed yet, red means sign-in failed.
- **Workbook Refresh** preference: manual (default) or automatic on sign-in.

## Variables (prompts)

- The *Set Variables* dialog opens from the ribbon, or since 2026.23 from the table and formula context menus. Viewers can change values too.
- **Workbook Variable Handling** (2026.23):
  - Default: save values and do not prompt;
  - Save values and prompt;
  - Save values and do not prompt;
  - Do not save values and always prompt. With this option, sources stay unrefreshed until the user goes through Prompts.
- **Supported variable types:**
  - SAC models: number, dimension member and currency variables.
  - Datasphere: input parameters, date parameters, story-filter currency variables. Value help, multiple values and ranges need *Multiple Entries* on the story filter.
  - BW: hierarchy, hierarchy-node, text and formula variables are supported.
  - S/4HANA Cloud: a **mandatory** hierarchy-node variable makes the query unusable.

## Source-specific limits

| Source | Limits |
|---|---|
| BW query | Interval filters defined in the query are shown and can be deleted, but not changed. Compounding is applied automatically. No switch between short/medium/long texts. Only the current version of a versioned hierarchy is shown. |
| S/4HANA (tunnel) query | Same as BW, plus: conditional formatting from the query is not shown and you cannot define your own; link nodes misbehave in the member selector. |
| S/4HANA Cloud query | Filters with ranges or excluded values are not displayed; remove them and recreate them in the add-in. No single-node selection. Hierarchy-node variable limits (see above). |
| Datasphere | Consumption of analytical-dataset views is deprecated; use analytic models. Restricted measures changed in Datasphere keep their old definition in Excel. The SELFDWC connection is hidden by default. |

## Data Access Language (SAC profile)

| Connection | Data source name | Metadata | Member texts |
|---|---|---|---|
| SAC model | yes | yes | no |
| Datasphere | no | no | yes |
| S/4HANA Cloud | yes | yes | yes |

## Commenting

- **Data point comments** (2025.15):
  - on cells of tables based on SAC models and on valid `SAP.GETDATA`/`SAP.SETDATA` cells;
  - synced with SAC in both directions;
  - edited only in the *Data Point Comment* tab, not in the sheet.
- **Who can do what:** viewers can add comments when the designer enables commenting (2026.09). Only designers can delete them.
- **Visibility:** comments appear only for users with the same data context.
- **Not available:**
  - on protected sheets;
  - on tables with `SAP.ASYMMETRICFILTER`.
- **Dimension comments** (2026.11): a comment column for row members of SAC models.

## Workbook panel

- **Preferences:**
  - Refresh;
  - Optimize Data Retrieval: Default/disabled, Enable optimized, Enable advanced data retrieval, Disable;
  - Mode for Data Formulas (online/offline);
  - Workbook Variable Handling;
  - Workbook Permissions;
  - Performance Profiling, which writes a *Performance Profiling (SAC)* sheet.
- **Components tab:**
  - lists every add-in object, grouped by sheet or by data source;
  - shows missing named ranges, which you can restore or delete;
  - has the *Table Title in Sheet* option.
