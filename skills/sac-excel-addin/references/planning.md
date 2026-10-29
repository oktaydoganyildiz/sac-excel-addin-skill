# Planning in Excel with the SAC add-in

Verified against SAP Help, "Planning" (deliverable `9bafb6f9d4034781b8040fc29a2cd99f`), and What's New up to version 2026.23 on 2026-10-29. Planning works only on **SAC planning models**, including models stored in SAP Datasphere (seamless planning). BW-IP and BPC planning are not supported (see `migration-from-analysis-office.md`).

## Write-back safety (read first)

Process Data writes to the edit version. **Publish** makes the data visible to everyone. Data action and multi action triggers change data on the server. Before you tell the user to run any of these, or write VBA that does, state what will change (model, version, scope). Do it only when the user asked for it.

## Basics

- **Mass data entry:** type several values, then choose `Process Data` (`Shift+Ctrl+Alt+Enter`). You plan on **one model at a time**. Process Data covers every sheet. Manual entries go first, and a second Process Data submits `SAP.SETDATA` values.
- **Versions:** *public* (all authorised users), *private* (creator only), *shared* (a private version shared with others). Editing a public version opens **public edit mode**. Others see the changes only after **Publish Data**. **Revert Data** discards the changes. Private versions can only be published in SAC.
- **Cell colours:** black font means input-ready (`SACDataCell`). Grey means read-only (`SACReadOnlyDataCell`). Excel still accepts typing in grey cells, but the value is not planned. Since 2025.23, data cells of models that are not enabled for planning also use black.
- **Hierarchy nodes in tables:** you can enter a value on a node. It is spread proportionally to the unchanged leaf members. Formulas behave differently (see `formulas.md` › SAP.SETDATA).
- **Input checks:**
  - Non-decimal input opens an error panel, and Process Data stays disabled until you fix it (redesigned in 2026.23).
  - Pasted values cannot skip cells in a row.
  - Up to 15 digits are allowed before the decimal separator. Excel itself keeps only 15 significant digits.
- **Excel formulas inside a planning table:**
  - A change counts only if it still differs after rounding to the cell's decimals.
  - Deleting a formula takes two steps: first the formula, then the value.
  - Changing the table layout, for example adding a dimension, removes the formulas.
- **Rules applied from SAC:** Data Access Control and validation rules (2025.17). Data disaggregation settings: based on data locking, validation rules and DAC (2026.11).
- **Data locking:** locked cells are **not marked** in Excel. You may be able to process them, but Publish is blocked.
- **Model changed in SAC while you work:** the add-in shows a version-mismatch message. Refresh.

## Enabling planning and the planning area

- Use `Planning ▸ Planning Enabled` in the table's **Planning** section of the builder panel (moved there in 2026.09), or `Ctrl+Alt+I`. Planning is on by default for designers and off for viewers.
- **Optimize Planning Area** builds a minimal planning area from the table context, which speeds things up. Members you add later are not input-ready until you switch the setting off and on again. The planning area is limited by the model preference and a hard cap of 100 M records. Since 2026.17, unfiltered dimensions are ignored when a default planning area exists.
- **Advanced Settings for Booked/Unbooked Data** (2026.09) has three options: *No Restrictions*, *Hide Unplannable Unbooked Data*, *Show Invalid Booked Data Only*.
- **Unbooked cells** show `-` (value 0). Turn them on with the dimension context menu `Unbooked Data` or the filter option `Show Unbooked Members`. Switching on propagates to the dimensions *below* in the panel. Switching off removes unbooked data for that dimension and the dimensions *above* it.
- **Two objects on one model:** a second table or formula data source on the same model stays plannable, and you can still process data. Publish and Revert are blocked while planning is disabled on the first table.
- **Revert** resets the version to public.

## Planning triggers

Both trigger types need **4 empty cells**. They are configured in the pane and run by clicking the blue trigger cells. The functions `SAP.GETPLANNINGTRIGGERNAME`, `SAP.GETPLANNINGTRIGGERPARAMETERS` and `SAP.GETPLANNINGTRIGGERPARAMETERVALUE` work with both types.

### Data action trigger (`Planning Trigger ▸ Add Data Action Trigger`)

- The SAC equivalent is the story's "data action starter".
- **Parameters:**
  - Dimension parameters take a *Fixed Value* or a *Table Filter*. Since 2025.23 a Table Filter also works for non-target versions and flat hierarchies, and since 2026.07 for "All Members".
  - Number parameters take a *Fixed Value* or a *Cell Reference*.
  - The SAC *Default Value* is preselected (2026.03) and cannot be changed in Excel.
- **Settings:**
  - *Refresh table on target model automatically* is on by default. It cannot run while the target table is in mass data entry mode.
  - *Planning Area Configuration*: Automatic / Empty (No Version Data) / Recommended / All Version Data.
  - *Publish to Target Version*: Do not Publish / Publish and fail if there are warnings / Publish and ignore warnings.
- A cross-model copy step with a `Source_Dimension` parameter can copy data from an analytic model to a planning model.
- If the data action's parameters change in SAC, reload it in Excel. A cardinality change (Any to Single) means the values must be defined again.
- Check the status in the SAC **Data Action Monitor**. The check icon in the trigger header validates the configuration.

### Multi action trigger (`Planning Trigger ▸ Add Multi Action Trigger`, since 2026.19)

- The SAC equivalent is the "multi action starter".
- Parameter types: Member, Number, String (with or without predefined values), Date, Time.
- **Validate** checks the setup against the backend. Progress shows in the SAC **Job Monitor**.
- The table refreshes automatically when the multi action changes its data.

### Triggers via VBA

`SAPOFXLA.VBAAPI.ExecutePlanningTrigger workbookName, "PlanningTrigger1"`; see `vba-api.md`.

## Calendar and stories

- You can link workbooks to SAC calendar tasks. **[B: SAP blog, July 2025]**
- *Add a table from Story* (2025.17) carries over the planning rights of viewers who are allowed to plan in the story (2026.19). It is not available for models with a user-managed (flexible) date dimension (2026.11).
- *Add a table from Data Analyzer* (2026.11) opens with SAC permissions, and planning is disabled for viewers (2026.15).
