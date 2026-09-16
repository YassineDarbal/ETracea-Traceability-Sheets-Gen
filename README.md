# ETracea Traceability Sheets Gen

VBA source code for the Excel traceability-sheet generator.

## Purpose

The tool generates one `.xlsx` traceability workbook per unique OF found in the source sheet. Rows belonging to the same OF are grouped into the same generated workbook.

## Workbook requirements

Required worksheets:

- `Dashboard`
- `OFs`
- `SCREEN VIEW`
- `INFORMATIONS`
- `TRACEABILITY`
- `HISTORY`

The `Dashboard` sheet can also be created automatically by the VBA macro described below.

## Build the Dashboard automatically

Import `modTG_Dashboard.bas`, then run:

`BuildDashboard`

The macro rebuilds the Dashboard with:

- an ETracea header and subtitle;
- a configuration card;
- a file-name input;
- an output-folder input;
- a `PARCOURIR` button using the Windows folder picker;
- a primary `GÉNÉRER LES FICHIERS` button;
- a secondary `VIDER LES OFs` button;
- a status card updated by the generation workflow.

If `TG_FileName` and `TG_OutputFolder` already exist, their current values are preserved before the Dashboard is rebuilt.

## Dashboard configuration

The workbook uses two **workbook-scoped named cells** so the Dashboard layout remains decoupled from the generation logic.

### `TG_FileName`

Created by `BuildDashboard` and linked to `Dashboard!F8`.

Contains the base name of generated files.

Example:

`Traçabilité Torquage Mécanique`

Do not include the OF number. The macro automatically appends a space followed by the 12-digit OF.

Generated example:

`Traçabilité Torquage Mécanique 000013969917.xlsx`

### `TG_OutputFolder`

Created by `BuildDashboard` and linked to `Dashboard!F10`.

Contains the destination directory for generated files.

The `PARCOURIR` button runs:

`SelectOutputFolder`

and fills this parameter through `Application.FileDialog(msoFileDialogFolderPicker)`.

If the directory does not exist when generation starts, the generator attempts to create it, including missing parent folders.

## Run macro

The Dashboard primary button is assigned automatically to:

`GenerateTraceabilityFiles`

The status card shows preparation, live per-OF progress (`current / total` plus the current OF number), success, warning, and error messages.

## Clear loaded OF data

The Dashboard secondary button is assigned automatically to:

`ClearOFs`

This action:

- keeps rows 1 and 2 of `OFs` untouched;
- deletes loaded rows starting from row 3 through the last used row;
- asks for confirmation before deleting;
- updates the Dashboard status after the sheet is cleared;
- restores the previous Excel `ScreenUpdating` state after the operation.

## Hidden template sheets

The template sheets can remain hidden in the generator workbook:

- `SCREEN VIEW`
- `INFORMATIONS`
- `TRACEABILITY`
- `HISTORY`

Using `xlSheetVeryHidden` is recommended for the generator workbook so users cannot unhide these sheets through the normal Excel interface.

During generation, the macro:

1. stores the original visibility state of all four template sheets;
2. temporarily makes them visible;
3. copies them into the new workbook;
4. immediately restores the original visibility state in the generator workbook;
5. forces the copied sheets to `xlSheetVisible` in the generated `.xlsx` file.

The original visibility state is also restored if an error occurs during the copy operation.

## Source data

- Source sheet: `OFs`
- First source row: `3`
- OF column: `A`
- Exported columns: `B:AC`
- OF numbers are normalized to 12 digits.

## Generated workbook

Each generated workbook contains:

- `SCREEN VIEW`
- `INFORMATIONS`
- `TRACEABILITY`
- `HISTORY`

For each OF:

- matching `OFs!B:AC` rows are copied as values to `TRACEABILITY!A3:AB...`;
- the 12-digit OF is written to `INFORMATIONS!B3`;
- the generated document title is written to `INFORMATIONS!B1`;
- the four copied sheets are visible in the generated workbook;
- the workbook is saved as `.xlsx` without VBA macros;
- an existing output file with the same name is replaced automatically.

## VBA modules

### `modTG_Config.bas`

Central constants for sheet names, source/destination layout, Dashboard named ranges, and output extension.

### `modTG_Helpers.bas`

Validation and helper functions, including:

- Dashboard/named-cell validation;
- OF grouping;
- 12-digit OF formatting;
- folder creation;
- filename cleaning;
- path construction;
- replacement of existing files;
- clearing old traceability values.

### `modTG_Dashboard.bas`

Dashboard UI generation and interaction, including:

- `BuildDashboard`;
- `SelectOutputFolder`;
- `ClearOFs`;
- `UpdateDashboardStatus`;
- automatic creation of the two workbook-scoped configuration names;
- automatic button creation and macro assignment.

### `modTG_Generator.bas`

Main generation workflow and per-OF workbook creation, including safe handling of hidden template sheets and live Dashboard status updates.

## Versioning

The `.bas` files in this repository are the source-controlled VBA modules. Changes should be made through versioned commits so previous working versions can always be recovered.
