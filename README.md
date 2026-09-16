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

## Dashboard configuration

The workbook uses two **workbook-scoped named cells** so the Dashboard layout can be changed without modifying the VBA code.

### `TG_FileName`

Contains the base name of generated files.

Example:

`Traçabilité Torquage Mécanique`

Do not include the OF number. The macro automatically appends a space followed by the 12-digit OF.

Generated example:

`Traçabilité Torquage Mécanique 000013969917.xlsx`

### `TG_OutputFolder`

Contains the destination directory for generated files.

Example:

`C:\Users\User11\Gamme de Traçabilité DF`

If the directory does not exist, the macro attempts to create it, including missing parent folders.

## Run macro

Assign the Dashboard button to:

`GenerateTraceabilityFiles`

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

This means the generator workbook can present only the Dashboard/source interface while the generated traceability files still expose their normal working sheets.

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

- matching `OFs!B:AC` rows are copied as values to `TRACEABILITY!A3:AB...`
- the 12-digit OF is written to `INFORMATIONS!B3`
- the generated document title is written to `INFORMATIONS!B1`
- the four copied sheets are visible in the generated workbook
- the workbook is saved as `.xlsx` without VBA macros
- an existing output file with the same name is replaced automatically

## VBA modules

### `modTG_Config.bas`

Central constants for sheet names, source/destination layout, Dashboard named ranges, and output extension.

### `modTG_Helpers.bas`

Validation and helper functions, including:

- Dashboard/named-cell validation
- OF grouping
- 12-digit OF formatting
- folder creation
- filename cleaning
- path construction
- replacement of existing files
- clearing old traceability values

### `modTG_Generator.bas`

Main generation workflow and per-OF workbook creation, including safe handling of hidden template sheets.

## Versioning

The `.bas` files in this repository are the source-controlled VBA modules. Changes should be made through versioned commits so previous working versions can always be recovered.
