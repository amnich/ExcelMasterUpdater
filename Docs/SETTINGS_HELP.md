# Settings - Metadata Columns Guide

The **Settings** view in Excel Master Updater includes a **Metadata Columns** (Kolumny metadanych) tab that allows you to automatically inject dynamic operational metadata into the updated Master Base file.

## Overview

Metadata tokens are special variables that the Excel Master Updater automatically replaces with actual processing context information (such as timestamps, file paths, or user names) whenever a row is inserted or updated in the base file. 

By mapping a column in your Base Excel file to a specific **Metadata Token**, you ensure that this column is always kept up to date with the latest operation details behind the scenes.

## Configuration Fields

When adding or modifying a metadata column rule, you must configure three fields:

1. **Column in base (Kolumna w bazie)**: The exact name of the column header in your Master Base Excel file where the metadata should be written. If the column does not exist, you should create it in the Excel file.
2. **Metadata token (Token metadanych)**: The predefined application variable to inject into the cell.
3. **Format (optional) (Format (opcjonalny))**: Allows formatting the injected value. This is primarily used for formatting dates and times using standard .NET formatting strings.

## Available Metadata Tokens

The following tokens are supported by the engine and can be selected from the dropdown:

- `ChangeDate`: The exact date and timestamp when the row was updated or inserted.
  - *Example Format*: `yyyy-MM-dd HH:mm` or `dd/MM/yyyy HH:mm:ss`.
- `SourceFileFullPath`: The complete absolute path to the incoming update file from which the data was merged.
- `SourceFileName`: The file name (with extension, without the folder path) of the incoming update file.
- `CurrentUser`: The Windows username (Environment.UserName) of the person who performed the update.
- `SourceRowNumber`: The exact row number in the incoming update file where this data originated. Useful for auditing and tracing back data to the source spreadsheet.
- `ImportBatchId`: A unique GUID identifier generated for the current application run/batch, allowing you to filter and group all rows modified during a single operation session.

## How to Configure

1. Open the **Settings (Ustawienia)** window from the top right menu of the main application.
2. Navigate to the **Metadata Columns (Kolumny metadanych)** tab.
3. Click **+ Add column (+ Dodaj kolumnę)** to define a new metadata rule.
4. Specify the Base Column Header name and select the appropriate Metadata Token from the list.
5. (Optional) Provide a formatting string if using the `ChangeDate` token.
6. Click **Save settings (Zapisz ustawienia)**.

Once configured, the engine will automatically intercept these columns during the mapping process and populate them in the base file during the *"Apply to base (Zastosuj do bazy)"* phase. You do not need to manually map these columns in the main mapping window.
