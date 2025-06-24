# IMLoader

A WPF desktop application for merging and consolidating data from multiple Excel files (.xlsx), designed for industrial inspection data workflows.

## Features

- **Master File Selection:**
  - Upload a master Excel file (original dataset).
  - Select the relevant worksheet/tab (default: `Inspection_Task`).
- **Additional Files:**
  - Add multiple Excel files to merge into the master.
  - For each file, select the worksheet/tab to use (default: `Inspection_Task`).
- **Header-Aware Merging:**
  - Only columns present in the master are merged.
  - Data is appended to the master, starting after the last data row.
  - Skips row 2 in all files (assumed to be filter/dropdown row); data starts at row 3.
- **Unit Number Extraction:**
  - Extracts the unit/plant number from each file's name (e.g., `CT004_IM Load_PROD_20210506.xlsx` → Unit: `4`).
  - Sets column A ("Unit") for every merged row to this extracted value.
- **Sheet Selection:**
  - User can select which sheet/tab to use for each file.
- **Export:**
  - Save the merged result as a new Excel file.
- **Status Feedback:**
  - UI displays status and error messages.

## File Structure

```
IMLoader/
├── App.xaml, App.xaml.cs           # WPF application entry point
├── AssemblyInfo.cs                 # Assembly-level config
├── ExcelFileModel.cs               # Model for Excel file metadata
├── ExcelHelper.cs                  # Core Excel logic (sheet/header/unit/merge)
├── IMLoader.csproj                 # Project file (NuGet, build config)
├── IMLoader.sln                    # Solution file
├── MainWindow.xaml                 # Main WPF UI layout
├── MainWindow.xaml.cs              # Main UI logic and event handlers
├── bin/, obj/                      # Build output and intermediate files
```

## NuGet Packages

- [ClosedXML](https://www.nuget.org/packages/ClosedXML/) (v0.105.0): For reading/writing Excel files.
- Microsoft.NET.Sdk.WindowsDesktop (WPF, via SDK)

## Setup & Usage

1. **Clone the repository** and open `IMLoader.sln` in Visual Studio 2022 or later.
2. **Restore NuGet packages** (should happen automatically on build).
3. **Build and run** the project.
4. **Workflow:**
   - Click **Select Master File** and choose your original Excel file.
   - Select the relevant sheet/tab (default: `Inspection_Task`).
   - Click **Add Files to Merge** and select one or more Excel files to merge.
   - For each, select the correct sheet/tab if needed.
   - Click **Merge and Save** to export the merged file.

## Technical Notes

- **ExcelHelper.cs**
  - `MergeFiles` handles all merging logic, header mapping, and unit number injection.
  - `ExtractUnitNumberFromFileName` uses regex to extract the unit number after `CT` and leading zeros.
  - Only columns present in the master are merged; extra columns in source files are ignored.
  - Data is always appended after the last used row in the master.
- **UI**
  - Built with WPF (XAML + C# code-behind).
  - Uses `ObservableCollection<ExcelFileModel>` for file tracking and data binding.
- **Extensibility**
  - To add new data validation, transformation, or support for other file formats, extend `ExcelHelper`.
  - For more advanced UI, consider using MVVM and data templates.

## Limitations & Assumptions

- Only `.xlsx` files are supported (not `.xls`).
- Assumes row 1 is always the header, row 2 is a filter row, and data starts at row 3.
- The unit number must be present in the filename in the form `CT0*<number>` (e.g., `CT004`, `CT21`).
- Only merges data from the selected sheet/tab in each file.

## Authors & Maintainers

- Initial implementation: [Your Name or Team]
- For handoff to another AI or developer, see code comments and this README for architecture and extension points. 