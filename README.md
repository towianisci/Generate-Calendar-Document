# Generate Calendar Document

## Description

This application generates a writable calendar for a specified year in Microsoft Word format (.docx). It creates a landscape-oriented document with one page per month, featuring a table-based calendar layout where users can write notes in each day's cell. Weekends are highlighted in orange for easy identification.

The calendar includes markings for:
- US federal holidays
- Christian holidays (Easter, Good Friday, Pentecost, Mother's Day, Father's Day)
- LDS Church-specific holidays (General Conference on the first weekend in April and October, Pioneer Day)
- California school holidays (Cesar Chavez Day)
- Key LDS Church history events (First Vision, Church Organization, etc.)

If multiple events occur on the same date, they are listed below the date in smaller italic text, separated by new lines.

## Dependencies

- Python 3.x
- `python-docx` library (install via `pip install python-docx`)

The script will automatically check for the required libraries at runtime and provide installation instructions if they are missing.

## How to Compile

To compile the script into a standalone executable using PyInstaller:

1. Run the provided build script: `./build.sh`

The build script will automatically check for and install the necessary dependencies (`python-docx` and `pyinstaller`) if they are not already present.

This will generate a single executable file named `generate_calendar` (or `generate_calendar.exe` on Windows) in the `dist/` directory.

Alternatively, run PyInstaller directly:
```
~/Library/Python/3.9/bin/pyinstaller --onefile generate_calendar.py
```

## How to Use

### Running the Python Script

1. Open a terminal and navigate to the project directory.
2. Run the script with an optional year argument:
   ```
   python generate_calendar.py [year]
   ```
   - `year`: Optional integer for the year (e.g., 2025). Defaults to the current year if not provided.

3. The output will be a .docx file named `Calendar_{year}_Writable.docx` in the current directory.

### Running the Compiled Executable

1. After compiling, locate the executable in the `dist/` directory.
2. Run it with the same optional year argument:
   ```
   ./dist/generate_calendar [year]
   ```
   (On Windows: `dist\generate_calendar.exe [year]`)

3. The output file will be generated as described above.

## Output

- A Microsoft Word document (.docx) with one page per month.
- Each page contains a table calendar with dates, holidays, and space for notes.
- File name: `Calendar_{year}_Writable.docx`

## Changelog

### Version 2.0 (December 7, 2025)
- Added automatic dependency checking in the Python script; it now checks for `python-docx` at runtime and provides installation instructions if missing.
- Updated the build script to automatically check for and install required dependencies (`python-docx` and `pyinstaller`) before compiling.
- Improved build script reliability with proper path handling.

### Version 1.8 (December 7, 2025)
- Initial release with calendar generation features.

## Author

Edward L. Thomas

## Version

2.0 (Updated December 7, 2025)
