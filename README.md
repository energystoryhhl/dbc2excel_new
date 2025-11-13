# DBC2Excel

A PyQt5-based GUI application for converting CAN database files between different formats, with a primary focus on DBC to Excel conversion and vice versa.

## Features

- **Multi-format Support**: Convert between various CAN database formats including:
  - DBC (Vector CANdb++)
  - Excel (XLS/XLSX)
  - DBF
  - KCD
  - ARXML
  - JSON
  - SYM

- **Interactive File Viewer**: Browse and visualize CAN message structures with a tree-based interface showing:
  - Frame ID (hexadecimal)
  - Frame Name
  - Cycle Time
  - Signal details (Name, Start Bit, Length, Min/Max values, etc.)

- **Signal Information Display**: View comprehensive signal properties including:
  - Start Bit
  - Signal Length
  - Maximum and Minimum values
  - Default values
  - Endianness (Little/Big Endian)
  - Factor and Offset
  - Signed/Unsigned status
  - Value tables

- **User-Friendly GUI**: 
  - Easy file opening and conversion
  - Expandable/collapsible tree view
  - Real-time conversion status updates
  - Quick access to converted files

## Requirements

- Python 3.x
- PyQt5
- canmatrix

## Installation

1. Clone the repository:
```bash
git clone https://github.com/energystoryhhl/dbc2excel_new.git
cd dbc2excel_new
```

2. Install required dependencies:
```bash
pip install PyQt5 canmatrix
```

## Usage

### Running the Application

Navigate to the `src` directory and run:

```bash
cd src
python main_py.py
```

### How to Use

1. **Open a File**: 
   - Click the "Open" button or use `Ctrl+O`
   - Select a supported CAN database file (DBC, Excel, etc.)

2. **View File Contents**:
   - The tree view displays all frames and their signals
   - Use `Ctrl+T` to expand all items
   - Use `Ctrl+Y` to collapse all items

3. **Convert File**:
   - Click the "Convert" button
   - Choose the output file format and location
   - Wait for the conversion to complete

4. **Open Converted File**:
   - After successful conversion, click "Open Converted File" to view the output

### Menu Shortcuts

- `Ctrl+O` - Open file
- `Ctrl+Q` - Close file
- `Ctrl+T` - Expand all
- `Ctrl+Y` - Collapse all

## Building Executable

The project includes a PyInstaller spec file (`main_py.spec`) for creating a standalone executable:

```bash
pyinstaller main_py.spec
```

## Project Structure

```
dbc2excel_new/
├── src/
│   ├── main_py.py          # Main application logic
│   ├── main_wd.py          # UI components (generated from Qt Designer)
│   ├── main_py.spec        # PyInstaller configuration
│   ├── icon/               # Application icons
│   └── tests/              # Test files
│       ├── test.dbc        # Sample DBC file
│       └── test.xlsx       # Sample Excel file
└── README.md
```

## File Format Details

### DBC Files
CAN database files that define the structure of CAN messages, signals, and network nodes.

### Excel Files
Spreadsheet format for viewing and editing CAN database information in a tabular format.

## Dependencies

- **PyQt5**: GUI framework
- **canmatrix**: Library for handling various CAN database formats

## License

This project is available for use and modification.

## Author

energystoryhhl

## Contributing

Contributions are welcome! Please feel free to submit issues or pull requests.

## Support

For issues, questions, or suggestions, please open an issue on the GitHub repository.
