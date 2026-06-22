# Ignition Tags Generator

A modern GUI application for converting Excel spreadsheets into Ignition OPC UA tag configurations in JSON format.

## Requirements

- Python 3.7 or higher
- See `requirements.txt` for dependencies

## Installation

1. **Clone or download the project**

2. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

### Running the Application

**Option 1: Direct Python execution**
```bash
python IgnitionTagsConverter.py
```

**Option 2: Using the executable** (if available)
```bash
IgnitionTagsConverter.exe
```

### How to Use

1. **Load UDT JSON**: Click "Select UDT JSON" and choose your UDT definition file
   - The application will automatically extract parameters from the UDT definitions

2. **Load XLSX File**: Click "Select XLSX File" and choose your Excel spreadsheet
   - Your spreadsheet should contain columns for:
     - Tag Name
     - Data Block
     - UDT Type
     - Tag History (TRUE/FALSE) - Optional, maps to Tag_Historized parameter

3. **Configure Parameters** (if needed):
   - **Driver Device Connection**: OPC UA driver device (default: `[Siemens Enhanced]`)
   - **PLCInstance**: PLC instance name (default: `PLC`)
   - **ServerName**: Ignition OPC UA Server name (default: `Ignition OPC UA Server`)
   - **Data Block Override**: Optional override for all data blocks (leave empty to use values from XLSX)
   - **Namespace**: OPC UA namespace (default: `1`)

4. **Export**: Click "📊 Export All Sheets"
   - Select where to save the output JSON file
   - The application will process all sheets (except "SCADA Signal" sheets)
   - Tags are organized by:
     - Sheet name (tab)
     - Data block name

### File Organization

The exported JSON organizes tags as follows:
```
Sheet Tab Name/
├── DataBlock1/
│   ├── Tag1 (UDT or Atomic)
│   ├── Tag2 (UDT or Atomic)
│   └── ...
├── DataBlock2/
│   └── ...
└── ...
```

## Tag History / Tag_Historized

The "Tag History" column in your Excel file controls the `Tag_Historized` parameter in the generated JSON:
- **TRUE**, "true", "1", or "yes" → `"value": "true"`
- **FALSE**, "false", "0", or other values → `"value": "false"`
- **Empty/Missing** → defaults to `"value": "true"`

This is written to the UDT instance parameters:
```json
"Tag_Historized": {
  "dataType": "String",
  "value": "true"
}
```

## Data Types Supported

### Atomic Types
- **Integers**: INT, DINT, LINT, SINT, USINT, UINT, UDINT, ULINT
- **Floating Point**: REAL, FLOAT, DOUBLE
- **Boolean**: BOOL
- **String**: STRING
- **Byte**: BYTE
- **Special**: ANALOG, DIGITAL, CALCULATED, COMM

### UDT Types
- Any custom UDT type defined in the UDT JSON file

## Example Input Files

### Excel Spreadsheet Structure
```
Tag Name      | Data Block | UDT Type       | Tag History
--------------|-----------|----------------|------------
Temperature   | DB10      | MeasurementUDT | TRUE
Pressure      | DB10      | MeasurementUDT | TRUE
Valve_Status  | DB20      | BOOL           | FALSE
```

### UDT JSON Structure
The UDT JSON should contain valid Ignition UDT definitions with parameters and tag structures.

## Troubleshooting

### "Could not auto-detect all required columns"
- Ensure your XLSX file has columns containing: "tag name", "data block", and "udt type" (case-insensitive)
- Adjust column names to match these patterns

### "Failed to load UDT JSON"
- Verify the JSON file is valid and not corrupted
- Check that the file contains proper UDT definitions

### "Failed to load XLSX"
- Ensure the file is a valid .xlsx (Excel 2007+) format
- Not compatible with .xls (older Excel format)

## Output

The exported JSON file contains:
- A `tags` array with folder structures
- Each folder organized by sheet name and data block
- Tags with OPC UA item paths for atomic types
- UDT instances with parameter values for UDT types

Example output structure:
```json
{
  "tags": [
    {
      "name": "Sheet1",
      "tagType": "Folder",
      "tags": [
        {
          "name": "DB10",
          "tagType": "Folder",
          "tags": [
            {
              "name": "Temperature",
              "typeId": "UDTS/MeasurementUDT",
              "tagType": "UdtInstance",
              "parameters": { ... }
            }
          ]
        }
      ]
    }
  ]
}
```
