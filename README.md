# Document Generator

A simple Windows batch and VBScript-based tool for generating invoice and credit note documents.

## Features

- **Invoice Generation**

  - Customer name
  - HS Code (optional)
  - Product name
  - Quantity
  - Unit amount
  - Tax class (v1, v3, exempt)

- **Credit Note Generation**
  - Credit note number (DCN)
  - Quantity
  - Amount
  - Tax class (v1, v3, exempt)

## Installation

1. No installation required
2. Simply download the batch file
3. Make sure you have write access to `C:\dpool\in` folder (will be created if it doesn't exist)

## Usage

1. Run `input_gui.bat`
2. Choose from the main menu:
   - 1: Generate Invoice
   - 2: Generate Credit Note
   - 3: Exit

### Output Format

Files are saved in `C:\dpool\in` with timestamps in their names.

#### Invoice Format

```
r_trp "Product Name" quantity * amount tax_class
```

or with HS code:

```
r_trp "HS_Code Product Name" quantity * amount tax_class
```

#### Credit Note Format

```
r_dcn "credit_note_number"
r_trp "total" quantity * amount tax_class
```

## Requirements

- Windows operating system
- Write access to `C:\dpool\in`
- Windows Scripting Host (WSH) enabled for VBScript support

## Error Handling

- All required fields must be filled
- Invalid inputs will cancel the operation
- Program automatically returns to main menu after completion or cancellation
- Temporary VBS files are automatically cleaned up on exit
