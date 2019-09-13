# xlpr: the Excel Parity Robot

Excel helper to generate data entry and comparison worksheets.

## Requirements

    pip3 install -r requirements.txt

## Usage

    python3 xlpr.py manual NAME_OF_FILE NUM_QUESTIONS NUM_PARTICIPANTS
    python3 xlpr.py auto INPUT_XLSX NUM_PARTICIPANTS OUTPUT_DIR

## Examples

### manual mode

    python3 xlpr.py manual PANAS_NOW 20 350

### automatic mode

To build a bunch of spreadsheets like AFCHRON did,

    python3 xlpr.py auto qrre_input.xlsx 350 ./output/

The automatic mode assumes certain things about the input excel spreadsheet,
and could certainly be improved.

### add columns mode

    python3 xlpr.py addcols existing_thing.xlsx 17

This will add 17 fresh question columns to an existing spreadsheet, in place.

## TODO

- Better automatic support with more options
- Conditional formatting "yellow out of range" support for addcols
- Support for alphabetical ranges (A-D) instead of 1-N, and support 
  conditional formatting for them still
- Sheet to help with visual vertical comparison of text dual entry
