# xlpr: the Excel Parity Robot

Excel helper to generate data entry and comparison worksheets.

## Requirements

    pip3 install -r requirements.txt

## Usage

    python3 xlpr.py manual NAME_OF_FILE NUM_QUESTIONS NUM_PARTICIPANTS
    python3 xlpr.py auto INPUT_XLSX NUM_PARTICIPANTS OUTPUT_DIR

## Examples

    python3 xlpr.py PANAS_NOW 20 350
    python3 xlpr.py qrre_input.xlsx 350 ./output/

## TODO

- Support for alphabetical ranges (A-D) instead of 1-N, and support 
  conditional formatting for them still
- Sheet to help with visual vertical comparison of text dual entry
