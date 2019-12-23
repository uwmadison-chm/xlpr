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

But it has some bugs.

### rebuild final comparison sheet

    python3 xlpr.py rebuild file.xlsx

This will rebuild the comparison sheet of an existing spreadsheet, in place.

To do it in a `bash` loop,

    for f in location/*.xlsx; do echo $f; python3 xlpr.py rebuild $f; done

### check various facts

    python3 xlpr.py check file.xlsx

This will compare the number of used rows/columns in sheets 1 and 2, and 
confirm that their question headers are pulling from the right places.

### build comparison sheets for day reconstructions

    python3 xlpr.py dr file.xlsx

Day reconstruction comparison sheets are slightly different, having 2 
additional sheets of "episodes"

## TODO

- Better automatic support with more options
- Conditional formatting "yellow out of range" support for addcols
- Support for . in out of range formatting
- Support for alphabetical ranges (A-D) instead of 1-N, and support 
  conditional formatting for them still
- Sheet to help with visual vertical comparison of text dual entry
