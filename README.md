# translate_excel
This Python script uses [translators](https://pypi.org/project/translators/) package to translate an Excel file from one language to another. The script supports multisheets.

## Usage
To run the script from the command line type:

`python translate_excel.py <filename> <source langauage> <destination language>`

where `<filename>` is the filename of the Excel file, `<source language>` is the original language and `<destination language>` is the language you wish to translate the Excel to. For the full list of available languages, see [translators](https://pypi.org/project/translators/). 

To test the script, use the file `test-ファイル名.xlsx` with source language set to `ja` (Japanese) and destination language set to `en` (English).

## Notes
- The script makes one HTTP request per cell with a delay of 1 second between separate requests, so it may take a while to translate large Excel files. This number can be tweaked in the code.
- The default engine is Google Translate, which can also be changed in the code.


