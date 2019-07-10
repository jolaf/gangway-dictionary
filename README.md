# gangway-dictionary
Simple script to automatically generate multi-language dictionary leaflets.

**NOTE** The basic language for this package is Russian, though the software itself is not language-dependent.

## Dependencies

* Python 3.6+
* `pip3 install docxtpl pygsheets`
* [Microsoft Word](https://products.office.com/word) to edit the layout template and print the generated leaflets.

## Usage

* Adjust dictionary content in [spreadsheet](https://docs.google.com/spreadsheets/d/1kbMuGJaRR4gYTr9yaobskiENStj48m8wqhKRjlIQ0Tc), if needed.
* Adjust layout in `GangwayDict-Template.docx`, if needed.
* Run `python3 GangwayDict.py`
* Check the generated `GangwayDict-EN.docx` and other files, one per language.
