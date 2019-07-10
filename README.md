# gangway-dictionary
Simple script to automatically generate multi-language dictionary leaflets.

**NOTE** The basic language for this package is Russian, though the software itself is not language-dependent.

## Dependencies

* Python 3.6+
* `pip3 install docxtpl pygsheets`
* [Microsoft Word](https://products.office.com/word) to edit the layout template and print the generated leaflets.

## Usage

* Create authentication token for Google API, as instructed [here](https://pygsheets.readthedocs.io/en/stable/authorization.html).
* Adjust dictionary content in [spreadsheet](https://docs.google.com/spreadsheets/d/1kbMuGJaRR4gYTr9yaobskiENStj48m8wqhKRjlIQ0Tc), if needed.
* Adjust layout in `GangwayDict-Template.docx`, if needed.
* Run `python3 GangwayDict.py` from console or terminal of your operating system.
* At first run, authorize in Google API by browsing to the URL provided on screen.
* Check the generated `GangwayDict-EN.docx` and other files, one per language.
