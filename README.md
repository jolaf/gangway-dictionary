# gangway-dictionary
Simple script to automatically generate multi-language dictionary leaflets.

The current version of generated leaflets is available right away in [DOCX](docx) and [PDF](pdf) formats.

**NOTE!** The basic language for this dictionary is **Russian**, though the software itself is not language-dependent.

## System requirements

Generation of [DOCX leaflets](docx) from the [DOCX layout template](GangwayDict-Template.docx?raw=true) and dictionary content stored in the [Google Spreadsheet](https://docs.google.com/spreadsheets/d/1kbMuGJaRR4gYTr9yaobskiENStj48m8wqhKRjlIQ0Tc) can be performed on any platform (Linux, MacOS, Windows).

To do it you will need:

* [Python 3.6](https://www.python.org/downloads/) or newer.
* `pip3 install docxtpl pygsheets`
* Internet access &ndash; to install `pip` modules and to get the dictionary content from the [Google Spreadsheet](https://docs.google.com/spreadsheets/d/1kbMuGJaRR4gYTr9yaobskiENStj48m8wqhKRjlIQ0Tc).
* [Adobe Reader](https://get.adobe.com/reader) to view and print generated [PDF leaflets](pdf).

If you want to

* View and print generated [DOCX leaflets](docx);
* Adjust the [DOCX layout template](GangwayDict-Template.docx?raw=true);
* Generate [PDF leaflets](pdf),

then you will need:

* [Microsoft Windows](https://www.microsoft.com/windows) and [Microsoft Word](https://products.office.com/word) to edit the [DOCX layout template](GangwayDict-Template.docx?raw=true), to print generated [DOCX leaflets](docx) and to generate [PDF leaflets](pdf).
* `pip3 install comtypes`

## Usage

* Create authentication token for Google API, as instructed [here](https://pygsheets.readthedocs.io/en/stable/authorization.html).
* Adjust dictionary content in the [Google Spreadsheet](https://docs.google.com/spreadsheets/d/1kbMuGJaRR4gYTr9yaobskiENStj48m8wqhKRjlIQ0Tc), if needed.
* Adjust layout in the [DOCX layout template](GangwayDict-Template.docx?raw=true), if needed.
* Run `python3 GangwayDict.py` from console or terminal of your operating system.
* At first run, authorize in Google API by browsing to the URL provided on screen.
* Check the generated [DOCX](docx) and [PDF](pdf) leaflets, one per language per format.
