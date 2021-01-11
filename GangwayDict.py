#!/usr/bin/env python3

from os.path import abspath
from traceback import format_exc
from typing import Any, List, Optional, Sequence, Tuple

try:
    from docxtpl import DocxTemplate # type: ignore
except ImportError as ex:
    raise ImportError(f"{type(ex).__name__}: {ex}\n\nPlease install docxtpl v0.6.3 or later: https://pypi.org/project/docxtpl/\n")

try:
    from pygsheets import authorize # type: ignore
    from pygsheets.worksheet import Worksheet # type: ignore
except ImportError as ex:
    raise ImportError(f"{type(ex).__name__}: {ex}\n\nPlease install pygsheets v2.0.1 or later: https://pypi.org/project/pygsheets/\n")

try:
    from comtypes import client as comClient, COMError # type: ignore
    def comErrorStr(e: COMError) -> str:
        details = ' '.join(str(d).replace('\r', '') for d in e.details if d)
        return f"{type(e).__name__}:{f' {e.text}' if e.text else ''} {details} {e.hresult}"
except ImportError as ex:
    comClient = None
    print(f"{type(ex).__name__}: {ex}\nWARNING: PDF generation will not be available.\nPlease run on Windows and install comtypes v1.1.7 or later: https://pypi.org/project/comtypes/\n")

AUTH_TOKEN: str = 'client_secret.json'
SCOPES: Sequence[str] = ('https://www.googleapis.com/auth/spreadsheets.readonly',)
SPREADSHEET_ID: str = '1kbMuGJaRR4gYTr9yaobskiENStj48m8wqhKRjlIQ0Tc'
ORIGINAL_TITLE: str = 'Русский'
LOCAL_PATTERN: str = '(местному)'
FILE_NAME: str = 'GangwayDict-%s'
TEMPLATE_FILE_NAME: str = FILE_NAME % 'Template.docx'
DOC_FILE_NAME: str = f'docx/{FILE_NAME}.docx'
PDF_FILE_NAME: str = f'pdf/{FILE_NAME}.pdf'

Table = Sequence[Sequence[str]]

class Block:
    def __init__(self, title: str, startRow: int, endRow: int, data: Sequence[str]) -> None:
        assert title
        assert startRow and endRow > startRow
        assert len(data) == endRow - startRow + 1
        self.title = title
        self.startRow = startRow
        self.endRow = endRow
        self.data = data

class Language:
    def __init__(self, headerRow: int, validateRow: int, data: Table, originals: Sequence[Block]) -> None:
        assert originals
        assert len(data) == 2
        self.native = data[0][headerRow]
        assert self.native
        self.name = data[1][headerRow]
        assert len(self.name) >= 5
        self.byName = self.name[:-1].lower()
        self.note = data[0][headerRow + 1]
        self.translator = data[0][validateRow]
        assert self.translator
        self.contact = data[1][validateRow]
        assert self.contact, f"No contact for {self.name}"
        print(self.native, self.name, self.translator, self.contact)
        self.docFileName = abspath(DOC_FILE_NAME % self.name)
        self.pdfFileName = abspath(PDF_FILE_NAME % self.name)
        self.data: Sequence[Tuple[str, Sequence[Tuple[str, str, str]]]] = tuple((block.title, tuple(zip(
                    (d.replace(LOCAL_PATTERN, self.byName) for d in block.data),
                    data[0][block.startRow : block.endRow + 1], data[1][block.startRow : block.endRow + 1]
                    ))) for block in originals)

    def renderDocx(self) -> None:
        doc = DocxTemplate(TEMPLATE_FILE_NAME)
        doc.render(self.__dict__)
        doc.save(self.docFileName)

    def renderPDF(self, msWord: Any) -> None:
        assert comClient
        try:
            wordDoc = msWord.Documents.Open(self.docFileName)
            try:
                wordDoc.SaveAs(self.pdfFileName, FileFormat = 17)
            finally:
                wordDoc.Close()
        except COMError as e:
            print(comErrorStr(e))
        except Exception as e:
            print(f"ERROR {type(e).__name__}: {e}\n{format_exc()}")

class GangwayDict:
    def __init__(self, worksheet: Worksheet) -> None:
        print("Parsing...")
        self.worksheet = worksheet
        data: Table = self.worksheet.get_all_values(majdim = 'COLUMNS', include_tailing_empty_rows = False)
        self.worksheet.unlink()
        #
        # Identify spreadsheet structure
        for (self.originalColumn, column) in enumerate(data):
            try:
                self.headerRow = column.index(ORIGINAL_TITLE)
                break
            except ValueError:
                pass
        else:
            assert False, "Original title not found"
        print(f"Original column: {self.excelColumn(self.originalColumn)}")
        print(f"Header row: {self.excelRow(self.headerRow)}")
        originals: Sequence[str] = data[self.originalColumn]
        for (self.validateRow, value) in reversed(tuple(enumerate(originals))):
            if value:
                break
        print(f"Validate row: {self.excelRow(self.validateRow)}")
        #
        # Get Originals
        titleRow: Optional[int] = None
        row: Optional[int] = None
        blocks: List[Block] = []
        for (row, value) in enumerate(originals[self.headerRow + 1 : self.validateRow], self.headerRow + 1):
            if value.isupper():
                if titleRow:
                    blocks.append(Block(originals[titleRow], titleRow + 1, row - 1, originals[titleRow + 1 : row]))
                titleRow = row
        assert titleRow, "No Originals blocks found"
        assert row > titleRow, "Last Originals block is empty"
        blocks.append(Block(originals[titleRow], titleRow + 1, row - 1, originals[titleRow + 1 : row]))
        self.originals: Sequence[Block] = tuple(blocks)
        #
        # Get languages
        col = self.originalColumn + 1
        languages: List[Language] = []
        while col < len(data):
            if data[col][self.validateRow].strip():
                languages.append(Language(self.headerRow, self.validateRow, data[col : col + 2], self.originals))
            col += 2
        self.languages: Sequence[Language] = tuple(languages)

    def render(self) -> None:
        print("Generating DOCX leaflets...")
        for language in self.languages:
            language.renderDocx()
        if comClient:
            print("Generating PDF leaflets...")
            try:
                try:
                    msWord = comClient.CreateObject('Word.Application')
                    for language in self.languages:
                        language.renderPDF(msWord)
                finally:
                    msWord.Quit()
            except COMError as e:
                print(comErrorStr(e))
            except Exception as e:
                print(f"ERROR {type(e).__name__}: {e}\n{format_exc()}")

    @staticmethod
    def excelColumn(col: int) -> str:
        if not col:
            return 'A'
        ret = ''
        while col:
            ret = f"{chr(ord('A') + col % 26)}{ret}"
            col //= 26
        return ret

    @staticmethod
    def excelRow(row: int) -> str:
        return f'{row + 1}'

def main() -> None:
    print("Loading Gangway Dictionary spreadsheet...")
    googleClient = authorize(client_secret = AUTH_TOKEN, scopes = SCOPES)
    spreadSheet = googleClient.open_by_key(SPREADSHEET_ID)
    GangwayDict(spreadSheet.worksheet()).render()
    print("DONE")

if __name__ == '__main__':
    main()
