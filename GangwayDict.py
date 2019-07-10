#!/usr/bin/env python3

from typing import Iterator, List, Optional, Sequence, Tuple

try:
    from docxtpl import DocxTemplate # type: ignore
except ImportError as ex:
    raise ImportError(f"{type(ex).__name__}: {ex}\n\nPlease install docxtpl v0.6.3 or later: https://pypi.org/project/docxtpl\n")

try:
    from pygsheets import authorize # type: ignore
    from pygsheets.worksheet import Worksheet # type: ignore
except ImportError as ex:
    raise ImportError(f"{type(ex).__name__}: {ex}\n\nPlease install pygsheets v2.0.1 or later: https://pypi.org/project/pygsheets\n")

AUTH_TOKEN: str = 'client_id.json'
SCOPES: Sequence[str] = ('https://www.googleapis.com/auth/spreadSheets.readonly',)
SPREADSHEET_ID: str = '1kbMuGJaRR4gYTr9yaobskiENStj48m8wqhKRjlIQ0Tc'
ORIGINAL_TITLE: str = 'Русский'
LOCAL_PATTERN: str = '(местному)'
DOC_FILE_NAME: str = 'GangwayDict-%s.docx'

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
        self.isoCode = data[0][headerRow]
        assert len(self.isoCode) == 2
        assert self.isoCode.isupper()
        self.name = data[1][headerRow]
        assert len(self.name) >= 5
        self.byName = f"{self.name[0].lower()}{self.name[1:-1]}"
        self.translator = data[0][validateRow]
        assert self.translator
        self.contact = data[1][validateRow]
        assert self.contact
        print(self.isoCode, self.name, self.translator, self.contact)
        self.data: Sequence[Tuple[str, Sequence[Tuple[str, str, str]]]] = tuple((block.title, tuple(zip(
                    (d.replace(LOCAL_PATTERN, self.byName) for d in block.data),
                    data[0][block.startRow : block.endRow + 1], data[1][block.startRow : block.endRow + 1]
                    ))) for block in originals)

    def render(self) -> None:
        doc = DocxTemplate(DOC_FILE_NAME % 'Template')
        doc.render(self.__dict__)
        doc.save(DOC_FILE_NAME % self.isoCode)

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
        print("Generating output files...")
        for language in self.languages:
            language.render()

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

    @classmethod
    def excelAddr(cls, row: int, col: int) -> str:
        return f'{cls.excelColumn(col)}{cls.excelRow(row)}'

def main() -> None:
    print("Loading Gangway Dictionary spreadsheet...")
    googleClient = authorize(client_secret = AUTH_TOKEN, scopes = SCOPES)
    spreadSheet = googleClient.open_by_key(SPREADSHEET_ID)
    GangwayDict(spreadSheet.worksheet()).render()
    print("DONE")

if __name__ == '__main__':
    main()
