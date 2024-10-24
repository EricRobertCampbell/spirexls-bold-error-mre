#!/usr/bin/env python3
""" Source: existing code, tutorial at https://www.e-iceblue.com/Tutorials/Python/Spire.XLS-for-Python/Program-Guide/Data/Python-Set-or-Change-Fonts-in-Excel.html
 """

from spire.xls import *
from spire.xls.common import *
from spire.doc import (
    Document as SpireDocument,
    ToPdfParameterList,
    PrivateFontPath,
    FileFormat,
)
from pathlib import Path
import random


file_path = Path("testbolded.xlsx")


def spire_doc_test() -> None:
    doc = SpireDocument()

    section = doc.AddSection()
    section.PageSetup.HeaderDistance = 0
    section.PageSetup.FooterDistance = 0
    section.PageSetup.Margins.Top = 87
    section.PageSetup.Margins.Left = 17
    section.PageSetup.Margins.Right = 17
    section.PageSetup.Margins.Bottom = 87
    paragraph = section.AddParagraph()
    paragraph.AppendText("Hello World!")
    doc.SaveToFile(f"testdoc.docx", FileFormat.Docx2013)
    # doc.SaveToFile(f"test{random.randint(1,1000000)}.docx", FileFormat.Docx2013)
    doc.Close()

    pdf_doc = SpireDocument()
    pdf_doc.LoadFromFile("testdoc.docx")
    ppl = ToPdfParameterList()
    ppl.IsEmbeddedAllFonts = True
    pdf_doc.SaveToFile("testdoc.pdf", ppl)
    pdf_doc.Close()


def main() -> None:

    print("generating Doc and PDF")
    spire_doc_test()
    print("Done generating Doc")

    wb = Workbook()
    wb.Worksheets.Clear()
    sheet = wb.Worksheets.Add("Test Sheet")

    # fails in a loop
    for pos in [(1, 1), (1, 5)]:
        sheet.Range[pos].Value = "Test Text"

        # Method 1
        # sheet.Range[pos].Style.Font.IsBold = True

        # Method 2
        # fontStyle = wb.Styles.Add('headerFontStyle')
        # fontStyle.Font.IsBold = True
        # flag = CellStyleFlag()
        # flag.FontBold = True
        # sheet.Range[pos].ApplyStyle(fontStyle, flag)

        # Also fails on italic - Method 1
        # sheet.Range[pos].Style.Font.IsItalic = True

        # Also fails on italic - Method 2
        # fontStyle = wb.Styles.Add("headerFontStyle")
        # fontStyle.Font.IsItalic = True

        # also fails when using an explicit range
        # sheet.Range[1, 1].Style.Font.IsBold = True

    wb.SaveToFile(str(file_path), ExcelVersion.Version2016)
    print(f"Done save to {file_path}")


if __name__ == "__main__":
    main()
