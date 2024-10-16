#!/usr/bin/env python3
""" Source: existing code, tutorial at https://www.e-iceblue.com/Tutorials/Python/Spire.XLS-for-Python/Program-Guide/Data/Python-Set-or-Change-Fonts-in-Excel.html
 """

from spire.xls import *
from spire.xls.common import *

file_path = "bolded.xlsx"
def main() -> None:
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
		sheet.Range[pos].Style.Font.IsItalic = True

		# Also fails on italic - Method 2
		fontStyle = wb.Styles.Add('headerFontStyle')
		fontStyle.Font.IsItalic = True
		flag = CellStyleFlag()
		flag.FontItalic = True
		sheet.Range[pos].ApplyStyle(fontStyle, flag)

    # also fails when using an explicit range
    # sheet.Range[1, 1].Style.Font.IsBold = True

	wb.SaveToFile(path, FileFormat.Version2016)
	print(f"Done save to {file_path}")

if __name__ == "__main__":
	main()
