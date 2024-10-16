Spire.Xls - Style Change Failure MRE

This is an MRE for a bug in Spire.Xls v14.7.3. When changing the style of a cell, consistently there's an error about `ffi_prep_cif_var failed`.

```
Traceback (most recent call last):
│  File "/home/eric/documents/spirexls-bold-error-mre/./generate_bold_excel.py", line 45, in <module>
│    main()
│  File "/home/eric/documents/spirexls-bold-error-mre/./generate_bold_excel.py", line 29, in main
│    sheet.Range[pos].Style.Font.IsItalic = True
│  File "/home/eric/documents/spirexls-bold-error-mre/venv/lib/python3.9/site-packages/spire/xls/ExcelFont.py", line 28, in IsItalic
│    CallCFunction(GetDllLibXls().ExcelFont_set_IsItalic, self.Ptr, value)
│  File "/home/eric/documents/spirexls-bold-error-mre/venv/lib/python3.9/site-packages/spire/xls/common/__init__.py", line 105, in CallCFunction
│    result = func(*args, **kwargs)
│RuntimeError: ffi_prep_cif_var failed
```

The included Python file should generate an Excel document with some bolded text. For the various ways I've attempted to do this, uncomment the marked lines of code in the `main()` function.

To run:

```
python -m venv venv
source venv/bin/activate
pip install -r requirements.txt
chmod u+x generate_bold_excel.py
./generate_bold_excel.py
```

This code was run on Ubuntu 22.04.1 LTS with Python 3.9.16.
