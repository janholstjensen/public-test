# RDKit for Excel
A simple Excel add-in that gives access to RDKit functions. OK - we now have more cowbell.

## Prerequisites when using standard Python 2.7
* Python 2.7 from python.org
* Pywin32 (win32com)
	* pip install pypiwin32
* Microsoft visual C 9.0 for python 2.7
	* Get VCForPython27.msi from https://www.microsoft.com/en-us/download/details.aspx?id=44266
* RDKit binaries from SourceForge
* pip install numpy
* pip install Pillow
.
* Register RDKit binaries - set PYTHONPATH and PATH environment variables.
* Ensure that you have the needed MSVC runtime libs for RDKit.
* Add "C:\Python27\Lib\site-packages\pywin32_system32" to PATH so "pythoncomloader27.dll" can be loaded by Excel.


## Prerequisites when using Conda (Python 2.7 version)
Assuming that you have Miniconda2 4.3.11 or later installed.

* Install RDKit
	* conda install -c rdkit rdkit

* Install Microsoft visual C 9.0 for python 2.7
	* Get VCForPython27.msi from https://www.microsoft.com/en-us/download/details.aspx?id=44266


# To compile and register add-in
Open command prompt as administrator.

Setup VC variables (your actual path to vcvarsall.bat will be different) and register add-in:

```
"C:\Users\esben\AppData\Local\Programs\Common\Microsoft\Visual C++ for Python\9.0\vcvarsall.bat"
python rdkitXL_server.py
```

Expected output:

```
C:\Apps\rdkit4excel>python rdkitXL_server.py
Compiling C:\Apps\rdkit4excel\RDKitXL.idl
Microsoft (R) 32b/64b MIDL Compiler Version 7.00.0555
Copyright (c) Microsoft Corporation. All rights reserved.
Processing C:\Apps\rdkit4excel\RDKitXL.idl
RDKitXL.idl
[...]
Processing C:\Users\jan\AppData\Local\Programs\Common\Microsoft\Visual C++ for Python\9.0\WinSDK\Include\ocidl.acfocidl.acf
Registering C:\Apps\rdkit4excel\RDKitXL.tlb
Registered: Python.RDKitXL

C:\Apps\rdkit4excel>
```

Register add-in in Excel via File -> options -> addins -> manage: Excel add-ins, GO -> Automation -> RDKitXL object.
You may get a message box asking "Cannot find add-in 'pythoncomloader27.dll'. Delete from list?". Answer "No" to this.


## Troubleshooting
To register server in debug mode, compile/register with 

```
python rdkitXL_server.py --debug
```

Open PythonWin and open the Tools -> Trace collector debugging tool to watch the messages and print statements


# Known BUGS
The IDL generation and compilation will fail if a default parameter contains " (double-quotes) in a string quoted by ' (single-quotes).
