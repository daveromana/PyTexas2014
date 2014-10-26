#Python and Spreadsheets: State of the Union, Oct 2014
##Kojo Idrissa
###Rouge Pythonista
###Educator & Accountant w/ an MBA

#Setup

##Not a lot of code here
+  Interesting code provided by the specific use case

##Context/Background
+  My role as a Python contributor
    *  Grow the community
    *  Bring non-programmers into the fold
+  Solutions for people who aren't professional developers
    *  Solutions obvious to experienced developers often aren't available to non-developers
+  Two primary audiences
    *  Spreadsheet users who want to "step their game up"
    *  Python developers confronted with spreadsheets

#Three approaches + "Other"
##Working with spreadsheet files
##Using Python code in a spreadsheet app
##Scripting spreadsheet apps

###Other: options I'm still investigating

#Working with spreadsheet files
-  Use case: You've been given a spreadsheet full of data to process, but not in a way that's convenient in your spreadsheet app.
    +  Perfect entry point for creating new Pythonistas. Spreadsheet hackers are ALMOST programmers.
-  [OpenPyXL](http://pythonhosted.org//openpyxl/)
    +  get it from [PyPi](https://pypi.python.org/pypi/openpyxl) 
    +  Python >=2.6
    +  works with .xlsx files
    +  I've demoed a larger example program
-  [Python-Excel](http://www.python-excel.org) packages
    +  xlrd; xlwt; xlutils
    +  Support for both *.xls and *.xlsx files
-  both under active development; my choice for a new person is OpenPyXL, as the documentation is better for getting started
-  For working with *.ods files, start at http://www.opendocumentformat.org/developers/ Python is at the top...as it **should** be. :-)
-  Example Results

#Running Python code IN a spreadsheet app
-  Use case: Adding functionality from Python to your spreadsheet-based workflow
    -  I see this as the largest growth area: replacing VB with Python
    -  For current Pythonistas confronted with spreadsheets
-  [Python for Excel with PyXLL (Enthought's Canopy plugin)](http://vimeo.com/89024595?utm_source=Python+Weekly+Newsletter&utm_campaign=42441d6a58-Python_Weekly_Issue_133_April_3_2014&utm_medium=email&utm_term=0_9e26887fc5-42441d6a58-312680573) pronounced "Pixel"
    +  Yes Glen, it's a DLL. :-)
    *  Centralized management of functions used in your spreadsheets. This is difficult to do with Excel, where each person uses their own formulas or VBA code
        -  The spreadsheet from Hong Kong gets used in Chicago, which gets sent to Houston. Changes made at every step, often based on "preference".
        -  No consistent formula methodology; differing computational assumptions; NONE of those people are programmers.
        -  No uniform method for the results that are generated
    *  Asynchronous functions allow you to keep Excel from "freezing" while Python is processing.
    +  Enthought's [PyXLL](https://www.enthought.com/products/pyxll/) page
-  xlwings and ExcelPython
    +  First five mintues of these [EuroPython 2014 Lightning Talks](http://www.youtube.com/watch?v=qDzeSGv28kU&feature=youtu.be&t=35s) focuses on these two packages. The authors are planning to merge them.
    -  xlwings: [Replace your VBA code with Python](http://xlwings.org/?utm_source=Python+Weekly+Newsletter&utm_campaign=42441d6a58-Python_Weekly_Issue_133_April_3_2014&utm_medium=email&utm_term=0_9e26887fc5-42441d6a58-312680573)
    -  [ExcelPython: call Python scripts from Excel/VBA](https://github.com/ericremoreynolds/excelpython)
+  [PyInEx](https://www.pytexas.org/2014/talks/35/); Python interpreter embedded IN Excel!!
    *  I won't talk about this, because there's already an entire presentation on it [scheduled for Sunday at 1pm](https://www.pytexas.org/2014/talks/35/). You should go see that. :-)

#Scripting spreadsheet apps (still researching)
-  Use Case: including a spreadsheet app as part of a Python-based work flow
+  Open/LibreOffice (via PyUno) 
    *  https://wiki.openoffice.org/wiki/PyUNO_bridge#Tutorial
    *  https://wiki.openoffice.org/wiki/Python
    *  https://www.google.com/search?client=safari&rls=en&q=libreoffice+python+uno&ie=UTF-8&oe=UTF-8&gfe_rd=cr&ei=UpvtU-WtLefU8gfDmIGoBg
    *  http://en.wikipedia.org/wiki/Universal_Network_Objects
    *  https://pypi.python.org/pypi/unotools/0.3.2
    *  http://api.libreoffice.org/examples/examples.html#python_examples
    *  [UNOTools](https://bitbucket.org/t2y/unotools)

#Issues/Caveats
-  vulnerability to changes in application or file format
    +  Excel/MS famous for this
    +  rise of Open Document formats may minimize this

#Other (still researching)
*  Google Docs spreadsheets with Python?
    -  [GSpread](https://pypi.python.org/pypi/gspread) in PyPi
-  PeeWee to explore CSV files?
    +  http://charlesleifer.com/blog/using-peewee-to-explore-csv-files/
    +  http://peewee.readthedocs.org/en/latest/peewee/playhouse.html#csv-loader
-  Other CSV file options
    +  Not my preferred approach

#That's all! Questions?
