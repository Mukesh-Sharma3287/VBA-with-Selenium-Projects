Attribute VB_Name = "MTable"
Option Explicit

''******************************************************************************
'' Module Name       : ModCustomer
'' Description       : This module is used to Extract the table data from webpage
''
''
'' Date             Developer             Action            Remarks
'' 12-Oct-2020     Mukesh Sharma          Created
''
''******************************************************************************

''******************************************************************************
'' Procedure           : ExtractedData()
'' Description         : To extract the table data from webpage
'' Arguments           : NA
'' Return              : NA
''******************************************************************************

'Pre-rquisite:
'Selenium wrapper should be installed in system
'Chrome Driver version should be same Chrome Browser Version
'Add Reference SeleniumWrapper Type Library

Public objChrmDriver As New SeleniumWrapper.WebDriver

Sub Extract_Table_Data()
    
    'Declare variables
    Dim strURL As String
    Dim tbl As SeleniumWrapper.Table
    Dim vArrayData() As Variant
    Dim iRows As Integer
    Dim iCols As Integer
    
    'Set URL
    strURL = "https://en.wikipedia.org/wiki"
    
    On Error GoTo Err_ExtractData
    
    'Open Url
    objChrmDriver.Start "chrome", strURL
    
    objChrmDriver.get "/1976_Summer_Olympics_medal_table"
    
    'Maximize the window
    objChrmDriver.windowMaximize
    
    'Clear contents
    wksTable.Range("A1").ClearContents
    
    
    'Set Table element
    Set tbl = objChrmDriver.findElementByXPath("//table[@class='wikitable sortable plainrowheaders jquery-tablesorter']").AsTable
    
    'Export data into Excel sheet
    vArrayData = tbl.GetData()

    iRows = UBound(vArrayData, 1)
    iCols = UBound(vArrayData, 2)

    'Export the data
    wksTable.Range(wksTable.Cells(1, 1), wksTable.Cells(iRows, iCols)).Value = vArrayData()
    
    On Error GoTo 0
     
    MsgBox "Done", vbInformation
    
    Set objChrmDriver = Nothing
    Exit Sub
    
Err_ExtractData:
    MsgBox Err.Description, vbInformation
    Set objChrmDriver = Nothing

End Sub
    
