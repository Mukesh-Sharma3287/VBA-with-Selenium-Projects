Attribute VB_Name = "ModFunctions"
Option Explicit

 Public wkbInv As Workbook
 Public wksInv As Worksheet
 Public lLastRow As Long
''******************************************************************************
'' Module Name       : ModFunctions
'' Description       : This module contains functions to use other module
''
''
'' Date             Developer             Action            Remarks
'' 10-Sep-2020     Mukesh Sharma          Created
''
''******************************************************************************

''******************************************************************************
'' Procedure           : Show_UserForm
'' Description         : To Show User form
'' Arguments           : NA
'' Return              : NA
''******************************************************************************

Sub Show_UserForm()
    
    On Error Resume Next
    frmTLPA.Show vbModeless
    On Error GoTo 0

End Sub

''******************************************************************************
'' Procedure           : fnGetInvoices
'' Description         : Get Invoices from download OTM portal
'' Arguments           : NA
'' Return              : Scripting.Dictionary: return invoices as dictionary
''******************************************************************************
 
Public Function fnGetInvoices() As Scripting.Dictionary
    
    Dim strFldPath As String
    Dim objDictDesc As New Scripting.Dictionary
    Dim objFSO As New FileSystemObject
    Dim objFl As Scripting.File
    Dim objFld As Scripting.Folder
    Dim strFilePath As String
    Dim vkey As Variant
    Dim lCounter As Long
    
    strFldPath = objFSO.BuildPath("C:\Users\" & Environ("UserName"), "\Downloads")

    Set objFld = objFSO.GetFolder(strFldPath)

    'Loop through the find the file name
    For Each objFl In objFld.Files
        If DateTime.TimeSerial(Hour(objFl.DateLastModified), Minute(objFl.DateLastModified), Second(objFl.DateLastModified)) > DateTime.TimeSerial(Hour(Now), Minute(Now) - 2, Second(Now)) And Format(Date, "mm/dd/yyyy") = Format(objFl.DateLastModified, "mm/dd/yyyy") Then
            strFilePath = objFl
            Exit For
        End If

    Next
    
    On Error Resume Next
    Set wkbInv = Workbooks.Open(strFilePath)
    
    Set wksInv = wkbInv.Sheets(1)
    
    lLastRow = wksInv.Cells(Rows.Count, "A").End(xlUp).Row
    
    'Call function to sort descending
    Call fnSorting(wksInv, lLastRow, "Desc")
    
    objDictDesc.RemoveAll
    For lCounter = 2 To lLastRow
        If wksInv.Range("V" & lCounter).Value <> Empty Then
             objDictDesc.Add wksInv.Range("A" & lCounter).Value, wksInv.Range("V" & lCounter).Value
        End If
     Next
     
    Set fnGetInvoices = objDictDesc
     
    
End Function

''******************************************************************************
'' Procedure           : fnDictAsc
'' Description         : Get Invoices from download OTM portal
'' Arguments           : NA
'' Return              : Scripting.Dictionary: return invoices Ascending oreder as dictionary
''******************************************************************************


Public Function fnDictAsc() As Scripting.Dictionary
    
    Dim objDictAsc As New Scripting.Dictionary
    Dim lCounter As Long
   
    'Call function to sort Ascending
    Call fnSorting(wksInv, lLastRow, "Asc")
    
    objDictAsc.RemoveAll
   
    For lCounter = 2 To lLastRow
        If wksInv.Range("V" & lCounter).Value = Empty Then
             objDictAsc.Add wksInv.Range("A" & lCounter).Value, wksInv.Range("V" & lCounter).Value
        End If
    Next
    
    Set fnDictAsc = objDictAsc
    
    wkbInv.Close False
    On Error GoTo 0
    
End Function

''******************************************************************************
'' Procedure           : fnSorting
'' Description         : Sorting Downloaded invoices from OTM protal as per codition
'' Arguments           : wksInv As Worksheet, lLastRow As Long, strSorting As String
'' Return              : NA
''******************************************************************************

Public Function fnSorting(wksInv As Worksheet, lLastRow As Long, strSorting As String) As String
    
    'Sorting
    With wksInv
        .Activate
        .Sort.SortFields.Clear
        If strSorting = "Desc" Then
            .Sort.SortFields.Add Key:=Range("V2:V" & lLastRow), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
        Else
            .Sort.SortFields.Add Key:=Range("V2:V" & lLastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        End If
        
    End With
    
    With wksInv.Sort
        .SetRange Range("A1:AX" & lLastRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    

End Function
    
 ''******************************************************************************
'' Procedure           : TextJoin
'' Description         : To join multiple invoices with comma
'' Arguments           : NA
'' Return              : NA
''******************************************************************************

Public Sub TextJoin()
    
    'Declare variables
    Dim lLastRow As Long
    Dim lCounter As Long
    Dim lLastCol As Long
    Dim lInnerCounter As Long
    Dim strInvoices As String
    
    lLastRow = wksTextJoin.Cells(Rows.Count, "B").End(xlUp).Row
    lLastCol = wksTextJoin.Cells(2, wksTextJoin.Columns.Count).End(xlToLeft).Column
    
    'Loop through club invoices
    For lCounter = 2 To lLastRow
        strInvoices = ""
        For lInnerCounter = 4 To lLastCol
            strInvoices = strInvoices & "," & wksTextJoin.Cells(lCounter, lInnerCounter).Value
        Next
        wksTextJoin.Range("C" & lCounter).Value = Right(strInvoices, Len(strInvoices) - 1)
    Next
    
    MsgBox "Done", vbInformation

End Sub

Public Function checkFileExist(dtDate As Date) As Boolean



    Dim strFldPath As String
    Dim objDictDesc As New Scripting.Dictionary
    Dim objFSO As New FileSystemObject
    Dim objFl As Scripting.File
    Dim objFld As Scripting.Folder
    Dim strFilePath As String
    Dim vkey As Variant
    Dim lCounter As Long
   
    checkFileExist = False
    strFldPath = objFSO.BuildPath("C:\Users\" & Environ("UserName"), "\Downloads")

    Set objFld = objFSO.GetFolder(strFldPath)

    'Loop through the find the file name
    For Each objFl In objFld.Files
        If InStr(1, objFl.Name, "REPORT", vbTextCompare) = 0 Then
            If DateTime.TimeSerial(Hour(objFl.DateLastModified), Minute(objFl.DateLastModified), Second(objFl.DateLastModified)) > DateTime.TimeSerial(Hour(dtDate), Minute(Now) - 1, Second(dtDate)) And Format(Date, "mm/dd/yyyy") = Format(objFl.DateLastModified, "mm/dd/yyyy") Then
                checkFileExist = True
                Exit For
            End If
        End If
    
    Next
    
   
End Function

Public Function NoOfInvoices(lRow As Long) As Long
    
    Dim vArray As Variant
    
    vArray = VBA.Split(wksInvoices.Range("B" & lRow), ",")
    
     NoOfInvoices = UBound(vArray) + 1
        
End Function
