Attribute VB_Name = "ModMainOTM"

Option Explicit

''******************************************************************************
'' Module Name       : ModMainOTM
'' Description       : This module is use to connect with OTM portal and perform activites
''
''
'' Date             Developer             Action            Remarks
'' 10-Sep-2020     Mukesh Sharma          Created
''
''******************************************************************************

''******************************************************************************
'' Procedure           : OTMWeb_Main
'' Description         : Perform Shimpent Planing activities in OTM Portal
'' Arguments           : NA
'' Return              : NA
''******************************************************************************

Sub Main()
    
    Dim lLastRow As Long
    Dim vInvoices As Variant
    Dim strInvoices As String
    
    On Error GoTo Err_Found
    lLastRow = wksInvoices.Cells(Rows.Count, "B").End(xlUp).Row
    'Validation
    If lLastRow = 1 Then
        MsgBox "Please Set the Invoices in Invoices Tab to proceed", vbCritical
        Exit Sub
    End If
    
    'Call function to login the OTM portal
    If ModLoginOTM.fnLoginOTM = False Then
        MsgBox "Unable to Login OTM Web portal, Please try ugain"
        Exit Sub
    End If
    
   
    For lCounter = 2 To lLastRow
        
        'Start Time
        wksInvoices.Range("H" & lCounter).Value = Format(Now, "dd-mmm-yyyy h:mm:ss")
        
        strInvoices = wksInvoices.Range("B" & lCounter).Value
        If strInvoices <> "" Then
            Call ModOTMWebProcess.fnOTMWebProcess(strInvoices, lCounter)
        End If
        
        'End Time
        wksInvoices.Range("I" & lCounter).Value = Format(Now, "dd-mmm-yyyy h:mm:ss")
    Next
    On Error GoTo 0
    
    objChrmDriver.Close
    Set objChrmDriver = Nothing
    
    MsgBox "Done", vbInformation
    
    Exit Sub
Err_Found:
    
    MsgBox Err.Description
    Set objChrmDriver = Nothing
   
End Sub

'''******************************************************************************
'' Procedure           : MainApproval
'' Description         : Perform Approval activities in OTM Portal
'' Arguments           : NA
'' Return              : NA
''******************************************************************************

Sub MainApproval()
    
    'Declare variables
    Dim lLastRow As Long
    Dim lCounter As Long
    Dim strMessage As String
    Dim vInvoices As Variant
    Dim strInvoices As String
    Dim bElement As Boolean
    Dim objElement As SeleniumWrapper.WebElement
    Dim objBy As New SeleniumWrapper.by
    Dim bPresent As Boolean
    Dim objElements As SeleniumWrapper.WebElementCollection
    Dim lDelay As Long
   
    bElement = False
    strMessage = "Not Found "
    lDelay = 15000
    
    'Error handling
    On Error GoTo Err_ElementNotFound
    lLastRow = wksApproval.Cells(Rows.Count, "B").End(xlUp).Row
    
    'Validation
    If lLastRow = 1 Then
        MsgBox "Please Set the Invoices in Invoices Tab to proceed", vbCritical
        Exit Sub
    End If
    
    'Call function to login the OTM portal
     If ModLoginOTM.fnLoginOTM = False Then
        MsgBox "Unable to Login OTM Web portal, Please try ugain"
        Exit Sub
    End If
    
    objChrmDriver.Wait 5000
    bElement = False
      'Click on Shipment Management-I
     If objChrmDriver.IsElementPresent(objBy.XPath("//img[contains(@id,'mainContentRegion')]"), lDelay) = True Then
        objChrmDriver.findElementsByXPath("//img[contains(@id,'mainContentRegion')]").Item(1).Click
         bElement = True
     End If

     If bElement = False Then
        strMessage = strMessage & ";" & "Click on Shipment Management-I"
     End If

     bElement = False
     objChrmDriver.Wait 5000
     'Click on Shipment Management-II
     If objChrmDriver.IsElementPresent(objBy.XPath("//img[contains(@id,'mainContentRegion')]"), lDelay) = True Then
        objChrmDriver.findElementsByXPath("//img[contains(@id,'mainContentRegion')]").Item(5).Click
         bElement = True
     End If


     If bElement = False Then
        strMessage = strMessage & ";" & "Click on Shipment Management-II"
     End If

     bElement = False
     objChrmDriver.Wait 5000
     'Click on Buy Shipments
     If objChrmDriver.IsElementPresent(objBy.XPath("//img[contains(@id,'mainContentRegion')]"), lDelay) = True Then
        objChrmDriver.findElementsByXPath("//img[contains(@id,'mainContentRegion')]").Item(13).Click
         bElement = True
     End If

      If bElement = False Then
        strMessage = strMessage & ";" & "Click on Buy Shipments"
     End If

     bElement = False
    'Switch frame
    objChrmDriver.Wait 5000
    If objChrmDriver.IsElementPresent(objBy.XPath("//iframe[contains(@id,'mainContentRegion')]"), lDelay) = True Then
         Set objElement = objChrmDriver.findElementByXPath("//iframe[contains(@id,'mainContentRegion')]")
         objChrmDriver.switchToFrame objElement
         bElement = True
     End If

      If bElement = False Then
        strMessage = strMessage & ";" & "Switch frame for Buy Shipment finder"
     End If

    For lCounter = 2 To lLastRow
        
        'Start Time
        wksApproval.Range("F" & lCounter).Value = Format(Now, "dd-mmm-yyyy h:mm:ss")
        
        strInvoices = wksApproval.Range("B" & lCounter).Value
        If strInvoices <> "" Then
            Call ModOTMApproval.fnOTMApproval(strInvoices, lCounter)
        End If
        
        'End Time
        wksApproval.Range("G" & lCounter).Value = Format(Now, "dd-mmm-yyyy h:mm:ss")
    Next
    On Error GoTo 0
    
    'Close current web window
    objChrmDriver.Close
    Set objChrmDriver = Nothing
    
    'Prompt the message to user
    MsgBox "Done", vbInformation
    Exit Sub

Err_ElementNotFound:
    objChrmDriver.Close
    Set objChrmDriver = Nothing
    wksApproval.Range("E" & lLastRow).Value = Err.Description
   
End Sub

