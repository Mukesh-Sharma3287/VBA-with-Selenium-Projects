Attribute VB_Name = "ModOTMApproval"
Option Explicit

''******************************************************************************
'' Procedure           : fnOTMWebProcess
'' Description         : Perform Invoices of Approval in OTM Portal
'' Arguments           : strInvoices As String, lLastRow As Long
'' Return              : NA
''******************************************************************************

Public Function fnOTMApproval(strInvoices As String, lLastRow As Long)

    bElement = False
    strMessage = "Not Found "
    lDelay = 15000
    
    'Error handling
    On Error GoTo Err_ElementNotFound
     bElement = False
     'Enter Shipment ID
     If objChrmDriver.IsElementPresent(objBy.XPath("//input[@name='shipment/xid']"), 30000) = True Then
        Set objElement = objChrmDriver.findElementByXPath("//input[@name='shipment/xid']")
        objElement.SendKeys strInvoices
         bElement = True
      End If

     If bElement = False Then
        strMessage = strMessage & ";" & "Enter Shipment ID"
     End If
    
    
    If InStr(1, strInvoices, ",", vbTextCompare) > 0 Then
         bElement = False
         'Select drop down item
         If objChrmDriver.IsElementPresent(objBy.XPath("//select[@name='shipment/xid_operator']"), lDelay) = True Then
            Set objSelect = objChrmDriver.findElementByXPath("//select[@name='shipment/xid_operator']").AsSelect
            
            If Len(strInvoices) > 10 Then
                objSelect.SelectByText "One Of"
                bElement = True
            End If
         End If
    
         If bElement = False Then
            strMessage = strMessage & ";" & "Select Drop down for One Of"
         End If
    End If
     
      bElement = False
     'Click on search button
     If objChrmDriver.IsElementPresent(objBy.XPath("//a[@name='search_button']"), lDelay) = True Then
        Set objElement = objChrmDriver.findElementByXPath("//a[@name='search_button']")
        objElement.Click
         bElement = True
     End If

      If bElement = False Then
        strMessage = strMessage & ";" & "Click on search button"
     End If

    bElement = False
     'Click on All Select button
    If objChrmDriver.IsElementPresent(objBy.XPath("//input[@id='rgSGSec.1.1.1.1.check']"), lDelay) = True Then
       objChrmDriver.findElementByXPath("//input[@id='rgSGSec.1.1.1.1.check']").Click
        bElement = True
    End If

    If bElement = False Then
       strMessage = strMessage & ";" & "Click on All select button"
    End If
           
    
     bElement = False
      'Click on Actions button
     If objChrmDriver.IsElementPresent(objBy.XPath("//a[@class='enButton']"), lDelay) = True Then
        objChrmDriver.findElementsByXPath("//a[@class='enButton']").Item(0).Click
         bElement = True
     End If

     If bElement = False Then
        strMessage = strMessage & ";" & "Click on action button"
     End If
    
    If objChrmDriver.IsElementPresent(objBy.ID("actionFrame")) = True Then
        Set objElement = objChrmDriver.FindElementById("actionFrame")
        objChrmDriver.switchToFrame objElement
    End If
    
       
     bElement = False
     'Click on Shipment Management
     If objChrmDriver.IsElementPresent(objBy.XPath("//tr[@id='actionTree.1_2']"), lDelay) = True Then
        objChrmDriver.findElementByXPath("//tr[@id='actionTree.1_2']").Click
         bElement = True
     End If

    If bElement = False Then
        strMessage = strMessage & ";" & "Click on Shipment Management"
     End If

     bElement = False
     If objChrmDriver.IsElementPresent(objBy.XPath("//span[@id='actionTree.1_2_7.l']"), lDelay) = True Then
        objChrmDriver.findElementByXPath("//span[@id='actionTree.1_2_7.l']").Click
         bElement = True
     End If

      If bElement = False Then
        strMessage = strMessage & ";" & "Click on Manual Action"
     End If

     bElement = False
     If objChrmDriver.IsElementPresent(objBy.XPath("//a[@id='actionTree.1_2_7_2.k']"), lDelay) = True Then
        objChrmDriver.findElementByXPath("//a[@id='actionTree.1_2_7_2.k']").Click
         bElement = True
     End If

      If bElement = False Then
        strMessage = strMessage & ";" & "Click on Manual approve"
      End If
    
    objChrmDriver.Wait 3000
    'Close the current window
     objChrmDriver.switchToWindow objChrmDriver.WindowHandles(1)
     objChrmDriver.Close
     objChrmDriver.switchToWindow objChrmDriver.WindowHandles(0)
    
     bElement = False
    'Switch frame
    objChrmDriver.Wait 2000
    If objChrmDriver.IsElementPresent(objBy.XPath("//iframe[contains(@id,'mainContentRegion')]"), lDelay) = True Then
         Set objElement = objChrmDriver.findElementByXPath("//iframe[contains(@id,'mainContentRegion')]")
         objChrmDriver.switchToFrame objElement
         bElement = True
     End If

     If bElement = False Then
       strMessage = strMessage & ";" & "Switch frame for Buy Shipment finder"
    End If
     
      bElement = False
     'Click on new query
     If objChrmDriver.IsElementPresent(objBy.XPath("//a[@class='enButton']"), lDelay) = True Then
        objChrmDriver.findElementsByXPath("//a[@class='enButton']").Item(1).Click
         bElement = True
     End If

     If bElement = False Then
        strMessage = strMessage & ";" & "Click on New Query"
     End If
     
 On Error GoTo 0
     
     If strMessage = "Not Found " Then
        wksApproval.Range("C" & lLastRow).Value = "Completed"
        wksApproval.Range("D" & lLastRow).Value = ""
     Else
        wksApproval.Range("C" & lLastRow).Value = "Not Completed"
        wksApproval.Range("D" & lLastRow).Value = strMessage
     End If
     
     Exit Function

Err_ElementNotFound:
     wksApproval.Range("E" & lLastRow).Value = Err.Description
     objChrmDriver.Close
End Function


