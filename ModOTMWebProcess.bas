Attribute VB_Name = "ModOTMWebProcess"
Option Explicit

''******************************************************************************
'' Procedure           : fnOTMWebProcess
'' Description         : Perform Invoices of Shimpent planing in OTM Portal
'' Arguments           : strInvoices As String, lLastRow As Long
'' Return              : NA
''******************************************************************************

Public Function fnOTMWebProcess(strInvoices As String, lLastRow As Long)

    'Declare variables
    Dim vInvoices As Variant
    Dim lItem As Long
    Dim lCounterInv As Long
    Dim vkey As Variant
    Dim objDictDesc As New Scripting.Dictionary
    Dim objDictAsc As New Scripting.Dictionary
    Dim iCount As Integer
    Dim dtDate As Date
    
    

    bElement = False
    strMessage = "Not Found "
    lDelay = 15000
    
    'Error handling
    On Error GoTo Err_ElementNotFound
    objChrmDriver.Wait 5000
    bElement = False
      'Click on Order management
     If objChrmDriver.IsElementPresent(objBy.XPath("//img[contains(@id,'mainContentRegion')]"), lDelay) = True Then
        objChrmDriver.findElementsByXPath("//img[contains(@id,'mainContentRegion')]").Item(0).Click
         bElement = True
     End If


     If bElement = False Then
        strMessage = strMessage & ";" & "Click on Order management1"
     End If

     bElement = False
     objChrmDriver.Wait 5000

     'Click on release management
     If objChrmDriver.IsElementPresent(objBy.XPath("//img[contains(@id,'mainContentRegion')]"), lDelay) = True Then
        objChrmDriver.findElementsByXPath("//img[contains(@id,'mainContentRegion')]").Item(5).Click
         bElement = True
     End If


     If bElement = False Then
        strMessage = strMessage & ";" & "Click on release management1"
     End If

     bElement = False
     objChrmDriver.Wait 5000
     'Click on Report
     If objChrmDriver.IsElementPresent(objBy.XPath("//img[contains(@id,'mainContentRegion')]"), lDelay) = True Then
        objChrmDriver.findElementsByXPath("//img[contains(@id,'mainContentRegion')]").Item(12).Click
         bElement = True
     End If

      If bElement = False Then
        strMessage = strMessage & ";" & "Click on report"
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
        strMessage = strMessage & ";" & "Switch frame for order release id"
     End If

     bElement = False
    
      'Enter Order Release ID
     If objChrmDriver.IsElementPresent(objBy.XPath("//input[@name='order_release/xid']"), 30000) = True Then
        Set objElement = objChrmDriver.findElementByXPath("//input[@name='order_release/xid']")
        objElement.SendKeys strInvoices
         bElement = True
      End If

     If bElement = False Then
        strMessage = strMessage & ";" & "Enter Order Release ID"
     End If

     bElement = False
    
    If InStr(1, strInvoices, ",", vbTextCompare) > 0 Then
         'Select drop down item
         If objChrmDriver.IsElementPresent(objBy.XPath("//select[@name='order_release/xid_operator']"), lDelay) = True Then
            Set objSelect = objChrmDriver.findElementByXPath("//select[@name='order_release/xid_operator']").AsSelect
            
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
     

     'Click on Export button
     If objChrmDriver.IsElementPresent(objBy.XPath("//a[@class='enButton']"), lDelay) = True Then
        objChrmDriver.findElementsByXPath("//a[@class='enButton']").Item(3).Click
         bElement = True
     End If

     If bElement = False Then
        strMessage = strMessage & ";" & "Click on export button"
     End If

     bElement = False

     'Click on OK
     If objChrmDriver.IsElementPresent(objBy.XPath("//button[@id='resultsPage:ExportPopupDialog::ok']"), lDelay) = True Then
        objChrmDriver.findElementByXPath("//button[@id='resultsPage:ExportPopupDialog::ok']").Click
         bElement = True
     End If

      If bElement = False Then
        strMessage = strMessage & ";" & "Click on export popup ok button"
     End If
     
     If ModFunctions.NoOfInvoices(lLastRow) > 100 Then
         'Click on All records
         If objChrmDriver.IsElementPresent(objBy.XPath("//a[@href='javascript:if (resultGrid.canContinue()) submitMyForm(-1) ']"), lDelay) = True Then
             objChrmDriver.findElementByXPath("//a[@href='javascript:if (resultGrid.canContinue()) submitMyForm(-1) ']").Click
             bElement = True
         End If
    
          If bElement = False Then
            strMessage = strMessage & ";" & "Click on export popup ok button"
         End If
    End If
    
    If InStr(1, strInvoices, ",", vbTextCompare) = 0 Then
        
        'Single Shipment ids
        Set objElement = objChrmDriver.findElementByXPath("//input[@value='" & "ULU." & strInvoices & "']")
        objElement.Click
        bElement = True
             
        If bElement = False Then
           strMessage = strMessage & ";" & "Click on items"
        End If
        
        'Call for action operation
        Call ModActionOpration.fnActionOperaton(1, lLastRow, "Single")
        
    Else
         objChrmDriver.Wait 5000
         Set objDictDesc = ModFunctions.fnGetInvoices
         Set objDictAsc = ModFunctions.fnDictAsc
         
         If objDictDesc.Count > 0 And objDictAsc.Count > 0 Then
             
             strBulkId = ""
             'With Amount
             bElement = False
            'Click on Total Pallet count for sorting
            If objChrmDriver.IsElementPresent(objBy.XPath("//div[@id='rgSGSec.1.2.1.21.val']"), lDelay) = True Then
                objChrmDriver.runScript "document.querySelector('.sgScroll').scrollBy(2500,0)"
                objChrmDriver.findElementByXPath("//div[@id='rgSGSec.1.2.1.21.val']").Click
                 bElement = True
             End If
        
             If bElement = False Then
                strMessage = strMessage & ";" & "Click on Total Pallet count for sorting"
             End If
             objChrmDriver.Wait 5000
            For Each vkey In objDictDesc.Keys
                If objDictDesc.Item(vkey) <> Empty Then
                    Set objElement = objChrmDriver.findElementByXPath("//input[@value='" & vkey & "']")
                    objChrmDriver.executeScript "arguments[0].scrollIntoView(true)", objElement
                    objElement.Click
                    bElement = True
                End If
             Next
                
            If bElement = False Then
               strMessage = strMessage & ";" & "Click on Item"
            End If
            'Call for action operation
            Call ModActionOpration.fnActionOperaton(1, lLastRow, "Single")
            'multiple shipments id
            objChrmDriver.Wait 1000
            
           'Empty
            bElement = False
             'Switch frame
            objChrmDriver.Wait 2000
            If objChrmDriver.IsElementPresent(objBy.XPath("//iframe[contains(@id,'mainContentRegion')]"), lDelay) = True Then
                 objChrmDriver.switchToFrame 1
                 bElement = True
             End If

             If bElement = False Then
                strMessage = strMessage & ";" & "Swtich Frame"
             End If
         
              bElement = False
               'Click on All Select button
              If objChrmDriver.IsElementPresent(objBy.XPath("//input[@id='rgSGSec.1.1.1.1.check']"), lDelay) = True Then
                 objChrmDriver.findElementByXPath("//input[@id='rgSGSec.1.1.1.1.check']").Click
                 objChrmDriver.findElementByXPath("//input[@id='rgSGSec.1.1.1.1.check']").Click
                  bElement = True
              End If
    
              If bElement = False Then
                 strMessage = strMessage & ";" & "Click on All select button"
              End If
    
            'Click on Total Pallet count for sorting
             If objChrmDriver.IsElementPresent(objBy.XPath("//div[@id='rgSGSec.1.2.1.21.val']"), lDelay) = True Then
                objChrmDriver.runScript "document.querySelector('.sgScroll').scrollBy(2500,0)"
                objChrmDriver.findElementByXPath("//div[@id='rgSGSec.1.2.1.21.val']").Click
                 bElement = True
             End If
                              
            For Each vkey In objDictAsc.Keys
                If objDictAsc.Item(vkey) = Empty Then
                    Set objElement = objChrmDriver.findElementByXPath("//input[@value='" & vkey & "']")
                    objChrmDriver.executeScript "arguments[0].scrollIntoView(true)", objElement
                    objElement.Click
                    bElement = True
                End If
             Next
          
              If bElement = False Then
                 strMessage = strMessage & ";" & "Click on Item"
              End If
              
              'Call for action operation
              Call ModActionOpration.fnActionOperaton(2, lLastRow, "Multiple")
        
        ElseIf objDictDesc.Count > 0 Then
             strBulkId = ""
             'With Value
             objChrmDriver.Wait 5000
            For Each vkey In objDictDesc.Keys
                If objDictDesc.Item(vkey) <> Empty Then
                    Set objElement = objChrmDriver.findElementByXPath("//input[@value='" & vkey & "']")
                    objChrmDriver.executeScript "arguments[0].scrollIntoView(true)", objElement
                    objElement.Click
                    bElement = True
                End If
             Next
                
            If bElement = False Then
               strMessage = strMessage & ";" & "Click on Item"
            End If
            'Call for action operation
            Call ModActionOpration.fnActionOperaton(1, lLastRow, "Single")
            'multiple shipments id
            objChrmDriver.Wait 10000
        
        ElseIf objDictAsc.Count > 0 Then
             strBulkId = ""
             'With Amount
             bElement = False
            objChrmDriver.Wait 5000
            For Each vkey In objDictAsc.Keys
                If objDictAsc.Item(vkey) = Empty Then
                    Set objElement = objChrmDriver.findElementByXPath("//input[@value='" & vkey & "']")
                    objChrmDriver.executeScript "arguments[0].scrollIntoView(true)", objElement
                    objElement.Click
                    bElement = True
                End If
             Next
                
            If bElement = False Then
               strMessage = strMessage & ";" & "Click on Item"
            End If
            'Call for action operation
            Call ModActionOpration.fnActionOperaton(2, lLastRow, "Single")
            'multiple shipments id
            objChrmDriver.Wait 1000
        End If
     End If
 
     bElement = False
     'Click on Home page
     If objChrmDriver.IsElementPresent(objBy.XPath("//a[@class='HomeIcon xgn p_AFTextOnly']"), lDelay) = True Then
        objChrmDriver.findElementByXPath("//a[@class='HomeIcon xgn p_AFTextOnly']").Click
         bElement = True
     End If

      If bElement = False Then
        strMessage = strMessage & ";" & "Click on Home button"
     End If

     bElement = False
     'Click on Business Automation
     If objChrmDriver.IsElementPresent(objBy.XPath("//img[contains(@id,'mainContentRegion')]"), lDelay) = True Then
        objChrmDriver.findElementsByXPath("//img[contains(@id,'mainContentRegion')]").Item(3).Click
         bElement = True
     End If

      If bElement = False Then
        strMessage = strMessage & ";" & "Click on Business automation"
      End If
     
     objChrmDriver.Wait 2000
      
     bElement = False
     'Click on Reporting
     If objChrmDriver.IsElementPresent(objBy.XPath("//img[contains(@id,'mainContentRegion')]"), lDelay) = True Then
        objChrmDriver.Wait 2000
        objChrmDriver.findElementsByXPath("//img[contains(@id,'mainContentRegion')]").Item(5).Click
        bElement = True
     End If

      If bElement = False Then
        strMessage = strMessage & ";" & "Click on Reporting2"
     End If

     bElement = False
     'Click on report
     If objChrmDriver.IsElementPresent(objBy.XPath("//img[contains(@id,'mainContentRegion')]"), lDelay) = True Then
        objChrmDriver.Wait 5000
        objChrmDriver.findElementsByXPath("//img[contains(@id,'mainContentRegion')]").Item(12).Click
         bElement = True
     End If

      If bElement = False Then
        strMessage = strMessage & ";" & "Click on reporting manager2"
     End If

     bElement = False
      'Switch frame
     If objChrmDriver.IsElementPresent(objBy.XPath("//iframe[contains(@id,'mainContentRegion')]"), lDelay) = True Then
        Set objElement = objChrmDriver.findElementByXPath("//iframe[contains(@id,'mainContentRegion')]")
        objChrmDriver.switchToFrame objElement
         bElement = True
     End If

      If bElement = False Then
        strMessage = strMessage & ";" & "Switch frame for EOD Report_THA_V3"
     End If

     bElement = False
     'Search EOD Rport Tha_V3
     If objChrmDriver.IsElementPresent(objBy.XPath("//td[@class='gridBodyCell']"), lDelay) = True Then
        Set objElements = objChrmDriver.findElementsByXPath("//td[@class='gridBodyCell']")
         bElement = True
     End If

      If bElement = False Then
        strMessage = strMessage & ";" & "searh EOD REPORT_THA-V3"
     End If

     bElement = False
     For Each objElement In objElements
        If objElement.Text = "EOD REPORT_THA _V3" Then
           objElement.findElementByXPath("following::td[3]").Click
            bElement = True
           Exit For
        End If
     Next

      If bElement = False Then
        strMessage = strMessage & ";" & "Click on link EOD REPORT_THA_V3"
     End If

     If InStr(1, strBulkId, ",", vbTextCompare) > 0 Then
          bElement = False
        'click on search
        If objChrmDriver.IsElementPresent(objBy.XPath("//a[@name='P_L_BULK_PLAN_ID_find']"), lDelay) = True Then
           objChrmDriver.findElementByXPath("//a[@name='P_L_BULK_PLAN_ID_find']").Click
            bElement = True
         End If
            
            If bElement = False Then
                strMessage = strMessage & ";" & "Click on Search"
            End If
    
            objChrmDriver.Wait 2000
            objChrmDriver.switchToWindow objChrmDriver.WindowHandles(1)
            objChrmDriver.Wait 1000
            
            bElement = False
            'feeds bulk id
            If objChrmDriver.IsElementPresent(objBy.XPath("//input[@name='bulk_plan/bulk_plan_xid']"), lDelay) = True Then
               objChrmDriver.findElementByXPath("//input[@name='bulk_plan/bulk_plan_xid']").SendKeys strBulkId  '"20200827-0048,20200827-0049"
                bElement = True
            End If
            
            If bElement = False Then
               strMessage = strMessage & ";" & "Enter bulk plan id"
            End If
            
            bElement = False
            'Select drop down item
            If objChrmDriver.IsElementPresent(objBy.XPath("//select[@name='bulk_plan/bulk_plan_xid_operator']"), lDelay) = True Then
                Set objSelect = objChrmDriver.findElementByXPath("//select[@name='bulk_plan/bulk_plan_xid_operator']").AsSelect
                objSelect.SelectByText "One Of"
                bElement = True
            End If
            
            If bElement = False Then
               strMessage = strMessage & ";" & "Select One OF 2"
            End If
            
             bElement = False
             'click on search button
             If objChrmDriver.IsElementPresent(objBy.XPath("//a[@name='search_button']"), lDelay) = True Then
                objChrmDriver.findElementByXPath("//a[@name='search_button']").Click
                 bElement = True
             End If
        
              If bElement = False Then
                strMessage = strMessage & ";" & "Click on search button"
             End If
            
             objChrmDriver.Wait 5000
             objChrmDriver.switchToWindow objChrmDriver.WindowHandles(1)
             objChrmDriver.Wait 1000
            
             bElement = False
             'click on all select
             If objChrmDriver.IsElementPresent(objBy.XPath("//input[@id='rgSGSec.1.1.1.1.check']"), lDelay) = True Then
                objChrmDriver.findElementByXPath("//input[@id='rgSGSec.1.1.1.1.check']").Click
                 bElement = True
             End If
        
              If bElement = False Then
                strMessage = strMessage & ";" & "Click on Select button"
              End If
             
              bElement = False
             'click on finish button
             If objChrmDriver.IsElementPresent(objBy.XPath("//a[@class='enButton']"), lDelay) = True Then
                objChrmDriver.findElementsByXPath("//a[@class='enButton']").Item(0).Click
                 bElement = True
             End If
        
              If bElement = False Then
                strMessage = strMessage & ";" & "Click on search button"
             End If
             
             
     'Switch window
    objChrmDriver.Wait 2000
    objChrmDriver.switchToWindow objChrmDriver.WindowHandles(0)
    objChrmDriver.Wait 1000
       
     bElement = False
       'Switch frame
     If objChrmDriver.IsElementPresent(objBy.XPath("//iframe[contains(@id,'mainContentRegion')]"), lDelay) = True Then
        Set objElement = objChrmDriver.findElementByXPath("//iframe[contains(@id,'mainContentRegion')]")
        objChrmDriver.switchToFrame objElement
         bElement = True
     End If

      If bElement = False Then
        strMessage = strMessage & ";" & "Switch frame for bulk plan id fill"
     End If
    
     
         bElement = False
         'click on submit button
         If objChrmDriver.IsElementPresent(objBy.XPath("//a[@name='submit_report']"), lDelay) = True Then
            objChrmDriver.findElementByXPath("//a[@name='submit_report']").Click
             bElement = True
         End If
    
          If bElement = False Then
            strMessage = strMessage & ";" & "Click on Submit button"
         End If
         
        objChrmDriver.Wait 2000
        objChrmDriver.switchToWindow objChrmDriver.WindowHandles(0)
        
        dtDate = Now()
        While ModFunctions.checkFileExist(dtDate) = False
            DoEvents
        Wend

             
    Else
            bElement = False
            'feeds bulk id
            If objChrmDriver.IsElementPresent(objBy.XPath("//input[@name='P_L_BULK_PLAN_ID']"), lDelay) = True Then
               objChrmDriver.findElementByXPath("//input[@name='P_L_BULK_PLAN_ID']").SendKeys strBulkId
                bElement = True
            End If
            
            If bElement = False Then
               strMessage = strMessage & ";" & "Enter bulk plan id"
            End If
            
             'click on submit button
         If objChrmDriver.IsElementPresent(objBy.XPath("//a[@name='submit_report']"), lDelay) = True Then
            objChrmDriver.findElementByXPath("//a[@name='submit_report']").Click
             bElement = True
         End If
    
          If bElement = False Then
            strMessage = strMessage & ";" & "Click on Submit button"
         End If
         
         objChrmDriver.Wait 2000
         objChrmDriver.switchToWindow objChrmDriver.WindowHandles(0)
        
          'Validation
            dtDate = Now()
            While ModFunctions.checkFileExist(dtDate) = False
                DoEvents
            Wend
      
    End If
        

     bElement = False
     'Click on Home page
     If objChrmDriver.IsElementPresent(objBy.XPath("//a[@class='HomeIcon xgn p_AFTextOnly']"), lDelay) = True Then
        objChrmDriver.findElementByXPath("//a[@class='HomeIcon xgn p_AFTextOnly']").Click
         bElement = True
     End If

      If bElement = False Then
        strMessage = strMessage & ";" & "Click on Home button"
     End If
        
      objChrmDriver.Wait 5000
      bElement = False
     'Click on Business Automation
     If objChrmDriver.IsElementPresent(objBy.XPath("//img[contains(@id,'mainContentRegion')]"), lDelay) = True Then
        objChrmDriver.findElementsByXPath("//img[contains(@id,'mainContentRegion')]").Item(3).Click
         bElement = True
     End If
    
    On Error GoTo 0
     
     If strMessage = "Not Found " Then
        wksInvoices.Range("E" & lLastRow).Value = "Completed"
        wksInvoices.Range("F" & lLastRow).Value = ""
        wksInvoices.Range("C" & lLastRow).Value = strBulkId
        wksInvoices.Range("D" & lLastRow).Value = strStatus
     Else
        wksInvoices.Range("E" & lLastRow).Value = "Not Completed"
        wksInvoices.Range("F" & lLastRow).Value = strMessage
        wksInvoices.Range("C" & lLastRow).Value = strBulkId
        wksInvoices.Range("D" & lLastRow).Value = strStatus
     End If
     
     Exit Function

Err_ElementNotFound:
     wksInvoices.Range("G" & lLastRow).Value = Err.Description
     objChrmDriver.Close
     objChrmDriver.switchToWindow objChrmDriver.WindowHandles(0)
     'Click on Home page
     If objChrmDriver.IsElementPresent(objBy.XPath("//a[@class='HomeIcon xgn p_AFTextOnly']"), lDelay) = True Then
        objChrmDriver.findElementByXPath("//a[@class='HomeIcon xgn p_AFTextOnly']").Click
     End If
End Function

