Attribute VB_Name = "ModActionOpration"
Option Explicit

Public Function fnActionOperaton(iCount As Integer, lLastRow As Long, strPalletStatus As String)
    
    'Click on Actions button
     bElement = False
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
     'Click on operation
     If objChrmDriver.IsElementPresent(objBy.XPath("//tr[@id='actionTree.1_1']"), lDelay) = True Then
        objChrmDriver.findElementByXPath("//tr[@id='actionTree.1_1']").Click
         bElement = True
     End If

    If bElement = False Then
        strMessage = strMessage & ";" & "Click on Operation"
     End If

     bElement = False
     If objChrmDriver.IsElementPresent(objBy.XPath("//span[@id='actionTree.1_1_1.l']"), lDelay) = True Then
        objChrmDriver.findElementByXPath("//span[@id='actionTree.1_1_1.l']").Click
         bElement = True
     End If

      If bElement = False Then
        strMessage = strMessage & ";" & "Click on created by shipment plan"
     End If
   
     bElement = False
     If objChrmDriver.IsElementPresent(objBy.XPath("//a[@id='actionTree.1_1_1_4.k']"), lDelay) = True Then
        objChrmDriver.findElementByXPath("//a[@id='actionTree.1_1_1_4.k']").Click
         bElement = True
     End If

      If bElement = False Then
        strMessage = strMessage & ";" & "Click on bulk plan"
      End If
    
     objChrmDriver.switchToWindow objChrmDriver.WindowHandles(1)
    
     bElement = False
     'Switch frame
     If objChrmDriver.IsElementPresent(objBy.XPath("//frame[@name='mainBody']"), lDelay) = True Then
        Set objElement = objChrmDriver.findElementByXPath("//frame[@name='mainBody']")
        objChrmDriver.switchToFrame objElement
         bElement = True
     End If

      If bElement = False Then
        strMessage = strMessage & ";" & "Switch frame of Shipment planning"
     End If
   
     bElement = False
     If objChrmDriver.IsElementPresent(objBy.XPath("//select[@name='qualifier/xid@ID']"), lDelay) = True Then
        Set objSelect = objChrmDriver.findElementByXPath("//select[@name='qualifier/xid@ID']").AsSelect
        If iCount = 1 Then
            objSelect.SelectByText "ULU_THA_PLAN"
        Else
             objSelect.SelectByText "ULU_THA_PLAN_ERU_OFF"
        End If
         bElement = True
     End If

      If bElement = False Then
        strMessage = strMessage & ";" & "Select dropdown ULU_THA_PLAN"
     End If

     bElement = False
     'ok button
      If objChrmDriver.IsElementPresent(objBy.XPath("//a[@name='ok']"), lDelay) = True Then
        'Click on release management
        objChrmDriver.findElementByXPath("//a[@name='ok']").Click
         bElement = True
     End If

     If bElement = False Then
        strMessage = strMessage & ";" & "Click on ok button"
     End If


    objChrmDriver.Wait 5000
    objChrmDriver.switchToWindow objChrmDriver.WindowHandles(1)
    objChrmDriver.Wait 1000

     bElement = False
       'Switch frame
     If objChrmDriver.IsElementPresent(objBy.XPath("//frame[@name='mainBody']"), lDelay) = True Then
        Set objElement = objChrmDriver.findElementByXPath("//frame[@name='mainBody']")
        objChrmDriver.switchToFrame objElement
         bElement = True
     End If

     If bElement = False Then
        strMessage = strMessage & ";" & "Switch frame of window bulk plan"
     End If

     bElement = False
     'Refresh page
     If objChrmDriver.IsElementPresent(objBy.XPath("//a[@name='refreshButton']"), lDelay) = True Then
         objChrmDriver.findElementByXPath("//a[@name='refreshButton']").Click
         bElement = True
     End If
     
     objChrmDriver.Wait 2000
     If bElement = False Then
        strMessage = strMessage & ";" & "Refresh page"
     End If
    
     bElement = False
    'get first table data
     If objChrmDriver.IsElementPresent(objBy.XPath("//div[@class='fieldLabel']"), lDelay) = True Then
        Set objElements = objChrmDriver.findElementsByXPath("//div[@class='fieldLabel']")
         bElement = True
     End If

      If bElement = False Then
        strMessage = strMessage & ";" & "get value bulk id"
     End If
   strStatus = "RUNNING"
   While strStatus = "RUNNING"
        'get first table data
        If objChrmDriver.IsElementPresent(objBy.XPath("//div[@class='fieldLabel']"), lDelay) = True Then
           Set objElements = objChrmDriver.findElementsByXPath("//div[@class='fieldLabel']")
            bElement = True
        End If
        For Each objElement In objElements
            If objElement.Text = "Status" Then
                objChrmDriver.Wait 100
                strStatus = objElement.findElementByXPath("following::div").Text
                If InStr(1, strStatus, "RUNNING", vbTextCompare) > 0 Then
                   If objChrmDriver.IsElementPresent(objBy.XPath("//a[@name='refreshButton']"), lDelay) = True Then
                       objChrmDriver.Wait 5000
                       objChrmDriver.findElementByXPath("//a[@name='refreshButton']").Click
                   End If
                   Exit For
                End If
                
               If strStatus = "COMPLETED" Then
                    Exit For
              End If
              
          End If
       Next
    Wend
    
     bElement = False
    'get first table data
     If objChrmDriver.IsElementPresent(objBy.XPath("//div[@class='fieldLabel']"), lDelay) = True Then
        Set objElements = objChrmDriver.findElementsByXPath("//div[@class='fieldLabel']")
         bElement = True
     End If

     For Each objElement In objElements
        If objElement.Text = "Bulk Plan ID" Then
            objChrmDriver.Wait 100
             If strPalletStatus = "Single" Then
                 strBulkId = objElement.findElementByXPath("following::div").Text
              Else
                strBulkId = strBulkId & "," & objElement.findElementByXPath("following::div").Text
              End If
              bElement = True
            bElement = True
        End If
   
        If objElement.Text = "Status" Then
             objChrmDriver.Wait 100
             strStatus = objElement.findElementByXPath("following::div").Text
       End If
     Next
   
   
       If bElement = False Then
           strMessage = strMessage & ";" & "bulk plain id :" & strBulkId & " Status: " & strStatus
        End If
       
        objChrmDriver.Wait 2000
        objChrmDriver.Close
        objChrmDriver.switchToWindow objChrmDriver.WindowHandles(0)
        
        If Left(strBulkId, 1) = "," Then
            strBulkId = Right(strBulkId, Len(strBulkId) - 1)
        End If
        wksInvoices.Range("C" & lLastRow).Value = strBulkId
        wksInvoices.Range("D" & lLastRow).Value = strStatus

End Function

