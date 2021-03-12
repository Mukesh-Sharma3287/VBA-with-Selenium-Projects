Attribute VB_Name = "MChallenge"
Option Explicit

''******************************************************************************
'' Module Name       : MChallenge
'' Description       : This module is used to fill the customer information in Google Chrome Webpage
''
''
'' Date             Developer             Action            Remarks
'' 12-Jan-2021     Mukesh Sharma          Created
''
''******************************************************************************

''******************************************************************************
'' Procedure           : Main
'' Description         : Fill the customer information in Google Chrome Webpage
'' Arguments           : NA
'' Return              : NA
''******************************************************************************

'Pre-rquisite:
'Selenium wrapper should be installed in system
'Chrome Driver version should be same Chrome Browser Version
'Add Reference SeleniumWrapper Type Library

Public objChrmDriver As New SeleniumWrapper.WebDriver
Public objElement As SeleniumWrapper.WebElement
Public objBy As New SeleniumWrapper.By
Public bPresent As Boolean
Public bElement As Boolean
Public strMessage As String
Public ldelay As Long

Sub Main()
    
    'Declare variables
    Dim strUrl As String
    Dim lLastRow As Long
    Dim lCounter As Long
    
    ldelay = 30000
    
    'Set URL
    strUrl = "http://www.rpachallenge.com/"
    
    'Open Url
    objChrmDriver.Start "chrome", strUrl
    
    objChrmDriver.get "/"
    
    'Maximize the window
     objChrmDriver.windowMaximize
    
    'Count Last row
    lLastRow = wksChallenge.Cells(Rows.Count, "A").End(xlUp).Row
    
    'Click on Start
     Set objElement = objChrmDriver.findElementByXPath("//button[@class='waves-effect col s12 m12 l12 btn-large uiColorButton']")
     objElement.Click
     
    'Loop through feed the employee details into web form
     On Error GoTo Err_Field:
     For lCounter = 2 To lLastRow
        
        Call fnFillRPAChallengeWebForm(lCounter)
        
     Next
    On Error GoTo 0
    
    MsgBox "Done", vbInformation
    Set objChrmDriver = Nothing
    
    Exit Sub
Err_Field:
     
     MsgBox Err.Description, vbCritical
     Set objChrmDriver = Nothing
 
End Sub

''******************************************************************************
'' Function            : fnFillRPAChallengeWebForm
'' Description         : To fill the customer information in Google Chrome Webpage
'' Arguments           : NA
'' Return              : NA
''******************************************************************************

Public Function fnFillRPAChallengeWebForm(lCounter As Long)
      
        On Error GoTo Err_Field:
        
        strMessage = "Not Found "
        
        bElement = False
        If objChrmDriver.isElementPresent(objBy.XPath("//input[@ng-reflect-name='labelFirstName']"), ldelay) = True Then
            Set objElement = objChrmDriver.findElementByXPath("//input[@ng-reflect-name='labelFirstName']")
            objElement.Clear
            objElement.SendKeys wksChallenge.Range("A" & lCounter).Value
            bElement = True
        End If
    
        If bElement = False Then
           strMessage = strMessage & ";" & "First Name"
        End If
        
        'Feed Last Name
        bElement = False
        If objChrmDriver.isElementPresent(objBy.XPath("//input[@ng-reflect-name='labelLastName']"), ldelay) = True Then
             Set objElement = objChrmDriver.findElementByXPath("//input[@ng-reflect-name='labelLastName']")
             objElement.Clear
             objElement.SendKeys wksChallenge.Range("B" & lCounter).Value
             bElement = True
        End If
    
        If bElement = False Then
           strMessage = strMessage & ";" & "Last Name"
        End If
        
      
      'Feed Company Name
       bElement = False
       If objChrmDriver.isElementPresent(objBy.XPath("//input[@ng-reflect-name='labelCompanyName']"), ldelay) = True Then
            Set objElement = objChrmDriver.findElementByXPath("//input[@ng-reflect-name='labelCompanyName']")
            objElement.Clear
            objElement.SendKeys wksChallenge.Range("C" & lCounter).Value
            bElement = True
        End If
    
        If bElement = False Then
           strMessage = strMessage & ";" & "Company Name"
        End If
            
      'Feed Role in Company
       bElement = False
       If objChrmDriver.isElementPresent(objBy.XPath("//input[@ng-reflect-name='labelRole']"), ldelay) = True Then
            Set objElement = objChrmDriver.findElementByXPath("//input[@ng-reflect-name='labelRole']")
            objElement.Clear
            objElement.SendKeys wksChallenge.Range("D" & lCounter).Value
            bElement = True
        End If
    
        If bElement = False Then
           strMessage = strMessage & ";" & "Role in Company"
        End If
        
      
      'Feed Address
       bElement = False
       If objChrmDriver.isElementPresent(objBy.XPath("//input[@ng-reflect-name='labelAddress']"), ldelay) = True Then
            Set objElement = objChrmDriver.findElementByXPath("//input[@ng-reflect-name='labelAddress']")
            objElement.Clear
            objElement.SendKeys wksChallenge.Range("E" & lCounter).Value
            bElement = True
        End If
    
        If bElement = False Then
           strMessage = strMessage & ";" & "Address"
        End If
            
      'Feed Email
       bElement = False
       If objChrmDriver.isElementPresent(objBy.XPath("//input[@ng-reflect-name='labelEmail']"), ldelay) = True Then
            Set objElement = objChrmDriver.findElementByXPath("//input[@ng-reflect-name='labelEmail']")
            objElement.Clear
            objElement.SendKeys wksChallenge.Range("F" & lCounter).Value
            bElement = True
        End If
    
        If bElement = False Then
           strMessage = strMessage & ";" & "Email"
        End If
      
      'Feed Phone no
       bElement = False
       If objChrmDriver.isElementPresent(objBy.XPath("//input[@ng-reflect-name='labelPhone']"), ldelay) = True Then
            Set objElement = objChrmDriver.findElementByXPath("//input[@ng-reflect-name='labelPhone']")
            objElement.Clear
            objElement.SendKeys wksChallenge.Range("G" & lCounter).Value
            bElement = True
       End If
    
       If bElement = False Then
           strMessage = strMessage & ";" & "Phone no"
       End If

      'Click on submit button
      bElement = False
      If objChrmDriver.isElementPresent(objBy.XPath("//input[@ng-reflect-name='labelPhone']"), ldelay) = True Then
            Set objElement = objChrmDriver.findElementByXPath("//input[@class='btn uiColorButton']")
            objElement.Click
            bElement = True
       End If
    
       If bElement = False Then
           strMessage = strMessage & ";" & "Click on submit button"
       End If
      On Error GoTo 0
      
      'Status of Data entry completion in Google Chrome Webpage
     If strMessage = "Not Found " Then
        wksChallenge.Range("H" & lCounter).Value = "Completed"
     Else
        wksChallenge.Range("H" & lCounter).Value = "Not Completed"
        wksChallenge.Range("J" & lCounter).Value = strMessage
    
     End If
     
     Exit Function
Err_Field:
     wksChallenge.Range("H" & lCounter).Value = "Not Completed"
     wksChallenge.Range("I" & lCounter).Value = Err.Description
      
End Function
