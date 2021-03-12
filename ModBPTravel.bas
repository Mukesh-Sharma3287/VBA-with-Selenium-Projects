Attribute VB_Name = "ModBPTravel"
Option Explicit

''******************************************************************************
'' Module Name       : ModMain
'' Description       : This module is used to connect with OTM portal and login with given credentils
''
''
'' Date             Developer             Action            Remarks
'' 5-Jan-2021     Mukesh Sharma          Created
''
''******************************************************************************

''******************************************************************************
'' Procedure           : Main
'' Description         : Perform Shimpent Planing activities in OTM Portal
'' Arguments           : NA
'' Return              : NA
''******************************************************************************

'Pre-rquisite:
'Selenium wrapper should be installed in system
'Chrome Driver version should be same Chrome Browser version
'Add Reference SeleniumWrapper Type Library

Public objChrmDriver As New SeleniumWrapper.WebDriver
Public objElement As SeleniumWrapper.WebElement
Public objBy As New SeleniumWrapper.By
Public objSelect As SeleniumWrapper.Select
Public bPresent As Boolean
Public bElement As Boolean
Public strMessage As String
Public ldelay As Long

Sub Main()
    
    Dim lLastRow As Long
    Dim strInvoices As String
    Dim lCounter As Long
    
    On Error GoTo Err_Found
    lLastRow = wksBPTravel.Cells(Rows.Count, "A").End(xlUp).Row
    
    'Validation
    If lLastRow = 1 Then
        MsgBox "Customer Details are not exist to proceed", vbCritical
        Exit Sub
    End If
    
    'Call function to login the BP Travel
    If fnLoginBPTravel = False Then
        MsgBox "Unable to Login BP Travel, Please try again"
        Exit Sub
    End If
   
    For lCounter = 2 To lLastRow

        'Start Time
        wksBPTravel.Range("Q" & lCounter).Value = Format(Now, "dd-mmm-yyyy h:mm:ss")
        
        'Call for BP travel
         Call fnCreateQuote(lCounter)
        
        'End Time
        wksBPTravel.Range("R" & lCounter).Value = Format(Now, "dd-mmm-yyyy h:mm:ss")

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

''******************************************************************************
'' Procedure           : fnLoginBPTravel
'' Description         : Login with Credentias in BP Travel Portal
'' Arguments           : NA
'' Return              : Boolean : if login successfully return False else return True
''******************************************************************************

Public Function fnLoginBPTravel() As Boolean
    
    Dim strURL As String
    Dim bFound As Boolean
    
    bElement = False
    strMessage = "Not Found "
    ldelay = 30000

    'Set URL
    strURL = "http://bptravel.blueprism.com"

    objChrmDriver.Start "chrome", strURL
    
    'Open Url
    objChrmDriver.get "/"

    'Maximize the window
    objChrmDriver.windowMaximize

     On Error GoTo Err_Field:
     
     'User Name
     If objChrmDriver.isElementPresent(objBy.XPath("//input[@id='username']"), ldelay) = True Then
         Set objElement = objChrmDriver.findElementByCssSelector("#username")
         objElement.Clear
         objElement.SendKeys "admin"
         bElement = True
     End If
        
     If bElement = False Then
        strMessage = strMessage & ";" & "User Name"
     End If
     
     bElement = False
    'Password
     If objChrmDriver.isElementPresent(objBy.XPath("//input[@id='password']"), ldelay) = True Then
        Set objElement = objChrmDriver.findElementByXPath("//input[@id='password']")
        objElement.Clear
        objElement.SendKeys "admin"
        bElement = True
     End If


     If bElement = False Then
        strMessage = strMessage & ";" & "Password"
     End If

     bElement = False
    
    'LogIn
     If objChrmDriver.isElementPresent(objBy.XPath("//span[@onclick='login()']"), ldelay) = True Then
        Set objElement = objChrmDriver.findElementByXPath("//span[@onclick='login()']")
        objElement.Click
        bElement = True
     End If


     If bElement = False Then
        strMessage = strMessage & ";" & "LogIn"
     End If
    
    
    'Check login status
    If strMessage = "Not Found " Then
        bFound = True
    Else
        bFound = False
    End If
    
    fnLoginBPTravel = bFound

Exit Function

Err_Field:
   Set objChrmDriver = Nothing
   bFound = False
   fnLoginBPTravel = bFound
    
End Function

''******************************************************************************
'' Function            : fnCreateQuote
'' Description         : To fill the customer information in Google Chrome Webpage
'' Arguments           : NA
'' Return              : NA
''******************************************************************************

Public Function fnCreateQuote(lCounter As Long)
      
        On Error GoTo Err_Field:
          bElement = False
         'Click on Create Quote
         If objChrmDriver.isElementPresent(objBy.XPath("//a[@href='createquote.html']"), ldelay) = True Then
            Set objElement = objChrmDriver.findElementByXPath("//a[@href='createquote.html']")
            objElement.Click
            bElement = True
         End If
    
         If bElement = False Then
            strMessage = strMessage & ";" & "Click on Create Quote"
         End If
         
        'From
        If objChrmDriver.isElementPresent(objBy.XPath("//select[@id='from']"), ldelay) = True Then
            Set objSelect = objChrmDriver.findElementByXPath("//select[@id='from']").AsSelect
            objSelect.selectByText wksBPTravel.Range("H" & lCounter)
            bElement = True
        End If
        
        If bElement = False Then
           strMessage = strMessage & ";" & "Select From"
        End If
        
        bElement = False
        
        'To
        If objChrmDriver.isElementPresent(objBy.XPath("//select[@id='to']"), ldelay) = True Then
            Set objSelect = objChrmDriver.findElementByXPath("//select[@id='to']").AsSelect
            objSelect.selectByText wksBPTravel.Range("I" & lCounter)
            bElement = True
        End If
        
        If bElement = False Then
           strMessage = strMessage & ";" & "Select To"
        End If
        
        'Departing
        bElement = False
        If objChrmDriver.isElementPresent(objBy.XPath("//input[@id='departing']"), ldelay) = True Then
            Set objElement = objChrmDriver.findElementByXPath("//input[@id='departing']")
            objElement.Clear
            objElement.SendKeys Format(wksBPTravel.Range("J" & lCounter), "dd/mm/yyyy")
            bElement = True
        End If
        
        If bElement = False Then
           strMessage = strMessage & ";" & "Departing"
        End If
        
        If wksBPTravel.Range("K" & lCounter).Value = "" Then
            'Check one way
            bElement = False
            If objChrmDriver.isElementPresent(objBy.XPath("//input[@id='oneway']"), ldelay) = True Then
                 Set objElement = objChrmDriver.findElementByXPath("//input[@id='oneway']")
                 objElement.Click
                 bElement = True
            End If
            
             If bElement = False Then
               strMessage = strMessage & ";" & "Check one way"
             End If
        Else
             'Return
            bElement = False
            If objChrmDriver.isElementPresent(objBy.XPath("//input[@id='returning']"), ldelay) = True Then
                Set objElement = objChrmDriver.findElementByXPath("//input[@id='returning']")
                objElement.Clear
                objElement.SendKeys Format(wksBPTravel.Range("K" & lCounter), "dd/mm/yyyy")
                bElement = True
            End If
            
            If bElement = False Then
               strMessage = strMessage & ";" & "Returning"
            End If
           
        End If
        
        'Adult
        If objChrmDriver.isElementPresent(objBy.XPath("//select[@id='adults']"), ldelay) = True Then
            Set objSelect = objChrmDriver.findElementByXPath("//select[@id='adults']").AsSelect
            objSelect.selectByText wksBPTravel.Range("L" & lCounter)
            bElement = True
        End If
        
        If bElement = False Then
           strMessage = strMessage & ";" & "Select Adult"
        End If
            
       'Children
        If objChrmDriver.isElementPresent(objBy.XPath("//select[@id='children']"), ldelay) = True Then
            Set objSelect = objChrmDriver.findElementByXPath("//select[@id='children']").AsSelect
            objSelect.selectByText wksBPTravel.Range("M" & lCounter)
            bElement = True
        End If
        
        If bElement = False Then
           strMessage = strMessage & ";" & "Select Children"
        End If
            
        'Name
        bElement = False
        If objChrmDriver.isElementPresent(objBy.XPath("//input[@id='name']"), ldelay) = True Then
            Set objElement = objChrmDriver.findElementByXPath("//input[@id='name']")
            objElement.Clear
            objElement.SendKeys wksBPTravel.Range("B" & lCounter) & " " & wksBPTravel.Range("C" & lCounter) & " " & wksBPTravel.Range("D" & lCounter)
            bElement = True
        End If
        
        If bElement = False Then
           strMessage = strMessage & ";" & "Name"
        End If
            
            
        'Postcode
        bElement = False
        If objChrmDriver.isElementPresent(objBy.XPath("//input[@id='postcode']"), ldelay) = True Then
            Set objElement = objChrmDriver.findElementByXPath("//input[@id='postcode']")
            objElement.Clear
            objElement.SendKeys wksBPTravel.Range("G" & lCounter)
            bElement = True
        End If
        
        If bElement = False Then
           strMessage = strMessage & ";" & "Postcode"
        End If
        
        
        'Telephone
        bElement = False
        If objChrmDriver.isElementPresent(objBy.XPath("//input[@id='telephone']"), ldelay) = True Then
            Set objElement = objChrmDriver.findElementByXPath("//input[@id='telephone']")
            objElement.Clear
            objElement.SendKeys wksBPTravel.Range("F" & lCounter)
            bElement = True
        End If
        
        If bElement = False Then
           strMessage = strMessage & ";" & "Telephone"
        End If
        
         'Email
        bElement = False
        If objChrmDriver.isElementPresent(objBy.XPath("//input[@id='email']"), ldelay) = True Then
            Set objElement = objChrmDriver.findElementByXPath("//input[@id='email']")
            objElement.Clear
            objElement.SendKeys wksBPTravel.Range("E" & lCounter)
            bElement = True
        End If
        
        If bElement = False Then
           strMessage = strMessage & ";" & "Email"
        End If
        
        If lCounter = 10 Then
            Debug.Print lCounter
        End If
        
        'Create Quote
        bElement = False
        If objChrmDriver.isElementPresent(objBy.XPath("//span[@class='button' and @onclick='createQuote()']"), ldelay) = True Then
            Set objElement = objChrmDriver.findElementByXPath("//span[@class='button' and @onclick='createQuote()']")
            objElement.Click
            bElement = True
        End If
        
        If bElement = False Then
           strMessage = strMessage & ";" & "Create Quote"
        End If
      On Error GoTo 0
      
      'Status of Data entry completion in Google Chrome Webpage
     If strMessage = "Not Found " Then
        wksBPTravel.Range("N" & lCounter).Value = "Completed"
     Else
        wksBPTravel.Range("N" & lCounter).Value = "Not Completed"
        wksBPTravel.Range("P" & lCounter).Value = strMessage
    
     End If
     
     Exit Function
Err_Field:
     
     wksBPTravel.Range("N" & lCounter).Value = "Not Completed"
     wksBPTravel.Range("O" & lCounter).Value = Err.Description
      
End Function


