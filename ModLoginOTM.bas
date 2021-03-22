Attribute VB_Name = "ModLoginOTM"
Option Explicit

''******************************************************************************
'' Procedure           : fnLoginOTM
'' Description         : Login with Credentias in OTM Portal
'' Arguments           : NA
'' Return              : Boolean : if login successfully return False else return True
''******************************************************************************

Public Function fnLoginOTM() As Boolean
    
    Dim strUrl As String
    Dim bFound As Boolean
   
   
    bElement = False
    strMessage = "Not Found "
    lDelay = 60000

    'Set URL
    strUrl = wksCred.Range("B3").Value
    
    objChrmDriver.Start "chrome", strUrl
    
    'Open Url
    objChrmDriver.Get "/"

    'Maximize the window
    objChrmDriver.windowMaximize

     On Error GoTo Err_Field:
     If objChrmDriver.IsElementPresent(objBy.XPath("//input[@name='username']"), lDelay) = True Then
        'User Name
         Set objElement = objChrmDriver.findElementByXPath("//input[@name='username']")
         objElement.Clear
         objElement.SendKeys wksCred.Range("B1").Value
         bElement = True
     End If

     bElement = False
    'Password
     If objChrmDriver.IsElementPresent(objBy.XPath("//input[@name='userpassword']"), lDelay) = True Then
        Set objElement = objChrmDriver.findElementByXPath("//input[@name='userpassword']")
        objElement.Clear
        objElement.SendKeys wksCred.Range("B2").Value
        bElement = True
     End If


     If bElement = False Then
        strMessage = strMessage & ";" & "Password"
     End If

     bElement = False
    
    'Submit
     If objChrmDriver.IsElementPresent(objBy.XPath("//input[@name='submitbutton']"), lDelay) = True Then
        Set objElement = objChrmDriver.findElementByXPath("//input[@name='submitbutton']")
        objElement.Click
        bElement = True
     End If


     If bElement = False Then
        strMessage = strMessage & ";" & "Submit"
     End If
    
    If strMessage = "Not Found " Then
        bFound = True
    Else
        bFound = False
    End If
    
    fnLoginOTM = bFound

Exit Function

Err_Field:
   Set objChrmDriver = Nothing
   bFound = False
   fnLoginOTM = bFound
    
End Function




