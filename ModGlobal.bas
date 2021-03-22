Attribute VB_Name = "ModGlobal"
Option Explicit

''******************************************************************************
'' Module Name       : ModGlobal
'' Description       : This module contains the global variables
''
''
'' Date             Developer             Action            Remarks
'' 25-Feb-2020     Mukesh Sharma          Created
''
''******************************************************************************

   'Declare variables
    Public objChrmDriver As New SeleniumWrapper.WebDriver
    Public objElement As SeleniumWrapper.WebElement
    Public objElements As SeleniumWrapper.WebElementCollection
    Public objSelect As SeleniumWrapper.Select
    Public objBy As New SeleniumWrapper.by
    Public objTable As SeleniumWrapper.Table
    Public lCounter As Long
    Public lDelay As Long
    Public bElement As Boolean
    Public bPresent As Boolean
    Public strBulkId As String
    Public strStatus As String
    Public strMessage As String
   
