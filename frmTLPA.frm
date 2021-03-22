VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTLPA 
   Caption         =   "Thailand Logistic Process Automation"
   ClientHeight    =   4290
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10215
   OleObjectBlob   =   "frmTLPA.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTLPA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
'' Procedure           : cmdProcess_Click
'' Description         : Process with Selection option
'' Arguments           : NA
'' Return              : Scripting.Dictionary: return invoices as dictionary
''******************************************************************************
Private Sub cmdProcess_Click()
    
    'Validation to Process
    If optJoin.Value = False And optApproval.Value = False And optShipment.Value = False Then
        MsgBox "Please Select at least one option to proceed", vbInformation
        Exit Sub
    End If
    
    'Process
    If optJoin.Value = True Then
        Call ModFunctions.TextJoin
    ElseIf optApproval.Value = True Then
        Call ModMainOTM.MainApproval
    ElseIf optShipment.Value = True Then
        Call ModMainOTM.Main
    End If
    
    
End Sub
