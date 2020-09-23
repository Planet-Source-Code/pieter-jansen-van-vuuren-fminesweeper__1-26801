VERSION 5.00
Begin VB.Form frmBestTimes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Best Times"
   ClientHeight    =   2805
   ClientLeft      =   1605
   ClientTop       =   1935
   ClientWidth     =   3150
   Icon            =   "frmBestTimes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   3150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3135
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   525
      Left            =   1890
      TabIndex        =   0
      Top             =   2250
      Width           =   1245
   End
End
Attribute VB_Name = "frmBestTimes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'Name:              frmBestTimes
'Created:           19-Jun-2000
'Description:       Short description of what it does
'Copyright:         Copyright 2000 Pieter van Vuuren. All Rights Reserved.
'
'Dependant On:
'
'Used By:
'
'Changes
'--------------------------------------------------------------------
'Developer:         Pieter van Vuuren
'Date:              21-Jun-2000
'Description:       Description of changes made
'--------------------------------------------------------------------
'********************************************************************
Option Explicit

' general declarations code (form)
Private Const LB_SETTABSTOPS As Long = &H192

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
  (ByVal hwnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long

'*******************************************************************************
' cmdOK_Click (SUB)
'
' PARAMETERS:
' None
'
' DESCRIPTION:
' ***Description goes here***
'*******************************************************************************
Private Sub cmdOK_Click()
    Unload Me
End Sub

'*******************************************************************************
' Form_Load (SUB)
'
' PARAMETERS:
' None
'
' DESCRIPTION:
' ***Description goes here***
'*******************************************************************************
Private Sub Form_Load()
    ReDim TabArray(0 To 2) As Long
    
    TabArray(0) = 0
    TabArray(1) = 21
    TabArray(2) = 60
    
    'clear any existing tabs
    Call SendMessage(List1.hwnd, LB_SETTABSTOPS, 0&, ByVal 0&)
    
    'set the list tabstops
    Call SendMessage(List1.hwnd, LB_SETTABSTOPS, 3&, TabArray(0))
    List1.Refresh
End Sub
