VERSION 5.00
Begin VB.Form frmCustom 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Custom"
   ClientHeight    =   1215
   ClientLeft      =   1560
   ClientTop       =   2205
   ClientWidth     =   3675
   ControlBox      =   0   'False
   Icon            =   "frmCustom.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   3675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtBombs 
      Height          =   285
      Left            =   1170
      TabIndex        =   6
      Text            =   "0"
      Top             =   780
      Width           =   1095
   End
   Begin VB.TextBox txtColumns 
      Height          =   285
      Left            =   1170
      TabIndex        =   4
      Text            =   "0"
      Top             =   450
      Width           =   1095
   End
   Begin VB.TextBox txtRows 
      Height          =   285
      Left            =   1170
      TabIndex        =   2
      Text            =   "0"
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   525
      Left            =   2340
      TabIndex        =   1
      Top             =   60
      Width           =   1245
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   525
      Left            =   2340
      TabIndex        =   0
      Top             =   630
      Width           =   1245
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Mines:"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   810
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Columns:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Rows:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   150
      Width           =   450
   End
End
Attribute VB_Name = "frmCustom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public IsCancelled As Boolean

Private Sub cmdCancel_Click()
    IsCancelled = True
    Hide
End Sub

Private Sub cmdOK_Click()
    IsCancelled = False
    Hide
End Sub

Private Sub txtBombs_LostFocus()
    If Not IsNumeric(txtBombs) Then
        txtBombs = Int((txtRows * txtColumns) * 0.2)
        Exit Sub
    End If
    If txtBombs > (txtRows * txtColumns) Then
        txtBombs = Int((txtRows * txtColumns) * 0.2)
    ElseIf txtBombs < 10 Then
        txtBombs = 10
    End If
End Sub

Private Sub txtColumns_LostFocus()
    If Not IsNumeric(txtColumns) Then
        txtColumns = 8
        Exit Sub
    End If
    If txtColumns > 30 Then
        txtColumns = 30
    End If
End Sub

Private Sub txtRows_LostFocus()
    If Not IsNumeric(txtRows) Then
        txtRows = 8
        Exit Sub
    End If
    If txtRows > 24 Then
        txtRows = 24
    End If
End Sub
