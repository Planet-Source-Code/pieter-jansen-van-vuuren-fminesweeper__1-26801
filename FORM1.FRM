VERSION 5.00
Begin VB.Form frmMineSweeper 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Minesweeper"
   ClientHeight    =   4365
   ClientLeft      =   4260
   ClientTop       =   3075
   ClientWidth     =   5370
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   Begin MineSweeper.MainPart MainPart1 
      Height          =   2835
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   5001
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMinesweeper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    'Counter1.Value = Counter1.Value + 1
End Sub

Private Sub Form_Load()
    'Counter1.TimerEnabled = True
    MainPart1.Move 0, 0
    Me.Move Me.Left, Me.Top, MainPart1.Width + (Me.Width - Me.ScaleWidth), MainPart1.Height + (Me.Height - Me.ScaleHeight)
    MainPart1.BuildGrid 16, 30, 99
End Sub

Private Sub Grid1_GridEvent(TileID As Long, GridEvents As GridEventsEnum)
    Debug.Print TileID, GridEvents
End Sub

Private Sub MainPart1_Resize()
    Me.Move Me.Left, Me.Top, MainPart1.Width + (Me.Width - Me.ScaleWidth), MainPart1.Height + (Me.Height - Me.ScaleHeight)
End Sub

Private Sub mnuFileExit_Click()
    End
End Sub
