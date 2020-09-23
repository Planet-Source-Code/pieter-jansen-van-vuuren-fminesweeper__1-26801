VERSION 5.00
Begin VB.Form frmMineSweeper 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Minesweeper"
   ClientHeight    =   4365
   ClientLeft      =   3105
   ClientTop       =   2700
   ClientWidth     =   5370
   Icon            =   "frmMineSweeper.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   5370
   Begin MineSweeper.MainPart MainPart1 
      Height          =   2835
      Left            =   90
      Top             =   90
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   5001
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Begin VB.Menu mnuGameNew 
         Caption         =   "&New"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuGameSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameBeginner 
         Caption         =   "&Beginner"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuGameIntermediate 
         Caption         =   "&Intermediate"
      End
      Begin VB.Menu mnuGameExpert 
         Caption         =   "&Expert"
      End
      Begin VB.Menu mnuGameCustom 
         Caption         =   "&Custom..."
      End
      Begin VB.Menu mnuGameSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameHint 
         Caption         =   "&Hint"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuGameMarks 
         Caption         =   "&Marks (?)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuGameSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameBestTimes 
         Caption         =   "Best &Times"
      End
      Begin VB.Menu mnuGameSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About Minesweeper..."
      End
   End
End
Attribute VB_Name = "frmMineSweeper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'Name:              frmMineSweeper
'Created:           15-Jun-2000
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
'Date:              15-Jun-2000
'Description:       Description of changes made
'--------------------------------------------------------------------
'********************************************************************
Option Explicit

Private m_oReg As CRegSettings

'*******************************************************************************
' Form_KeyDown (SUB)
'
' PARAMETERS:
' (In/Out) - KeyCode - Integer -
' (In/Out) - Shift   - Integer -
'
' DESCRIPTION:
' ***Description goes here***
'*******************************************************************************
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Screen.MousePointer = vbHourglass
    If KeyCode = vbKeyEscape Then
        Me.WindowState = vbMinimized
    'ElseIf KeyCode = vbKeyF2 Then
    '    mnuGameNew_Click
    End If
    'Screen.MousePointer = vbDefault
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
    Screen.MousePointer = vbHourglass
    MainPart1.Move 0, 0
    Me.Move Me.Left, Me.Top, MainPart1.Width + (Me.Width - Me.ScaleWidth), MainPart1.Height + (Me.Height - Me.ScaleHeight)
    Set m_oReg = New CRegSettings
    
    MainPart1.BuildGrid m_oReg.GetSetting("LastGrid", "Rows", "8"), _
        m_oReg.GetSetting("LastGrid", "Cols", "8"), _
        m_oReg.GetSetting("LastGrid", "Bombs", "10")

    mnuGameBeginner.Checked = False
    mnuGameIntermediate.Checked = False
    mnuGameExpert.Checked = False
    mnuGameCustom.Checked = False
    Select Case m_oReg.GetSetting("LastGrid", "Type", "Beginner")
    Case "Beginner"
        mnuGameBeginner.Checked = True
    Case "Intermediate"
        mnuGameIntermediate.Checked = True
    Case "Expert"
        mnuGameExpert.Checked = True
    Case "Custom"
        mnuGameCustom.Checked = True
    End Select
    Screen.MousePointer = vbDefault
End Sub

'*******************************************************************************
' Form_Paint (SUB)
'
' PARAMETERS:
' None
'
' DESCRIPTION:
' ***Description goes here***
'*******************************************************************************
Private Sub Form_Paint()
    Me.Move Me.Left, Me.Top, MainPart1.Width + (Me.Width - Me.ScaleWidth), MainPart1.Height + (Me.Height - Me.ScaleHeight)
End Sub

'*******************************************************************************
' Form_Unload (SUB)
'
' PARAMETERS:
' (In/Out) - Cancel - Integer -
'
' DESCRIPTION:
' ***Description goes here***
'*******************************************************************************
Private Sub Form_Unload(Cancel As Integer)
    With m_oReg
        .SaveSetting "LastGrid", "Rows", MainPart1.Rows
        .SaveSetting "LastGrid", "Cols", MainPart1.Cols
        .SaveSetting "LastGrid", "Bombs", MainPart1.Bombs
    End With 'm_oReg
    Set m_oReg = Nothing
End Sub

'*******************************************************************************
' MainPart1_GridEvent (SUB)
'
' PARAMETERS:
' (In/Out) - TileID     - Long           -
' (In/Out) - GridEvents - GridEventsEnum -
'
' DESCRIPTION:
' ***Description goes here***
'*******************************************************************************
Private Sub MainPart1_GridEvent(TileID As Long, GridEvents As GridEventsEnum)
    Select Case GridEvents
    Case geNew
    Case geBombed
    Case geDone
        GetBestTimes
    Case geExposed
    Case geFlagged
    Case geQuestion
    Case geUnQuestion
    End Select
End Sub

'*******************************************************************************
' MainPart1_Resize (SUB)
'
' PARAMETERS:
' None
'
' DESCRIPTION:
' ***Description goes here***
'*******************************************************************************
Private Sub MainPart1_Resize()
    Me.Move Me.Left, Me.Top, MainPart1.Width + (Me.Width - Me.ScaleWidth), MainPart1.Height + (Me.Height - Me.ScaleHeight)
End Sub

'*******************************************************************************
' mnuGameBeginner_Click (SUB)
'
' PARAMETERS:
' None
'
' DESCRIPTION:
' ***Description goes here***
'*******************************************************************************
Private Sub mnuGameBeginner_Click()
    Screen.MousePointer = vbHourglass
    m_oReg.SaveSetting "LastGrid", "Type", "Beginner"
    mnuGameBeginner.Checked = True
    mnuGameIntermediate.Checked = False
    mnuGameExpert.Checked = False
    mnuGameCustom.Checked = False
    MainPart1.BuildGrid 8, 8, 10
    Screen.MousePointer = vbDefault
End Sub

'*******************************************************************************
' mnuGameBestTimes_Click (SUB)
'
' PARAMETERS:
' None
'
' DESCRIPTION:
' ***Description goes here***
'*******************************************************************************
Private Sub mnuGameBestTimes_Click()
    ShowBestTimes
End Sub

'*******************************************************************************
' mnuGameCustom_Click (SUB)
'
' PARAMETERS:
' None
'
' DESCRIPTION:
' ***Description goes here***
'*******************************************************************************
Private Sub mnuGameCustom_Click()
    With frmCustom
        .txtRows = MainPart1.Rows
        .txtColumns = MainPart1.Cols
        .txtBombs = MainPart1.Bombs
        .Show vbModal, Me
        If .IsCancelled Then
            Set frmCustom = Nothing
            Exit Sub
        End If
        MainPart1.Rows = .txtRows
        MainPart1.Cols = .txtColumns
        MainPart1.Bombs = .txtBombs
    End With
    Set frmCustom = Nothing
    Screen.MousePointer = vbHourglass
    m_oReg.SaveSetting "LastGrid", "Type", "Custom"
    mnuGameBeginner.Checked = False
    mnuGameIntermediate.Checked = False
    mnuGameExpert.Checked = False
    mnuGameCustom.Checked = True
    MainPart1.BuildGrid
    Screen.MousePointer = vbDefault
End Sub

'*******************************************************************************
' mnuGameExit_Click (SUB)
'
' PARAMETERS:
' None
'
' DESCRIPTION:
' ***Description goes here***
'*******************************************************************************
Private Sub mnuGameExit_Click()
    End
End Sub

'*******************************************************************************
' mnuGameExpert_Click (SUB)
'
' PARAMETERS:
' None
'
' DESCRIPTION:
' ***Description goes here***
'*******************************************************************************
Private Sub mnuGameExpert_Click()
    Screen.MousePointer = vbHourglass
    m_oReg.SaveSetting "LastGrid", "Type", "Expert"
    mnuGameBeginner.Checked = False
    mnuGameIntermediate.Checked = False
    mnuGameExpert.Checked = True
    mnuGameCustom.Checked = False
    MainPart1.BuildGrid 16, 30, 99
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuGameHint_Click()
    MainPart1.Hint False
End Sub

'*******************************************************************************
' mnuGameIntermediate_Click (SUB)
'
' PARAMETERS:
' None
'
' DESCRIPTION:
' ***Description goes here***
'*******************************************************************************
Private Sub mnuGameIntermediate_Click()
    Screen.MousePointer = vbHourglass
    m_oReg.SaveSetting "LastGrid", "Type", "Intermediate"
    mnuGameBeginner.Checked = False
    mnuGameIntermediate.Checked = True
    mnuGameExpert.Checked = False
    mnuGameCustom.Checked = False
    MainPart1.BuildGrid 16, 16, 40
    Screen.MousePointer = vbDefault
End Sub

'*******************************************************************************
' mnuGameNew_Click (SUB)
'
' PARAMETERS:
' None
'
' DESCRIPTION:
' ***Description goes here***
'*******************************************************************************
Private Sub mnuGameNew_Click()
    Screen.MousePointer = vbHourglass
    MainPart1.SmilyFace = SmilyDown
    DoEvents
    MainPart1.BuildGrid
    MainPart1.SmilyFace = SmilyUp
    Screen.MousePointer = vbDefault
End Sub

'*******************************************************************************
' GetBestTimes (SUB)
'
' PARAMETERS:
' None
'
' DESCRIPTION:
' ***Description goes here***
'*******************************************************************************
Private Sub GetBestTimes()
Dim tBestTimes      As String
Dim vBestTimes      As Variant

Dim vTime           As Variant

Dim i               As Integer

Dim iInsertionPoint As Integer

Dim tUser           As String

    tBestTimes = m_oReg.GetSetting("BestTimes" & MainPart1.Rows & MainPart1.Cols & MainPart1.Bombs, "Top 10", "999,Anonymous~998,Anonymous~996,Anonymous~997,Anonymous~994,Anonymous~992,Anonymous~995,Anonymous~993,Anonymous~991,Anonymous~990,Anonymous")
    
    vBestTimes = Split(tBestTimes, "~")
    
    InsertSortStringsStart vBestTimes, True, True
    
    vTime = Split(vBestTimes(UBound(vBestTimes)), ",")
    If MainPart1.TimerValue < CInt(vTime(0)) Then
        'This is a new Top ten Time
        'Find Insertion point
        For i = UBound(vBestTimes) To LBound(vBestTimes) Step -1
            vTime = Split(vBestTimes(i), ",")
            If MainPart1.TimerValue < CInt(vTime(0)) Then
                tUser = vTime(1)
                iInsertionPoint = i
            End If
        Next 'i
        'Move other times down
        For i = UBound(vBestTimes) To iInsertionPoint Step -1
            If i > LBound(vBestTimes) Then
                vBestTimes(i) = vBestTimes(i - 1)
            End If
        Next 'i
        vBestTimes(iInsertionPoint) = Format(MainPart1.TimerValue, "@@@") & "," & InputBox("Enter your name", "Top 10 Score", tUser)
    End If
    
    tBestTimes = Join(vBestTimes, "~")
    
    m_oReg.SaveSetting "BestTimes" & MainPart1.Rows & MainPart1.Cols & MainPart1.Bombs, "Top 10", tBestTimes
    
End Sub

'*******************************************************************************
' ShowBestTimes (SUB)
'
' PARAMETERS:
' None
'
' DESCRIPTION:
' ***Description goes here***
'*******************************************************************************
Private Sub ShowBestTimes()
Dim tBestTimes As String
Dim vBestTimes As Variant

Dim vTime      As Variant

Dim i          As Integer

    tBestTimes = m_oReg.GetSetting("BestTimes" & MainPart1.Rows & MainPart1.Cols & MainPart1.Bombs, "Top 10", "999,Anonymous~998,Anonymous~996,Anonymous~997,Anonymous~994,Anonymous~992,Anonymous~995,Anonymous~993,Anonymous~991,Anonymous~990,Anonymous")
    
    vBestTimes = Split(tBestTimes, "~")
    
    InsertSortStringsStart vBestTimes, True, False
    
    For i = LBound(vBestTimes) To UBound(vBestTimes)
        vTime = Split(vBestTimes(i), ",")
        frmBestTimes.List1.AddItem CStr(i + 1) & "." & vbTab & vTime(0) & vbTab & vTime(1)
    Next 'i

    frmBestTimes.Show vbModal, Me
    Set frmBestTimes = Nothing
End Sub

