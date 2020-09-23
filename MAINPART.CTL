VERSION 5.00
Begin VB.UserControl MainPart 
   CanGetFocus     =   0   'False
   ClientHeight    =   5520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6675
   ScaleHeight     =   5520
   ScaleWidth      =   6675
   Begin MineSweeper.Grid Grid1 
      Height          =   2010
      Left            =   210
      Top             =   840
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   3545
   End
   Begin MineSweeper.TopPart TopPart1 
      Height          =   555
      Left            =   210
      Top             =   120
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   979
   End
End
Attribute VB_Name = "MainPart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'********************************************************************
'Name:              MainPart
'Created:           14-Jun-2000
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

Public Event Resize()

Private m_TwipsPerPixelX As Long
Private m_TwipsPerPixelY As Long

Public Event TimerEvent(TimerEvents As TimerEventsEnum)
'Event Declarations:
Event GridEvent(TileID As Long, GridEvents As GridEventsEnum) 'MappingInfo=Grid1,Grid1,-1,GridEvent


'*******************************************************************************
' Rows (PROPERTY GET)
'*******************************************************************************
Public Property Get Rows() As Long
    Rows = Grid1.Rows
End Property
Public Property Let Rows(New_Rows As Long)
    Grid1.Rows = New_Rows
End Property
'*******************************************************************************
' Cols (PROPERTY GET)
'*******************************************************************************
Public Property Get Cols() As Long
    Cols = Grid1.Cols
End Property
Public Property Let Cols(New_Cols As Long)
    Grid1.Cols = New_Cols
End Property
'*******************************************************************************
' Bombs (PROPERTY GET)
'*******************************************************************************
Public Property Get Bombs() As Long
    Bombs = Grid1.Bombs
End Property
Public Property Let Bombs(New_Bombs As Long)
    Grid1.Bombs = New_Bombs
End Property

'*******************************************************************************
' Grid1_GridEvent (SUB)
'
' PARAMETERS:
' (In/Out) - TileID     - Long           -
' (In/Out) - GridEvents - GridEventsEnum -
'
' DESCRIPTION:
' ***Description goes here***
'*******************************************************************************
Private Sub Grid1_GridEvent(TileID As Long, GridEvents As GridEventsEnum)
    TopPart1.BombCount = Grid1.Bombs - Grid1.TilesFlagged
    Select Case GridEvents
    Case geNew
    Case geBombed
        TopPart1.SmilyFace = SmilyLost
        TopPart1.PauseTimer
    Case geDone
        TopPart1.SmilyFace = SmilyWon
        TopPart1.PauseTimer
    Case geExposed
    Case geFlagged
    Case geQuestion
    Case geUnQuestion
    End Select
    RaiseEvent GridEvent(TileID, GridEvents)
End Sub

'*******************************************************************************
' Grid1_MouseDown (SUB)
'
' PARAMETERS:
' (In/Out) - TileID - Long    -
' (In/Out) - Button - Integer -
' (In/Out) - Shift  - Integer -
' (In/Out) - X      - Single  -
' (In/Out) - Y      - Single  -
'
' DESCRIPTION:
' ***Description goes here***
'*******************************************************************************
Private Sub Grid1_MouseDown(TileID As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    TopPart1.SmilyFace = SmilyO
    If Grid1.Enabled _
    And Not TopPart1.TimerEnabled Then
        TopPart1.StartTimer
    End If
End Sub

'*******************************************************************************
' Grid1_MouseUp (SUB)
'
' PARAMETERS:
' (In/Out) - TileID - Long    -
' (In/Out) - Button - Integer -
' (In/Out) - Shift  - Integer -
' (In/Out) - X      - Single  -
' (In/Out) - Y      - Single  -
'
' DESCRIPTION:
' ***Description goes here***
'*******************************************************************************
Private Sub Grid1_MouseUp(TileID As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Grid1.Enabled Then
        TopPart1.SmilyFace = SmilyUp
    End If
End Sub

'*******************************************************************************
' Grid1_Resize (SUB)
'
' PARAMETERS:
' None
'
' DESCRIPTION:
' ***Description goes here***
'*******************************************************************************
Private Sub Grid1_Resize()
    TopPart1.Width = Grid1.Width
    UserControl_Resize
    RaiseEvent Resize
End Sub

'*******************************************************************************
' TopPart1_Click (SUB)
'
' PARAMETERS:
' None
'
' DESCRIPTION:
' ***Description goes here***
'*******************************************************************************
Private Sub TopPart1_Click()
    With TopPart1
        .BombCount = Grid1.Bombs
        .PauseTimer
        .ResetTimer
    End With
    Grid1.DrawGrid
End Sub

'*******************************************************************************
' TopPart1_TimerEvent (SUB)
'
' PARAMETERS:
' (In/Out) - TimerEvents - TimerEventsEnum -
'
' DESCRIPTION:
' ***Description goes here***
'*******************************************************************************
Private Sub TopPart1_TimerEvent(TimerEvents As TimerEventsEnum)
    RaiseEvent TimerEvent(TimerEvents)
End Sub

'*******************************************************************************
' UserControl_Initialize (SUB)
'
' PARAMETERS:
' None
'
' DESCRIPTION:
' ***Description goes here***
'*******************************************************************************
Private Sub UserControl_Initialize()
    m_TwipsPerPixelY = Screen.TwipsPerPixelY
    m_TwipsPerPixelX = Screen.TwipsPerPixelX
    TopPart1.Move 6 * m_TwipsPerPixelX, 6 * m_TwipsPerPixelY, Grid1.Width
    Grid1.Move (6 * m_TwipsPerPixelX), (12 * m_TwipsPerPixelY) + TopPart1.Height
    UserControl.Height = TopPart1.Height + Grid1.Height + (18 * m_TwipsPerPixelY)
    UserControl.Width = TopPart1.Width + (12 * m_TwipsPerPixelX)
    TopPart1.BombCount = Grid1.Bombs
    TopPart1.ResetTimer
End Sub

'*******************************************************************************
' UserControl_Paint (SUB)
'
' PARAMETERS:
' None
'
' DESCRIPTION:
' ***Description goes here***
'*******************************************************************************
Private Sub UserControl_Paint()
    
    m_TwipsPerPixelY = Screen.TwipsPerPixelY
    m_TwipsPerPixelX = Screen.TwipsPerPixelX
    
    UserControl.Cls
    
    UserControl.Line (0, 0)-(UserControl.Width, 0), RGB(255, 255, 255)
    UserControl.Line (0, m_TwipsPerPixelY)-(UserControl.Width - m_TwipsPerPixelX, m_TwipsPerPixelY), RGB(255, 255, 255)
    UserControl.Line (0, m_TwipsPerPixelY + m_TwipsPerPixelY)-(UserControl.Width - m_TwipsPerPixelX - m_TwipsPerPixelX, m_TwipsPerPixelY + m_TwipsPerPixelY), RGB(255, 255, 255)
    
    UserControl.Line (0, 0)-(0, UserControl.Height), RGB(255, 255, 255)
    UserControl.Line (m_TwipsPerPixelX, 0)-(m_TwipsPerPixelX, UserControl.Height - m_TwipsPerPixelY), RGB(255, 255, 255)
    UserControl.Line (m_TwipsPerPixelX + m_TwipsPerPixelX, 0)-(m_TwipsPerPixelX + m_TwipsPerPixelX, UserControl.Height - m_TwipsPerPixelY - m_TwipsPerPixelY), RGB(255, 255, 255)
    
    UserControl.Line (m_TwipsPerPixelX, UserControl.Height - m_TwipsPerPixelY)-(UserControl.Width, UserControl.Height - m_TwipsPerPixelY), QBColor(8)
    UserControl.Line (m_TwipsPerPixelX + m_TwipsPerPixelX, UserControl.Height - (m_TwipsPerPixelY + m_TwipsPerPixelY))-(UserControl.Width, UserControl.Height - (m_TwipsPerPixelY + m_TwipsPerPixelY)), QBColor(8)
    UserControl.Line (m_TwipsPerPixelX + m_TwipsPerPixelX + m_TwipsPerPixelX, UserControl.Height - m_TwipsPerPixelY - m_TwipsPerPixelY - m_TwipsPerPixelY)-(UserControl.Width - m_TwipsPerPixelX, UserControl.Height - (m_TwipsPerPixelY + m_TwipsPerPixelY + m_TwipsPerPixelY)), QBColor(8)
    
    UserControl.Line (UserControl.Width - m_TwipsPerPixelX, m_TwipsPerPixelY)-(UserControl.Width - m_TwipsPerPixelX, UserControl.Height), QBColor(8)
    UserControl.Line (UserControl.Width - m_TwipsPerPixelX - m_TwipsPerPixelX, m_TwipsPerPixelY + m_TwipsPerPixelY)-(UserControl.Width - m_TwipsPerPixelX - m_TwipsPerPixelX, UserControl.Height), QBColor(8)
    UserControl.Line (UserControl.Width - m_TwipsPerPixelX - m_TwipsPerPixelX - m_TwipsPerPixelX, m_TwipsPerPixelY + m_TwipsPerPixelY + m_TwipsPerPixelY)-(UserControl.Width - m_TwipsPerPixelX - m_TwipsPerPixelX - m_TwipsPerPixelX, UserControl.Height), QBColor(8)
    
End Sub

'*******************************************************************************
' UserControl_Resize (SUB)
'
' PARAMETERS:
' None
'
' DESCRIPTION:
' ***Description goes here***
'*******************************************************************************
Private Sub UserControl_Resize()
    If UserControl.Height <> TopPart1.Height + Grid1.Height + (18 * m_TwipsPerPixelY) Then
        UserControl.Height = TopPart1.Height + Grid1.Height + (18 * m_TwipsPerPixelY)
    End If
    If UserControl.Width <> TopPart1.Width + (12 * m_TwipsPerPixelX) Then
        UserControl.Width = TopPart1.Width + (12 * m_TwipsPerPixelX)
    End If
End Sub

'*******************************************************************************
' BuildGrid (SUB)
'
' PARAMETERS:
' (In/Out) - Rows - Variant -
' (In/Out) - Rows - Cols    -
' (In/Out) - Rows - Bombs   -
'
' DESCRIPTION:
' ***Description goes here***
'*******************************************************************************
Public Sub BuildGrid(Optional Rows, Optional Cols, Optional Bombs)
    With Grid1
        If Not IsMissing(Rows) Then
            .Rows = CLng(Rows)
        End If
        If Not IsMissing(Cols) Then
            .Cols = CLng(Cols)
        End If
        If Not IsMissing(Bombs) Then
            .Bombs = CLng(Bombs)
        End If
        .DrawGrid
        TopPart1.BombCount = .Bombs
    End With 'Grid1
    With TopPart1
        .PauseTimer
        .ResetTimer
        .SmilyFace = SmilyUp
    End With 'TopPart1
End Sub

'*******************************************************************************
' StopTimer (SUB)
'
' PARAMETERS:
' None
'
' DESCRIPTION:
' ***Description goes here***
'*******************************************************************************
Public Sub StopTimer()
    TopPart1.PauseTimer
End Sub
'*******************************************************************************
' StartTimer (SUB)
'
' PARAMETERS:
' None
'
' DESCRIPTION:
' ***Description goes here***
'*******************************************************************************
Public Sub StartTimer()
    TopPart1.StartTimer
End Sub

'*******************************************************************************
' ReverseTimer (PROPERTY GET)
'*******************************************************************************
Public Property Get ReverseTimer() As Boolean
    ReverseTimer = TopPart1.ReverseTimer
End Property

'*******************************************************************************
' ReverseTimer (PROPERTY LET)
'*******************************************************************************
Public Property Let ReverseTimer(ByVal New_ReverseTimer As Boolean)
    TopPart1.ReverseTimer() = New_ReverseTimer
    PropertyChanged "ReverseTimer"
End Property

'*******************************************************************************
' UserControl_ReadProperties (SUB)
'
' PARAMETERS:
' (In/Out) - PropBag - PropertyBag -
'
' DESCRIPTION:
' ***Description goes here***
'*******************************************************************************
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    TopPart1.ReverseTimer = PropBag.ReadProperty("ReverseTimer", 0)
    TopPart1.TimerValue = PropBag.ReadProperty("TimerValue", 0)
    TopPart1.TimerEnabled = PropBag.ReadProperty("TimerEnabled", False)
    TopPart1.SmilyFace = PropBag.ReadProperty("SmilyFace", SmilyUp)
End Sub

'*******************************************************************************
' UserControl_WriteProperties (SUB)
'
' PARAMETERS:
' (In/Out) - PropBag - PropertyBag -
'
' DESCRIPTION:
' ***Description goes here***
'*******************************************************************************
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("ReverseTimer", TopPart1.ReverseTimer, 0)
    Call PropBag.WriteProperty("TimerValue", TopPart1.TimerValue, 0)
    Call PropBag.WriteProperty("TimerEnabled", TopPart1.TimerEnabled, False)
    Call PropBag.WriteProperty("SmilyFace", SmilyUp, SmilyUp)
End Sub

'*******************************************************************************
' TimerValue (PROPERTY GET)
'*******************************************************************************
Public Property Get TimerValue() As Long
    TimerValue = TopPart1.TimerValue
End Property

'*******************************************************************************
' TimerValue (PROPERTY LET)
'*******************************************************************************
Public Property Let TimerValue(ByVal New_TimerValue As Long)
    TopPart1.TimerValue() = New_TimerValue
    PropertyChanged "TimerValue"
End Property

'*******************************************************************************
' TimerEnabled (PROPERTY GET)
'*******************************************************************************
Public Property Get TimerEnabled() As Boolean
    TimerEnabled = TopPart1.TimerEnabled
End Property

'*******************************************************************************
' TimerEnabled (PROPERTY LET)
'*******************************************************************************
Public Property Let TimerEnabled(ByVal New_TimerEnabled As Boolean)
    TopPart1.TimerEnabled() = New_TimerEnabled
    PropertyChanged "TimerEnabled"
End Property

Public Property Let SmilyFace(ByVal New_SmilyFace As SmilyEnum)
    TopPart1.SmilyFace() = New_SmilyFace
    PropertyChanged "SmilyFace"
End Property

Public Sub Hint(Optional DoHint As Boolean = False)
    Grid1.Hint DoHint
    TopPart1.IncrementTimer 1
End Sub

