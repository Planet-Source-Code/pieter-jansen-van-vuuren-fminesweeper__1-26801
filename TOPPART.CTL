VERSION 5.00
Begin VB.UserControl TopPart 
   CanGetFocus     =   0   'False
   ClientHeight    =   885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6330
   ScaleHeight     =   885
   ScaleWidth      =   6330
   Begin MineSweeper.Counter cntTimer 
      Height          =   375
      Left            =   5070
      Top             =   180
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
   End
   Begin MineSweeper.Counter cntBombs 
      Height          =   375
      Left            =   240
      Top             =   180
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      Value           =   10
   End
   Begin VB.Image imgSmily 
      Height          =   525
      Left            =   2370
      Top             =   150
      Width           =   1245
   End
End
Attribute VB_Name = "TopPart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'********************************************************************
'Name:              TopPart
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

Public Event Click()

Private m_TwipsPerPixelX As Long
Private m_TwipsPerPixelY As Long

Public Enum SmilyEnum
    SmilyUp = 401
    SmilyDown = 402
    SmilyO = 403
    SmilyWon = 404
    SmilyLost = 405
End Enum

Public Event TimerEvent(TimerEvents As TimerEventsEnum)

'*******************************************************************************
' BombCount (PROPERTY LET)
'*******************************************************************************
Public Property Let BombCount(New_BombCount As Long)
    cntBombs.Value = New_BombCount
End Property

'*******************************************************************************
' BombCount (PROPERTY GET)
'*******************************************************************************
Public Property Get BombCount() As Long
    BombCount = cntBombs.Value
End Property

'*******************************************************************************
' TimerEnabled (PROPERTY LET)
'*******************************************************************************
Public Property Let TimerEnabled(New_TimerEnabled As Boolean)
    cntTimer.TimerEnabled = New_TimerEnabled
End Property
'*******************************************************************************
' TimerEnabled (PROPERTY GET)
'*******************************************************************************
Public Property Get TimerEnabled() As Boolean
    TimerEnabled = cntTimer.TimerEnabled
End Property

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
    cntTimer.TimerEnabled = True
End Sub
'*******************************************************************************
' PauseTimer (SUB)
'
' PARAMETERS:
' None
'
' DESCRIPTION:
' ***Description goes here***
'*******************************************************************************
Public Sub PauseTimer()
    cntTimer.TimerEnabled = False
End Sub
'*******************************************************************************
' ResetTimer (SUB)
'
' PARAMETERS:
' None
'
' DESCRIPTION:
' ***Description goes here***
'*******************************************************************************
Public Sub ResetTimer()
    cntTimer.Value = 0
End Sub

'*******************************************************************************
' TimerValue (PROPERTY LET)
'*******************************************************************************
Public Property Let TimerValue(New_TimerValue As Long)
    cntTimer.Value = New_TimerValue
End Property

'*******************************************************************************
' TimerValue (PROPERTY GET)
'*******************************************************************************
Public Property Get TimerValue() As Long
    TimerValue = cntTimer.Value
End Property

'*******************************************************************************
' SmilyFace (PROPERTY LET)
'*******************************************************************************
Public Property Let SmilyFace(New_SmilyFace As SmilyEnum)
    imgSmily.Picture = LoadResPicture(New_SmilyFace, vbResBitmap)
End Property

Private Sub cntTimer_TimerEvent(TimerEvents As TimerEventsEnum)
    RaiseEvent TimerEvent(TimerEvents)
End Sub

'*******************************************************************************
' imgSmily_Click (SUB)
'
' PARAMETERS:
' None
'
' DESCRIPTION:
' ***Description goes here***
'*******************************************************************************
Private Sub imgSmily_Click()
    RaiseEvent Click
End Sub

'*******************************************************************************
' imgSmily_MouseDown (SUB)
'
' PARAMETERS:
' (In/Out) - Button - Integer -
' (In/Out) - Shift  - Integer -
' (In/Out) - X      - Single  -
' (In/Out) - Y      - Single  -
'
' DESCRIPTION:
' ***Description goes here***
'*******************************************************************************
Private Sub imgSmily_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SmilyFace = SmilyDown
End Sub

'*******************************************************************************
' imgSmily_MouseUp (SUB)
'
' PARAMETERS:
' (In/Out) - Button - Integer -
' (In/Out) - Shift  - Integer -
' (In/Out) - X      - Single  -
' (In/Out) - Y      - Single  -
'
' DESCRIPTION:
' ***Description goes here***
'*******************************************************************************
Private Sub imgSmily_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SmilyFace = SmilyUp
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
    imgSmily.Move 26 * m_TwipsPerPixelX, 26 * m_TwipsPerPixelY, 26 * m_TwipsPerPixelX, 26 * m_TwipsPerPixelY
    SmilyFace = SmilyUp
    cntBombs.Move 6 * m_TwipsPerPixelX, 6 * m_TwipsPerPixelY
    cntTimer.Move UserControl.Width - 6 * m_TwipsPerPixelX, 6 * m_TwipsPerPixelY
    UserControl.Height = (12 * m_TwipsPerPixelY) + cntBombs.Height
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
    
    UserControl.Line (0, 0)-(UserControl.Width, 0), QBColor(8)
    UserControl.Line (0, m_TwipsPerPixelY)-(UserControl.Width - m_TwipsPerPixelX, m_TwipsPerPixelY), QBColor(8)
    
    UserControl.Line (0, 0)-(0, UserControl.Height), QBColor(8)
    UserControl.Line (m_TwipsPerPixelX, 0)-(m_TwipsPerPixelX, UserControl.Height - m_TwipsPerPixelY), QBColor(8)
    
    UserControl.Line (m_TwipsPerPixelX, UserControl.Height - m_TwipsPerPixelY)-(UserControl.Width, UserControl.Height - m_TwipsPerPixelY), RGB(255, 255, 255)
    UserControl.Line (m_TwipsPerPixelX + m_TwipsPerPixelX, UserControl.Height - m_TwipsPerPixelY - m_TwipsPerPixelY)-(UserControl.Width, UserControl.Height - m_TwipsPerPixelY - m_TwipsPerPixelY), RGB(255, 255, 255)
    
    UserControl.Line (UserControl.Width - m_TwipsPerPixelX, m_TwipsPerPixelY)-(UserControl.Width - m_TwipsPerPixelX, UserControl.Height), RGB(255, 255, 255)
    UserControl.Line (UserControl.Width - m_TwipsPerPixelX - m_TwipsPerPixelX, m_TwipsPerPixelY + m_TwipsPerPixelY)-(UserControl.Width - m_TwipsPerPixelX - m_TwipsPerPixelX, UserControl.Height), RGB(255, 255, 255)

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
    If UserControl.Height <> (12 * m_TwipsPerPixelY) + cntBombs.Height Then
        UserControl.Height = (12 * m_TwipsPerPixelY) + cntBombs.Height
    End If
    cntBombs.Move 6 * m_TwipsPerPixelX, 6 * m_TwipsPerPixelY
    cntTimer.Move UserControl.Width - (6 * m_TwipsPerPixelX) - cntTimer.Width, 6 * m_TwipsPerPixelY
    imgSmily.Move UserControl.Width / 2 - imgSmily.Width / 2, UserControl.Height / 2 - imgSmily.Height / 2
End Sub
Public Property Get ReverseTimer() As Boolean
    ReverseTimer = cntTimer.ReverseTimer
End Property

Public Property Let ReverseTimer(ByVal New_ReverseTimer As Boolean)
    cntTimer.ReverseTimer() = New_ReverseTimer
    PropertyChanged "ReverseTimer"
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    cntTimer.ReverseTimer = PropBag.ReadProperty("ReverseTimer", 0)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("ReverseTimer", cntTimer.ReverseTimer, 0)
End Sub

Public Sub IncrementTimer(IncrementBy As Long)
    cntTimer.Value = cntTimer.Value + IncrementBy
End Sub
