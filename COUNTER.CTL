VERSION 5.00
Begin VB.UserControl Counter 
   CanGetFocus     =   0   'False
   ClientHeight    =   825
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1380
   FillStyle       =   0  'Solid
   ScaleHeight     =   825
   ScaleWidth      =   1380
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   780
      Top             =   150
   End
   Begin VB.Image imgCounter 
      Height          =   345
      Index           =   0
      Left            =   120
      Top             =   120
      Width           =   195
   End
End
Attribute VB_Name = "Counter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'********************************************************************
'Name:              Counter
'Created:           13-Jun-2000
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

Private Enum Counters
    Counter0 = 101
    Counter1 = 102
    Counter2 = 103
    Counter3 = 104
    Counter4 = 105
    Counter5 = 106
    Counter6 = 107
    Counter7 = 108
    Counter8 = 109
    Counter9 = 110
    CounterNeg = 111
End Enum

Const m_TimerMin = 0
Const m_TimerMax = 999

'Default Property Values:
Const m_def_ReverseTimer = False
Const m_def_TimerEnabled = 0
Const m_def_Value = 0

'Property Variables:
Dim m_ReverseTimer As Boolean
Dim m_TimerEnabled As Boolean
Dim m_Value        As Integer

Public Enum TimerEventsEnum
    TimerMinOut
    TimerMaxOut
End Enum

Public Event TimerEvent(TimerEvents As TimerEventsEnum)
'*******************************************************************************
' Timer1_Timer (SUB)
'
' PARAMETERS:
' None
'
' DESCRIPTION:
' ***Description goes here***
'*******************************************************************************
Private Sub Timer1_Timer()
    If ReverseTimer Then
        Value = m_Value - 1
        If Value = m_TimerMin Then
            TimerEnabled = False
            RaiseEvent TimerEvent(TimerMinOut)
        End If
    Else
        Value = m_Value + 1
        If Value = m_TimerMax Then
            TimerEnabled = False
            RaiseEvent TimerEvent(TimerMaxOut)
        End If
    End If
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
    Load imgCounter(1)
    Load imgCounter(2)
    
    Value = 0
    
    imgCounter(1).Visible = True
    imgCounter(2).Visible = True
    
    imgCounter(0).Move Screen.TwipsPerPixelX, Screen.TwipsPerPixelY
    imgCounter(1).Move 14 * Screen.TwipsPerPixelX, Screen.TwipsPerPixelY
    imgCounter(2).Move 27 * Screen.TwipsPerPixelX, Screen.TwipsPerPixelY
End Sub

'*******************************************************************************
' UserControl_Paint (SUB)
'
' PARAMETERS:
' None
'
' DESCRIPTION:
' Draw border
'*******************************************************************************
Private Sub UserControl_Paint()
    UserControl.Cls
    UserControl.Line (0, 0)-(40 * Screen.TwipsPerPixelX, 0), QBColor(8)
    UserControl.Line (0, 0)-(0, 24 * Screen.TwipsPerPixelY), QBColor(8)
    UserControl.Line (40 * Screen.TwipsPerPixelX, Screen.TwipsPerPixelY)-(40 * Screen.TwipsPerPixelX, 24 * Screen.TwipsPerPixelY), RGB(255, 255, 255) 'QBColor(1)
    UserControl.Line (Screen.TwipsPerPixelX, 24 * Screen.TwipsPerPixelY)-(41 * Screen.TwipsPerPixelX, 24 * Screen.TwipsPerPixelY), RGB(255, 255, 255) 'QBColor(1)
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
    UserControl.Width = 41 * Screen.TwipsPerPixelX
    UserControl.Height = 25 * Screen.TwipsPerPixelY
End Sub

'*******************************************************************************
' Value (PROPERTY GET)
'*******************************************************************************
Public Property Get Value() As Integer
    Value = m_Value
End Property

'*******************************************************************************
' Value (PROPERTY LET)
'*******************************************************************************
Public Property Let Value(ByVal New_Value As Integer)
Attribute Value.VB_Description = "Value being displayed"
Dim tNumber     As String
Dim vNumbers    As Variant

    m_Value = New_Value
    PropertyChanged "Value"
    'Update the display
    tNumber = Format$(m_Value, "0\,0\,0")
    vNumbers = Split(tNumber, ",")
    LoadNumber 0, CInt(vNumbers(0))
    LoadNumber 1, CInt(vNumbers(1))
    LoadNumber 2, CInt(vNumbers(2))
    
    If m_Value < 0 Then
        LoadNumber 0, CInt(-1)
    End If
End Property

'*******************************************************************************
' UserControl_InitProperties (SUB)
'
' PARAMETERS:
' None
'
' DESCRIPTION:
' ***Description goes here***
'*******************************************************************************
Private Sub UserControl_InitProperties()
'Initialize Properties for User Control
    Value = m_def_Value
    m_TimerEnabled = m_def_TimerEnabled
    m_ReverseTimer = m_def_ReverseTimer
End Sub

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
'Load property values from storage
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    m_TimerEnabled = PropBag.ReadProperty("TimerEnabled", m_def_TimerEnabled)
    m_ReverseTimer = PropBag.ReadProperty("ReverseTimer", m_def_ReverseTimer)
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
'Write property values to storage
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("TimerEnabled", m_TimerEnabled, m_def_TimerEnabled)
    Call PropBag.WriteProperty("ReverseTimer", m_ReverseTimer, m_def_ReverseTimer)
End Sub

'*******************************************************************************
' LoadNumber (SUB)
'
' PARAMETERS:
' (In/Out) - ImageIndex - Integer -
' (In/Out) - Number     - Integer -
'
' DESCRIPTION:
' Load the number bitmap from the resource file
'*******************************************************************************
Private Sub LoadNumber(ImageIndex As Integer, Number As Integer)
Dim iCounter    As Integer

    Select Case Number
    Case -1
        iCounter = CounterNeg
    Case 0
        iCounter = Counter0
    Case 1
        iCounter = Counter1
    Case 2
        iCounter = Counter2
    Case 3
        iCounter = Counter3
    Case 4
        iCounter = Counter4
    Case 5
        iCounter = Counter5
    Case 6
        iCounter = Counter6
    Case 7
        iCounter = Counter7
    Case 8
        iCounter = Counter8
    Case 9
        iCounter = Counter9
    End Select
    imgCounter(ImageIndex).Picture = LoadResPicture(iCounter, vbResBitmap)
End Sub

'*******************************************************************************
' TimerEnabled (PROPERTY GET)
'*******************************************************************************
Public Property Get TimerEnabled() As Boolean
Attribute TimerEnabled.VB_Description = "If true starts a timer"
    TimerEnabled = m_TimerEnabled
End Property

'*******************************************************************************
' TimerEnabled (PROPERTY LET)
'*******************************************************************************
Public Property Let TimerEnabled(ByVal New_TimerEnabled As Boolean)
    m_TimerEnabled = New_TimerEnabled
    PropertyChanged "TimerEnabled"
    Timer1.Enabled = m_TimerEnabled
End Property

'*******************************************************************************
' ReverseTimer (PROPERTY GET)
'*******************************************************************************
Public Property Get ReverseTimer() As Boolean
    ReverseTimer = m_ReverseTimer
End Property

'*******************************************************************************
' ReverseTimer (PROPERTY LET)
'*******************************************************************************
Public Property Let ReverseTimer(ByVal New_ReverseTimer As Boolean)
    m_ReverseTimer = New_ReverseTimer
    PropertyChanged "ReverseTimer"
End Property

