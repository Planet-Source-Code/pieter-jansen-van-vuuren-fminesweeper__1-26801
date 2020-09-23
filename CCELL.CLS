VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'********************************************************************
'Name:              cCell
'Created:           06-Jun-2000
'Description:       Short description of what it does
'Copyright:         Copyright 2000 Pieter van Vuuren. All Rights Reserved.
'
'Dependant On:
'
'Used By:
'
'Changes
'--------------------------------------------------------------------
'Developer:         Pieter
'Date:              06-Jun-2000
'Description:       Description of changes made
'--------------------------------------------------------------------
'********************************************************************
Option Explicit

'My ID
Private m_ID          As Integer

'Row and Col Location
Private m_Row         As Integer
Private m_Col         As Integer

'Buddies
Private m_N           As Integer
Private m_NE          As Integer
Private m_E           As Integer
Private m_SE          As Integer
Private m_S           As Integer
Private m_SW          As Integer
Private m_W           As Integer
Private m_NW          As Integer

'I've got a bomb
Private m_fBomb       As Boolean

Private m_fExposed    As Boolean
Private m_fFlagged    As Boolean
Private m_fQuestioned As Boolean

'*******************************************************************************
' Questioned (PROPERTY LET)
'*******************************************************************************
Public Property Let Questioned(fQuestioned As Boolean)
    m_fQuestioned = fQuestioned
End Property
'*******************************************************************************
' Questioned (PROPERTY GET)
'*******************************************************************************
Public Property Get Questioned() As Boolean
    Questioned = m_fQuestioned
End Property

'*******************************************************************************
' Flagged (PROPERTY LET)
'*******************************************************************************
Public Property Let Flagged(fFlagged As Boolean)
    m_fFlagged = fFlagged
End Property
'*******************************************************************************
' Flagged (PROPERTY GET)
'*******************************************************************************
Public Property Get Flagged() As Boolean
    Flagged = m_fFlagged
End Property

'*******************************************************************************
' Exposed (PROPERTY LET)
'*******************************************************************************
Public Property Let Exposed(fExposed As Boolean)
    m_fExposed = fExposed
End Property
'*******************************************************************************
' Exposed (PROPERTY GET)
'*******************************************************************************
Public Property Get Exposed() As Boolean
    Exposed = m_fExposed
End Property

'*******************************************************************************
' Bomb (PROPERTY LET)
'*******************************************************************************
Public Property Let Bomb(fBomb As Boolean)
    m_fBomb = fBomb
End Property
'*******************************************************************************
' Bomb (PROPERTY GET)
'*******************************************************************************
Public Property Get Bomb() As Boolean
    Bomb = m_fBomb
End Property

'*******************************************************************************
' ID (PROPERTY LET)
'*******************************************************************************
Public Property Let ID(lID As Integer)
    m_ID = lID
End Property
'*******************************************************************************
' ID (PROPERTY GET)
'*******************************************************************************
Public Property Get ID() As Integer
    ID = m_ID
End Property

'*******************************************************************************
' Row (PROPERTY LET)
'*******************************************************************************
Public Property Let Row(lRow As Integer)
    m_Row = lRow
End Property
'*******************************************************************************
' Row (PROPERTY GET)
'*******************************************************************************
Public Property Get Row() As Integer
    Row = m_Row
End Property

'*******************************************************************************
' Col (PROPERTY LET)
'*******************************************************************************
Public Property Let Col(lCol As Integer)
    m_Col = lCol
End Property
'*******************************************************************************
' Col (PROPERTY GET)
'*******************************************************************************
Public Property Get Col() As Integer
    Col = m_Col
End Property

'*******************************************************************************
' N (PROPERTY LET)
'*******************************************************************************
Public Property Let N(lN As Integer)
    m_N = lN
End Property
'*******************************************************************************
' N (PROPERTY GET)
'*******************************************************************************
Public Property Get N() As Integer
    N = m_N
End Property

'*******************************************************************************
' NE (PROPERTY LET)
'*******************************************************************************
Public Property Let NE(lNE As Integer)
    m_NE = lNE
End Property
'*******************************************************************************
' NE (PROPERTY GET)
'*******************************************************************************
Public Property Get NE() As Integer
    NE = m_NE
End Property

'*******************************************************************************
' E (PROPERTY LET)
'*******************************************************************************
Public Property Let E(lE As Integer)
    m_E = lE
End Property
'*******************************************************************************
' E (PROPERTY GET)
'*******************************************************************************
Public Property Get E() As Integer
    E = m_E
End Property

'*******************************************************************************
' SE (PROPERTY LET)
'*******************************************************************************
Public Property Let SE(lSE As Integer)
    m_SE = lSE
End Property
'*******************************************************************************
' SE (PROPERTY GET)
'*******************************************************************************
Public Property Get SE() As Integer
    SE = m_SE
End Property

'*******************************************************************************
' S (PROPERTY LET)
'*******************************************************************************
Public Property Let S(lS As Integer)
    m_S = lS
End Property
'*******************************************************************************
' S (PROPERTY GET)
'*******************************************************************************
Public Property Get S() As Integer
    S = m_S
End Property

'*******************************************************************************
' SW (PROPERTY LET)
'*******************************************************************************
Public Property Let SW(lSW As Integer)
    m_SW = lSW
End Property
'*******************************************************************************
' SW (PROPERTY GET)
'*******************************************************************************
Public Property Get SW() As Integer
    SW = m_SW
End Property

'*******************************************************************************
' W (PROPERTY LET)
'*******************************************************************************
Public Property Let W(lW As Integer)
    m_W = lW
End Property
'*******************************************************************************
' W (PROPERTY GET)
'*******************************************************************************
Public Property Get W() As Integer
    W = m_W
End Property

'*******************************************************************************
' NW (PROPERTY LET)
'*******************************************************************************
Public Property Let NW(lNW As Integer)
    m_NW = lNW
End Property
'*******************************************************************************
' NW (PROPERTY GET)
'*******************************************************************************
Public Property Get NW() As Integer
    NW = m_NW
End Property

'*******************************************************************************
' Class_Initialize (SUB)
'
' PARAMETERS:
' None
'
' DESCRIPTION:
' ***Description goes here***
'*******************************************************************************
Private Sub Class_Initialize()
    m_fQuestioned = False
    m_fFlagged = False
    m_fExposed = False
    m_fBomb = False
End Sub
