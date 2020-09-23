VERSION 5.00
Begin VB.UserControl Grid 
   CanGetFocus     =   0   'False
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Image imgCell 
      Height          =   375
      Index           =   0
      Left            =   390
      Top             =   210
      Visible         =   0   'False
      Width           =   435
   End
End
Attribute VB_Name = "Grid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'********************************************************************
'Name:              Grid
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

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Event GridEvent(TileID As Long, GridEvents As GridEventsEnum)
Public Event MouseDown(TileID As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(TileID As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event Resize()
Public Event Click()

Public Enum GridEventsEnum
    geNew
    geBombed
    geFlagged
    geQuestion
    geUnQuestion
    geExposed
    geDone
End Enum

Private Enum BombCounters
    BombCounter0 = 202
    BombCounter1 = 301
    BombCounter2 = 302
    BombCounter3 = 303
    BombCounter4 = 304
    BombCounter5 = 305
    BombCounter6 = 306
    BombCounter7 = 307
    BombCounter8 = 308
End Enum

Private Enum Tiles
    Tile = 201
    EmptyTile = 202
    Bomb = 203
    ExplodedBomb = 204
    Flag = 205
    FlagWrong = 206
    Question = 207
    QuestionDown = 208
End Enum

Private m_TwipsPerPixelX  As Long
Private m_TwipsPerPixelY  As Long

Private m_Height          As Long
Private m_Width           As Long

Private m_TilesFlagged    As Long
Private m_TilesQuestioned As Long
Private m_TilesExposed    As Long

Private m_oGrid           As cGrid

Private m_lShift          As Long
Private m_fLeftButton     As Boolean

'Default Property Values:
Const m_def_Bombs = 10
Const m_def_Rows = 8
Const m_def_Cols = 8

'Property Variables:
Dim m_Bombs               As Long
Dim m_Rows                As Long
Dim m_Cols                As Long

'*******************************************************************************
' TilesFlagged (PROPERTY GET)
'*******************************************************************************
Public Property Get TilesFlagged() As Long
    TilesFlagged = m_TilesFlagged
End Property

'*******************************************************************************
' TilesQuestioned (PROPERTY GET)
'*******************************************************************************
Public Property Get TilesQuestioned() As Long
    TilesQuestioned = m_TilesQuestioned
End Property

'*******************************************************************************
' TilesExposed (PROPERTY GET)
'*******************************************************************************
Public Property Get TilesExposed() As Long
    TilesExposed = m_TilesExposed
End Property

'*******************************************************************************
' imgCell_Click (SUB)
'
' PARAMETERS:
' (In/Out) - Index - Integer -
'
' DESCRIPTION:
' ***Description goes here***
'*******************************************************************************
Private Sub imgCell_Click(Index As Integer)
    RaiseEvent Click
End Sub

'*******************************************************************************
' imgCell_MouseDown (SUB)
'
' PARAMETERS:
' (In/Out) - Index  - Integer -
' (In/Out) - Button - Integer -
' (In/Out) - Shift  - Integer -
' (In/Out) - X      - Single  -
' (In/Out) - Y      - Single  -
'
' DESCRIPTION:
' ***Description goes here***
'*******************************************************************************
Private Sub imgCell_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim oCell   As cCell
Dim lIdx    As Integer

    m_lShift = 0
    
    'For simultanaeous click of both the L and R mouse buttons
    If Shift = 0 _
    And Button = vbLeftButton Then
        m_fLeftButton = True
    End If
    
    If (Shift = vbShiftMask _
    And Button = vbLeftButton) _
    Or (m_fLeftButton _
    And Button = vbRightButton) Then
        With m_oGrid
            If Not .Item(Index).Exposed Then
                m_lShift = 1
            Else
                m_lShift = 2
            End If
            
            'N
            lIdx = .Item(Index).N
            If lIdx <> 0 Then
                Set oCell = .Item(lIdx)
                If Not oCell.Exposed _
                And Not oCell.Questioned _
                And Not oCell.Flagged Then
                    imgCell(lIdx).Picture = LoadResPicture(EmptyTile, vbResBitmap)
                ElseIf oCell.Questioned Then
                    imgCell(lIdx).Picture = LoadResPicture(QuestionDown, vbResBitmap)
                End If
                Set oCell = Nothing
            End If
            
            'NE
            lIdx = .Item(Index).NE
            If lIdx <> 0 Then
                Set oCell = .Item(lIdx)
                If Not oCell.Exposed _
                And Not oCell.Questioned _
                And Not oCell.Flagged Then
                    imgCell(lIdx).Picture = LoadResPicture(EmptyTile, vbResBitmap)
                ElseIf oCell.Questioned Then
                    imgCell(lIdx).Picture = LoadResPicture(QuestionDown, vbResBitmap)
                End If
                Set oCell = Nothing
            End If
            
            'E
            lIdx = .Item(Index).E
            If lIdx <> 0 Then
                Set oCell = .Item(lIdx)
                If Not oCell.Exposed _
                And Not oCell.Questioned _
                And Not oCell.Flagged Then
                    imgCell(lIdx).Picture = LoadResPicture(EmptyTile, vbResBitmap)
                ElseIf oCell.Questioned Then
                    imgCell(lIdx).Picture = LoadResPicture(QuestionDown, vbResBitmap)
                End If
                Set oCell = Nothing
            End If
            
            'SE
            lIdx = .Item(Index).SE
            If lIdx <> 0 Then
                Set oCell = .Item(lIdx)
                If Not oCell.Exposed _
                And Not oCell.Questioned _
                And Not oCell.Flagged Then
                    imgCell(lIdx).Picture = LoadResPicture(EmptyTile, vbResBitmap)
                ElseIf oCell.Questioned Then
                    imgCell(lIdx).Picture = LoadResPicture(QuestionDown, vbResBitmap)
                End If
                Set oCell = Nothing
            End If
            
            'S
            lIdx = .Item(Index).S
            If lIdx <> 0 Then
                Set oCell = .Item(lIdx)
                If Not oCell.Exposed _
                And Not oCell.Questioned _
                And Not oCell.Flagged Then
                    imgCell(lIdx).Picture = LoadResPicture(EmptyTile, vbResBitmap)
                ElseIf oCell.Questioned Then
                    imgCell(lIdx).Picture = LoadResPicture(QuestionDown, vbResBitmap)
                End If
                Set oCell = Nothing
            End If
            
            'SW
            lIdx = .Item(Index).SW
            If lIdx <> 0 Then
                Set oCell = .Item(lIdx)
                If Not oCell.Exposed _
                And Not oCell.Questioned _
                And Not oCell.Flagged Then
                    imgCell(lIdx).Picture = LoadResPicture(EmptyTile, vbResBitmap)
                ElseIf oCell.Questioned Then
                    imgCell(lIdx).Picture = LoadResPicture(QuestionDown, vbResBitmap)
                End If
                Set oCell = Nothing
            End If
            
            'W
            lIdx = .Item(Index).W
            If lIdx <> 0 Then
                Set oCell = .Item(lIdx)
                If Not oCell.Exposed _
                And Not oCell.Questioned _
                And Not oCell.Flagged Then
                    imgCell(lIdx).Picture = LoadResPicture(EmptyTile, vbResBitmap)
                ElseIf oCell.Questioned Then
                    imgCell(lIdx).Picture = LoadResPicture(QuestionDown, vbResBitmap)
                End If
                Set oCell = Nothing
            End If
            
            'NW
            lIdx = .Item(Index).NW
            If lIdx <> 0 Then
                Set oCell = .Item(lIdx)
                If Not oCell.Exposed _
                And Not oCell.Questioned _
                And Not oCell.Flagged Then
                    imgCell(lIdx).Picture = LoadResPicture(EmptyTile, vbResBitmap)
                ElseIf oCell.Questioned Then
                    imgCell(lIdx).Picture = LoadResPicture(QuestionDown, vbResBitmap)
                End If
                Set oCell = Nothing
            End If
        
            If Not .Item(Index).Exposed _
            And Not .Item(Index).Flagged Then
                If Not .Item(Index).Questioned Then
                    imgCell(Index).Picture = LoadResPicture(EmptyTile, vbResBitmap)
                Else 'If .Item(Index).Questioned Then
                    imgCell(Index).Picture = LoadResPicture(QuestionDown, vbResBitmap)
                End If
            End If
        End With
    Else
        If Not m_oGrid.Item(Index).Exposed _
        And Not m_oGrid.Item(Index).Flagged _
        And Button = vbLeftButton Then
            imgCell(Index).Picture = LoadResPicture(EmptyTile, vbResBitmap)
        End If
    End If
    
    RaiseEvent MouseDown(CLng(Index), Button, Shift, X, Y)

End Sub

'*******************************************************************************
' CheckButtons (SUB)
'
' PARAMETERS:
' (In/Out) - Index - Integer -
'
' DESCRIPTION:
' ***Description goes here***
'*******************************************************************************
Private Sub CheckButtons(Index As Integer)
    
    CheckButtons2 m_oGrid.Item(Index).N
    CheckButtons2 m_oGrid.Item(Index).NE
    CheckButtons2 m_oGrid.Item(Index).E
    CheckButtons2 m_oGrid.Item(Index).SE
    CheckButtons2 m_oGrid.Item(Index).S
    CheckButtons2 m_oGrid.Item(Index).SW
    CheckButtons2 m_oGrid.Item(Index).W
    CheckButtons2 m_oGrid.Item(Index).NW
    
End Sub

'*******************************************************************************
' CheckButtons2 (SUB)
'
' PARAMETERS:
' (In/Out) - iID - Integer -
'
' DESCRIPTION:
' Recursive routine to clear grid cells after an empty cell was clicked
'*******************************************************************************
Private Sub CheckButtons2(iID As Integer)
    If iID > 0 Then
        If Not m_oGrid.Item(iID).Bomb _
        And Not m_oGrid.Item(iID).Exposed Then
            imgCell_MouseUp iID, vbLeftButton, 0, 0, 0
        End If
    End If
End Sub

'*******************************************************************************
' imgCell_MouseUp (SUB)
'
' PARAMETERS:
' (In/Out) - Index  - Integer -
' (In/Out) - Button - Integer -
' (In/Out) - Shift  - Integer -
' (In/Out) - X      - Single  -
' (In/Out) - Y      - Single  -
'
' DESCRIPTION:
' ***Description goes here***
'*******************************************************************************
Private Sub imgCell_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lBC     As Long

Dim oCell   As cCell
Dim lIdx    As Integer

    If m_lShift <> 0 Then
        With m_oGrid
            'N
            lIdx = .Item(Index).N
            If lIdx <> 0 Then
                Set oCell = .Item(lIdx)
                With oCell
                    If Not .Exposed _
                    And Not .Questioned _
                    And Not .Flagged Then
                        imgCell(lIdx).Picture = LoadResPicture(Tile, vbResBitmap)
                    ElseIf .Questioned Then
                        imgCell(lIdx).Picture = LoadResPicture(Question, vbResBitmap)
                    End If
                End With
                Set oCell = Nothing
            End If
            
            'NE
            lIdx = .Item(Index).NE
            If lIdx <> 0 Then
                Set oCell = .Item(lIdx)
                With oCell
                    If Not .Exposed _
                    And Not .Questioned _
                    And Not .Flagged Then
                        imgCell(lIdx).Picture = LoadResPicture(Tile, vbResBitmap)
                    ElseIf .Questioned Then
                        imgCell(lIdx).Picture = LoadResPicture(Question, vbResBitmap)
                    End If
                End With
                Set oCell = Nothing
            End If
            
            'E
            lIdx = .Item(Index).E
            If lIdx <> 0 Then
                Set oCell = .Item(lIdx)
                With oCell
                    If Not .Exposed _
                    And Not .Questioned _
                    And Not .Flagged Then
                        imgCell(lIdx).Picture = LoadResPicture(Tile, vbResBitmap)
                    ElseIf .Questioned Then
                        imgCell(lIdx).Picture = LoadResPicture(Question, vbResBitmap)
                    End If
                End With
                Set oCell = Nothing
            End If
            
            'SE
            lIdx = .Item(Index).SE
            If lIdx <> 0 Then
                Set oCell = .Item(lIdx)
                With oCell
                    If Not .Exposed _
                    And Not .Questioned _
                    And Not .Flagged Then
                        imgCell(lIdx).Picture = LoadResPicture(Tile, vbResBitmap)
                    ElseIf .Questioned Then
                        imgCell(lIdx).Picture = LoadResPicture(Question, vbResBitmap)
                    End If
                End With
                Set oCell = Nothing
            End If
            
            'S
            lIdx = .Item(Index).S
            If lIdx <> 0 Then
                Set oCell = .Item(lIdx)
                With oCell
                    If Not .Exposed _
                    And Not .Questioned _
                    And Not .Flagged Then
                        imgCell(lIdx).Picture = LoadResPicture(Tile, vbResBitmap)
                    ElseIf .Questioned Then
                        imgCell(lIdx).Picture = LoadResPicture(Question, vbResBitmap)
                    End If
                End With
                Set oCell = Nothing
            End If
            
            'SW
            lIdx = .Item(Index).SW
            If lIdx <> 0 Then
                Set oCell = .Item(lIdx)
                With oCell
                    If Not .Exposed _
                    And Not .Questioned _
                    And Not .Flagged Then
                        imgCell(lIdx).Picture = LoadResPicture(Tile, vbResBitmap)
                    ElseIf .Questioned Then
                        imgCell(lIdx).Picture = LoadResPicture(Question, vbResBitmap)
                    End If
                End With
                Set oCell = Nothing
            End If
            
            'W
            lIdx = .Item(Index).W
            If lIdx <> 0 Then
                Set oCell = .Item(lIdx)
                With oCell
                    If Not .Exposed _
                    And Not .Questioned _
                    And Not .Flagged Then
                        imgCell(lIdx).Picture = LoadResPicture(Tile, vbResBitmap)
                    ElseIf .Questioned Then
                        imgCell(lIdx).Picture = LoadResPicture(Question, vbResBitmap)
                    End If
                End With
                Set oCell = Nothing
            End If
            
            'NW
            lIdx = .Item(Index).NW
            If lIdx <> 0 Then
                Set oCell = .Item(lIdx)
                With oCell
                    If Not .Exposed _
                    And Not .Questioned _
                    And Not .Flagged Then
                        imgCell(lIdx).Picture = LoadResPicture(Tile, vbResBitmap)
                    ElseIf .Questioned Then
                        imgCell(lIdx).Picture = LoadResPicture(Question, vbResBitmap)
                    End If
                End With
                Set oCell = Nothing
            End If
            
            If m_lShift = 1 Then
                imgCell(Index).Picture = LoadResPicture(Tile, vbResBitmap)
            ElseIf m_lShift = 2 Then
            
            End If
        End With
    End If
    
Dim lFlagged As Long
    
    If Not (m_oGrid.Item(Index).Exposed) Or (m_lShift <> 1) Then
        If m_lShift = 2 Then
            'Check if there is any flagged tiles
            lFlagged = 0
            With m_oGrid
                'N
                lIdx = .Item(Index).N
                If lIdx <> 0 Then
                    If .Item(lIdx).Flagged Then
                        lFlagged = lFlagged + 1
                    End If
                End If
                
                'NE
                lIdx = .Item(Index).NE
                If lIdx <> 0 Then
                    If .Item(lIdx).Flagged Then
                        lFlagged = lFlagged + 1
                    End If
                End If
                
                'E
                lIdx = .Item(Index).E
                If lIdx <> 0 Then
                    If .Item(lIdx).Flagged Then
                        lFlagged = lFlagged + 1
                    End If
                End If
                
                'SE
                lIdx = .Item(Index).SE
                If lIdx <> 0 Then
                    If .Item(lIdx).Flagged Then
                        lFlagged = lFlagged + 1
                    End If
                End If
                
                'S
                lIdx = .Item(Index).S
                If lIdx <> 0 Then
                    If .Item(lIdx).Flagged Then
                        lFlagged = lFlagged + 1
                    End If
                End If
                
                'SW
                lIdx = .Item(Index).SW
                If lIdx <> 0 Then
                    If .Item(lIdx).Flagged Then
                        lFlagged = lFlagged + 1
                    End If
                End If
                
                'W
                lIdx = .Item(Index).W
                If lIdx <> 0 Then
                    If .Item(lIdx).Flagged Then
                        lFlagged = lFlagged + 1
                    End If
                End If
                
                'NW
                lIdx = .Item(Index).NW
                If lIdx <> 0 Then
                    If .Item(lIdx).Flagged Then
                        lFlagged = lFlagged + 1
                    End If
                End If
                
                'If the BombCount = Nr of tiles flagged
                'Then click on all the other tiles
                If .BombCount(Index) = lFlagged Then
                    m_lShift = 0
                    With .Item(Index)
                        If .N <> 0 Then imgCell_MouseUp .N, vbLeftButton, 0, 0, 0
                        If .NE <> 0 Then imgCell_MouseUp .NE, vbLeftButton, 0, 0, 0
                        If .E <> 0 Then imgCell_MouseUp .E, vbLeftButton, 0, 0, 0
                        If .SE <> 0 Then imgCell_MouseUp .SE, vbLeftButton, 0, 0, 0
                        If .S <> 0 Then imgCell_MouseUp .S, vbLeftButton, 0, 0, 0
                        If .SW <> 0 Then imgCell_MouseUp .SW, vbLeftButton, 0, 0, 0
                        If .W <> 0 Then imgCell_MouseUp .W, vbLeftButton, 0, 0, 0
                        If .NW <> 0 Then imgCell_MouseUp .NW, vbLeftButton, 0, 0, 0
                    End With
                End If
            End With
        ElseIf Not (m_fLeftButton _
        And Button = vbRightButton) Then
            Select Case Button
            Case vbLeftButton
                If Not m_oGrid.Item(Index).Flagged Then
                    m_oGrid.Item(Index).Exposed = True
                    If m_oGrid.Item(Index).Bomb Then
                        imgCell(Index).Picture = LoadResPicture(ExplodedBomb, vbResBitmap)
                        FinishGrid
                        RaiseEvent GridEvent(CLng(Index), geBombed)
                    Else
                        Select Case m_oGrid.BombCount(Index)
                        Case 0
                            lBC = BombCounter0
                        Case 1
                            lBC = BombCounter1
                        Case 2
                            lBC = BombCounter2
                        Case 3
                            lBC = BombCounter3
                        Case 4
                            lBC = BombCounter4
                        Case 5
                            lBC = BombCounter5
                        Case 6
                            lBC = BombCounter6
                        Case 7
                            lBC = BombCounter7
                        Case 8
                            lBC = BombCounter8
                        End Select
                        imgCell(Index).Picture = LoadResPicture(lBC, vbResBitmap)
                        If lBC = BombCounter0 Then
                            CheckButtons Index
                        End If
                        CheckIfWon
                    End If
                End If
            Case vbRightButton
                If Not m_oGrid.Item(Index).Exposed Then
                    If m_oGrid.Item(Index).Questioned Then
                        m_TilesQuestioned = m_TilesQuestioned - 1
                        m_oGrid.Item(Index).Questioned = False
                        imgCell(Index).Picture = LoadResPicture(Tile, vbResBitmap)
                        RaiseEvent GridEvent(CLng(Index), geUnQuestion)
                    ElseIf m_oGrid.Item(Index).Flagged Then
                        m_TilesFlagged = m_TilesFlagged - 1
                        m_oGrid.Item(Index).Flagged = False
                        m_TilesQuestioned = m_TilesQuestioned + 1
                        m_oGrid.Item(Index).Questioned = True
                        imgCell(Index).Picture = LoadResPicture(Question, vbResBitmap)
                        RaiseEvent GridEvent(CLng(Index), geQuestion)
                    Else
                        m_TilesFlagged = m_TilesFlagged + 1
                        m_oGrid.Item(Index).Flagged = True
                        imgCell(Index).Picture = LoadResPicture(Flag, vbResBitmap)
                        RaiseEvent GridEvent(CLng(Index), geFlagged)
                    End If
                End If
            End Select
        End If
    End If
    
    RaiseEvent MouseUp(CLng(Index), Button, Shift, X, Y)

    m_lShift = 0
    
    'For simultanaeous click of both the L and R mouse buttons
    If Shift = 0 _
    And Button = vbLeftButton Then
        m_fLeftButton = False
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
    m_TwipsPerPixelY = Screen.TwipsPerPixelY
    m_TwipsPerPixelX = Screen.TwipsPerPixelX
    Set m_oGrid = New cGrid
    imgCell(0).Move -17 * m_TwipsPerPixelX, -17 * m_TwipsPerPixelY, 16 * m_TwipsPerPixelX, 16 * m_TwipsPerPixelY
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
    UserControl.Line (0, m_TwipsPerPixelY + m_TwipsPerPixelY)-(UserControl.Width - m_TwipsPerPixelX - m_TwipsPerPixelX, m_TwipsPerPixelY + m_TwipsPerPixelY), QBColor(8)
    
    UserControl.Line (0, 0)-(0, UserControl.Height), QBColor(8)
    UserControl.Line (m_TwipsPerPixelX, 0)-(m_TwipsPerPixelX, UserControl.Height - m_TwipsPerPixelY), QBColor(8)
    UserControl.Line (m_TwipsPerPixelX + m_TwipsPerPixelX, 0)-(m_TwipsPerPixelX + m_TwipsPerPixelX, UserControl.Height - m_TwipsPerPixelY - m_TwipsPerPixelY), QBColor(8)
    
    UserControl.Line (m_TwipsPerPixelX, UserControl.Height - m_TwipsPerPixelY)-(UserControl.Width, UserControl.Height - m_TwipsPerPixelY), RGB(255, 255, 255)
    UserControl.Line (m_TwipsPerPixelX + m_TwipsPerPixelX, UserControl.Height - (m_TwipsPerPixelY + m_TwipsPerPixelY))-(UserControl.Width, UserControl.Height - (m_TwipsPerPixelY + m_TwipsPerPixelY)), RGB(255, 255, 255)
    UserControl.Line (m_TwipsPerPixelX + m_TwipsPerPixelX + m_TwipsPerPixelX, UserControl.Height - m_TwipsPerPixelY - m_TwipsPerPixelY - m_TwipsPerPixelY)-(UserControl.Width - m_TwipsPerPixelX, UserControl.Height - (m_TwipsPerPixelY + m_TwipsPerPixelY + m_TwipsPerPixelY)), RGB(255, 255, 255)
    
    UserControl.Line (UserControl.Width - m_TwipsPerPixelX, m_TwipsPerPixelY)-(UserControl.Width - m_TwipsPerPixelX, UserControl.Height), RGB(255, 255, 255)
    UserControl.Line (UserControl.Width - m_TwipsPerPixelX - m_TwipsPerPixelX, m_TwipsPerPixelY + m_TwipsPerPixelY)-(UserControl.Width - m_TwipsPerPixelX - m_TwipsPerPixelX, UserControl.Height), RGB(255, 255, 255)
    UserControl.Line (UserControl.Width - m_TwipsPerPixelX - m_TwipsPerPixelX - m_TwipsPerPixelX, m_TwipsPerPixelY + m_TwipsPerPixelY + m_TwipsPerPixelY)-(UserControl.Width - m_TwipsPerPixelX - m_TwipsPerPixelX - m_TwipsPerPixelX, UserControl.Height), RGB(255, 255, 255)
    
End Sub

'*******************************************************************************
' Rows (PROPERTY GET)
'*******************************************************************************
Public Property Get Rows() As Long
    Rows = m_Rows
End Property

'*******************************************************************************
' Rows (PROPERTY LET)
'*******************************************************************************
Public Property Let Rows(ByVal New_Rows As Long)
Attribute Rows.VB_Description = "Number of Rows"
    If New_Rows < 8 Then
        New_Rows = 8
    End If
    m_Rows = New_Rows
    m_Height = (m_Rows * imgCell(0).Height) + (6 * m_TwipsPerPixelY)
    UserControl.Height = m_Height
    PropertyChanged "Rows"
End Property

'*******************************************************************************
' Cols (PROPERTY GET)
'*******************************************************************************
Public Property Get Cols() As Long
Attribute Cols.VB_Description = "Number of Columns"
    Cols = m_Cols
End Property

'*******************************************************************************
' Cols (PROPERTY LET)
'*******************************************************************************
Public Property Let Cols(ByVal New_Cols As Long)
    If New_Cols < 8 Then
        New_Cols = 8
    End If
    m_Cols = New_Cols
    m_Width = (m_Cols * imgCell(0).Width) + (6 * m_TwipsPerPixelX)
    UserControl.Width = m_Width
    PropertyChanged "Cols"
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
    m_TwipsPerPixelY = Screen.TwipsPerPixelY
    m_TwipsPerPixelX = Screen.TwipsPerPixelX
    Rows = m_def_Rows
    Cols = m_def_Cols
    Bombs = m_def_Bombs
    DrawGrid
End Sub

'*******************************************************************************
' UserControl_ReadProperties (SUB)
'
' PARAMETERS:
' (In/Out) - PropBag - PropertyBag -
'
' DESCRIPTION:
' Load property values from storage
'*******************************************************************************
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Rows = PropBag.ReadProperty("Rows", m_def_Rows)
    Cols = PropBag.ReadProperty("Cols", m_def_Cols)
    Bombs = PropBag.ReadProperty("Bombs", m_def_Bombs)
    DrawGrid
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
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
    If UserControl.Width <> m_Width Then
        UserControl.Width = m_Width
    End If
    If UserControl.Height <> m_Height Then
        UserControl.Height = m_Height
    End If
End Sub

'*******************************************************************************
' UserControl_Terminate (SUB)
'
' PARAMETERS:
' None
'
' DESCRIPTION:
' ***Description goes here***
'*******************************************************************************
Private Sub UserControl_Terminate()
    Set m_oGrid = Nothing
End Sub

'*******************************************************************************
' UserControl_WriteProperties (SUB)
'
' PARAMETERS:
' (In/Out) - PropBag - PropertyBag -
'
' DESCRIPTION:
' Write property values to storage
'*******************************************************************************
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Rows", m_Rows, m_def_Rows)
    Call PropBag.WriteProperty("Cols", m_Cols, m_def_Cols)
    Call PropBag.WriteProperty("Bombs", m_Bombs, m_def_Bombs)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

'*******************************************************************************
' Bombs (PROPERTY GET)
'*******************************************************************************
Public Property Get Bombs() As Long
    Bombs = m_Bombs
End Property

'*******************************************************************************
' Bombs (PROPERTY LET)
'*******************************************************************************
Public Property Let Bombs(ByVal New_Bombs As Long)
Attribute Bombs.VB_Description = "The number of Mines"
    m_Bombs = New_Bombs
    PropertyChanged "Bombs"
End Property

'*******************************************************************************
' DrawGrid (SUB)
'
' PARAMETERS:
' None
'
' DESCRIPTION:
' Draw the grid
'*******************************************************************************
Public Sub DrawGrid()
Dim cell    As cCell
Dim iID     As Integer
Dim iCol    As Integer
Dim iRow    As Integer
Dim iWidth  As Integer
Dim iHeight As Integer
Dim ID      As Long

    Screen.MousePointer = vbHourglass
    
    m_TilesFlagged = 0
    m_TilesQuestioned = 0
    m_TilesExposed = 0
    
    m_oGrid.Initialize m_Rows, m_Cols, m_Bombs
    
    iWidth = imgCell(0).Width
    iHeight = imgCell(0).Height
    
    For ID = 1 To imgCell.UBound
        Unload imgCell(ID)
    Next 'ID
        
    For Each cell In m_oGrid
        iID = cell.ID
        iCol = cell.Col
        iRow = cell.Row
        
        Load imgCell(iID)
        With imgCell(iID)
            .Move (iCol - 1) * (iWidth) + (3 * m_TwipsPerPixelX), (iRow - 1) * (iHeight) + (3 * m_TwipsPerPixelY), iWidth, iHeight
            .Picture = LoadResPicture(Tile, vbResBitmap)
            .Visible = True
        End With 'imgCell(iID)
    Next 'cell
    
    Enabled = True

    RaiseEvent Resize
    RaiseEvent GridEvent(0, geNew)
    
    Screen.MousePointer = vbDefault
    
End Sub

'*******************************************************************************
' FinishGrid (SUB)
'
' PARAMETERS:
' (In/Out) - ShowAll - Boolean -
'
' DESCRIPTION:
' After a bomb was clicked show where all the bombs are
'*******************************************************************************
Private Sub FinishGrid(Optional ShowAll As Boolean = False)
Dim oCell As cCell

    Screen.MousePointer = vbHourglass
    
    For Each oCell In m_oGrid
        With oCell
            If Not .Exposed Then
                If .Bomb Then
                    If Not .Flagged Then
                        imgCell(.ID).Picture = LoadResPicture(Bomb, vbResBitmap)
                    End If
                Else
                    If .Flagged Then
                        imgCell(.ID).Picture = LoadResPicture(FlagWrong, vbResBitmap)
                    End If
                    .Flagged = False
                    If ShowAll Then
                        imgCell_MouseDown .ID, vbLeftButton, 0, 0, 0
                    End If
                End If
            End If
        End With 'oCell
    Next 'oCell
    Enabled = False
    
    Screen.MousePointer = vbDefault
End Sub

'*******************************************************************************
' CheckIfWon (SUB)
'
' PARAMETERS:
' None
'
' DESCRIPTION:
' ***Description goes here***
'*******************************************************************************
Private Sub CheckIfWon()
Dim oCell       As cCell
Dim fWon        As Boolean

    fWon = True
    For Each oCell In m_oGrid
        With oCell
            If (Not .Exposed _
            And Not .Bomb) _
            Or (.Flagged _
            And Not .Bomb) Then
                fWon = False
                Exit For
            End If
        End With 'oCell
    Next 'oCell
    If fWon Then
        Enabled = False
        RaiseEvent GridEvent(0, geDone)
    End If
End Sub

'*******************************************************************************
' Enabled (PROPERTY GET)
'*******************************************************************************
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

'*******************************************************************************
' Enabled (PROPERTY LET)
'*******************************************************************************
Public Property Let Enabled(ByVal New_Enabled As Boolean)
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'*******************************************************************************
' Hint (SUB)
'
' PARAMETERS:
' (In/Out) - DoHint - Boolean - If True will perform the mouseup event
'
' DESCRIPTION:
' ***Description goes here***
'*******************************************************************************
Public Sub Hint(Optional DoHint As Boolean = False)
Dim oCell  As cCell
Dim fFound As Boolean

    Screen.MousePointer = vbHourglass
    
    For Each oCell In m_oGrid
        With oCell
            If Not .Exposed Then
                If Not .Bomb Then
                    If Not .Flagged Then
                        If m_oGrid.BombCount(.ID) = 0 Then
                            If DoHint Then
                                imgCell_MouseUp .ID, vbLeftButton, 0, 0, 0
                            Else
                                imgCell(.ID).Picture = LoadResPicture(EmptyTile, vbResBitmap)
                                DoEvents
                                Sleep 200
                                imgCell(.ID).Picture = LoadResPicture(Tile, vbResBitmap)
                            End If
                            fFound = True
                            Exit For
                        End If
                    End If
                End If
            End If
        End With 'oCell
    Next 'oCell
    
    If Not fFound Then
        fFound = False
        For Each oCell In m_oGrid
            With oCell
                If Not .Exposed Then
                    If Not .Bomb Then
                        If Not .Flagged Then
                            If DoHint Then
                                imgCell_MouseUp .ID, vbLeftButton, 0, 0, 0
                            Else
                                imgCell(.ID).Picture = LoadResPicture(EmptyTile, vbResBitmap)
                                DoEvents
                                Sleep 200
                                imgCell(.ID).Picture = LoadResPicture(Tile, vbResBitmap)
                            End If
                            fFound = True
                            Exit For
                        End If
                    End If
                End If
            End With 'oCell
        Next 'oCell
    End If
    
    Screen.MousePointer = vbDefault
End Sub
