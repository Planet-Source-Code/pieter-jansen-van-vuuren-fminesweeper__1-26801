VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'********************************************************************
'Name:              cGrid
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

Private m_Col   As Collection

Private m_Rows  As Integer
Private m_Cols  As Integer

Private m_Bombs As Integer

'*******************************************************************************
' GetID (FUNCTION)
'
' PARAMETERS:
' (In/Out) - Row - Integer -
' (In/Out) - Col - Integer -
'
' RETURN VALUE:
' Integer - Cell ID
'
' DESCRIPTION:
' Calculates Cell ID given the row and column
'*******************************************************************************
Public Function GetID(Row As Integer, Col As Integer) As Integer
    GetID = (Row * m_Cols) - (m_Cols - Col)
End Function

'*******************************************************************************
' Rows (PROPERTY LET)
'*******************************************************************************
Public Property Let Rows(lRows As Integer)
    m_Rows = lRows
End Property
'*******************************************************************************
' Rows (PROPERTY GET)
'*******************************************************************************
Public Property Get Rows() As Integer
    Rows = m_Rows
End Property

'*******************************************************************************
' Cols (PROPERTY LET)
'*******************************************************************************
Public Property Let Cols(lCols As Integer)
    m_Cols = lCols
End Property
'*******************************************************************************
' Cols (PROPERTY GET)
'*******************************************************************************
Public Property Get Cols() As Integer
    Cols = m_Cols
End Property

'*******************************************************************************
' Bombs (PROPERTY LET)
'*******************************************************************************
Public Property Let Bombs(lBombs As Integer)
    m_Bombs = lBombs
End Property
'*******************************************************************************
' Bombs (PROPERTY GET)
'*******************************************************************************
Public Property Get Bombs() As Integer
    Bombs = m_Bombs
End Property

'*******************************************************************************
' BombCount (PROPERTY GET)
'
' DESCRIPTION:
' Calculates the number of Bombs around a cell
'*******************************************************************************
Public Property Get BombCount(ID As Integer) As Integer
Dim cell As cCell

    BombCount = 0
    Set cell = m_Col(ID)
    
    With cell
        If .N <> 0 Then
            If m_Col(.N).Bomb Then
                BombCount = BombCount + 1
            End If
        End If
        If .NE <> 0 Then
            If m_Col(.NE).Bomb Then
                BombCount = BombCount + 1
            End If
        End If
        If .E <> 0 Then
            If m_Col(.E).Bomb Then
                BombCount = BombCount + 1
            End If
        End If
        If .SE <> 0 Then
            If m_Col(.SE).Bomb Then
                BombCount = BombCount + 1
            End If
        End If
        If .S <> 0 Then
            If m_Col(.S).Bomb Then
                BombCount = BombCount + 1
            End If
        End If
        If .SW <> 0 Then
            If m_Col(.SW).Bomb Then
                BombCount = BombCount + 1
            End If
        End If
        If .W <> 0 Then
            If m_Col(.W).Bomb Then
                BombCount = BombCount + 1
            End If
        End If
        If .NW <> 0 Then
            If m_Col(.NW).Bomb Then
                BombCount = BombCount + 1
            End If
        End If
    End With
    
End Property

'*******************************************************************************
' Add (FUNCTION)
'
' PARAMETERS:
' (In/Out) - ID  - Integer -
' (In/Out) - Row - Integer -
' (In/Out) - Col - Integer -
' (In/Out) - N   - Integer -
' (In/Out) - NE  - Integer -
' (In/Out) - E   - Integer -
' (In/Out) - SE  - Integer -
' (In/Out) - S   - Integer -
' (In/Out) - SW  - Integer -
' (In/Out) - W   - Integer -
' (In/Out) - NW  - Integer -
'
' RETURN VALUE:
' cCell -
'
' DESCRIPTION:
' Add a Cell to the Grid Collection
'*******************************************************************************
Public Function Add(ID As Integer, Row As Integer, Col As Integer, N As Integer, NE As Integer, E As Integer, SE As Integer, S As Integer, SW As Integer, W As Integer, NW As Integer) As cCell
    Set Add = New cCell
    With Add
        .ID = ID
        .Row = Row
        .Col = Col
        .N = N
        .NE = NE
        .E = E
        .SE = SE
        .S = S
        .SW = SW
        .W = W
        .NW = NW
    End With 'Add
    m_Col.Add Add, CStr(ID)
End Function

'*******************************************************************************
' Initialize (SUB)
'
' PARAMETERS:
' (In/Out) - Rows - Variant -
' (In/Out) - Rows - Cols    -
' (In/Out) - Rows - NoBomb  -
'
' DESCRIPTION:
' Initializes a Grid
'*******************************************************************************
Public Sub Initialize(Optional Rows, _
                      Optional Cols, _
                      Optional NoBomb)
                      
Dim iR  As Integer
Dim iC  As Integer
Dim iID As Integer

Dim lN  As Integer
Dim lNE As Integer
Dim lE  As Integer
Dim lSE As Integer
Dim lS         As Integer
Dim lSW        As Integer
Dim lW         As Integer
Dim lNW        As Integer
    
Dim iCellCount As Integer
Dim i          As Integer
Dim j          As Integer

    If Not IsMissing(Rows) Then
        m_Rows = CInt(Rows)
    End If
    If Not IsMissing(Cols) Then
        m_Cols = CInt(Cols)
    End If
    If Not IsMissing(NoBomb) Then
        m_Bombs = CInt(NoBomb)
    End If

    Randomize
    
    Set m_Col = New Collection
    
    iCellCount = m_Rows * m_Cols
    ReDim aBombID(1 To m_Bombs)
'    For i = 1 To m_Bombs
'        iID = 0
'        While iID = 0
'            iID = CInt((iCellCount - 2) * Rnd + 1)
'            For j = 1 To m_Bombs
'                If aBombID(j) = iID Then
'                    iID = 0
'                    Exit For
'                End If
'            Next 'j
'            If iID <> 0 Then
'                aBombID(i) = iID
'            End If
'        Wend
'    Next 'i
    Dim tBombs As String
    tBombs = ""
    For i = 1 To m_Bombs
        iID = 0
        While iID = 0
            iID = CInt((iCellCount - 2) * Rnd + 1)
            'For j = 1 To m_Bombs
            '    If aBombID(j) = iID Then
            '        iID = 0
            '        Exit For
            '    End If
            'Next 'j
            'If iID <> 0 Then
            If InStr(tBombs, "," & iID & ",") = 0 Then
                aBombID(i) = iID
                tBombs = tBombs & "," & iID & ","
            Else
                iID = 0
            End If
        Wend
    Next 'i
'    Dim iBomb As Long
'    iBomb = 1
'    For iR = 1 To m_Rows
'        For iC = 1 To m_Cols
'            'Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
'            If Int((1 - 0 + 1) * Rnd + 0) = 1 Then
'                aBombID(iBomb) = GetID(iR, iC)
'                iBomb = iBomb + 1
'                If iBomb > m_Bombs Then
'                    Exit For
'                End If
'            End If
'        Next 'iC
'        If iBomb > m_Bombs Then
'            Exit For
'        End If
'    Next 'iR
    
    For iR = 1 To m_Rows
        For iC = 1 To m_Cols
            iID = GetID(iR, iC)
            
            lN = iID - m_Cols
            lNE = iID - m_Cols + 1
            lE = iID + 1
            lSE = iID + m_Cols + 1
            lS = iID + m_Cols
            lSW = iID + m_Cols - 1
            lW = iID - 1
            lNW = iID - m_Cols - 1
            
            If iR = 1 And iC = 1 Then
                lN = 0
                lNE = 0
                lSW = 0
                lW = 0
                lNW = 0
            ElseIf iR = 1 And iC <> 1 And iC <> m_Cols Then
                lN = 0
                lNE = 0
                lNW = 0
            ElseIf iR = 1 And iC = m_Cols Then
                lN = 0
                lNE = 0
                lE = 0
                lSE = 0
                lNW = 0
            ElseIf iR <> 1 And iR <> m_Rows And iC = 1 Then
                lSW = 0
                lW = 0
                lNW = 0
            ElseIf iR <> 1 And iR <> m_Rows And iC <> 1 And iC <> m_Cols Then
            ElseIf iR <> 1 And iR <> m_Rows And iC = m_Cols Then
                lNE = 0
                lE = 0
                lSE = 0
            ElseIf iR = m_Rows And iC = 1 Then
                lSE = 0
                lS = 0
                lSW = 0
                lW = 0
                lNW = 0
            ElseIf iR = m_Rows And iC <> 1 And iC <> m_Cols Then
                lSE = 0
                lS = 0
                lSW = 0
            ElseIf iR = m_Rows And iC = m_Cols Then
                lNE = 0
                lE = 0
                lSE = 0
                lSE = 0
                lS = 0
                lSW = 0
            End If
            Add iID, iR, iC, lN, lNE, lE, lSE, lS, lSW, lW, lNW
        Next 'iC
    Next 'iR

    For i = 1 To NoBomb
        m_Col(CStr(aBombID(i))).Bomb = True
    Next 'i
    
End Sub

'*******************************************************************************
' Item (FUNCTION)
'
' PARAMETERS:
' (In/Out) - ID - Integer -
'
' RETURN VALUE:
' cCell -
'
' DESCRIPTION:
' ***Description goes here***
'*******************************************************************************
Public Function Item(ID As Integer) As cCell
    Set Item = m_Col(CStr(ID))
End Function

'*******************************************************************************
' Count (FUNCTION)
'
' PARAMETERS:
' None
'
' RETURN VALUE:
' Long -
'
' DESCRIPTION:
' The number of cells in the grid
'*******************************************************************************
Public Function Count() As Long
    Count = m_Col.Count
End Function

'*******************************************************************************
' NewEnum (FUNCTION)
'
' PARAMETERS:
' None
'
' RETURN VALUE:
' IUnknown -
'
' DESCRIPTION:
' NewEnum must return the IUnknown interface of a
' collection's enumerator.
'*******************************************************************************
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
   Set NewEnum = m_Col.[_NewEnum]
End Function

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
    m_Rows = 8
    m_Cols = 8
End Sub

'*******************************************************************************
' Class_Terminate (SUB)
'
' PARAMETERS:
' None
'
' DESCRIPTION:
' ***Description goes here***
'*******************************************************************************
Private Sub Class_Terminate()
    Set m_Col = Nothing
End Sub
