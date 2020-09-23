Attribute VB_Name = "MStringSortRoutines"
'*******************************************************************************
' MODULE:       MStringSortRoutines
' FILENAME:     C:\My Code\vb\Sorting\MStringSortRoutines.bas
' AUTHOR:       Phil Fresle
' CREATED:      01-Dec-1999
' COPYRIGHT:    Copyright 1999 Frez Systems Limited. All Rights Reserved.
'*******************************************************************************
Option Explicit

'*******************************************************************************
' InsertSortStringsStart (SUB)
'
' PARAMETERS:
' (In/Out) - ListArray()    - String  - Array to sort
' (In)     - bAscending     - Boolean - True to sort ascending, false descending
' (In)     - bCaseSensitive - Boolean - True for a case sensitive sort, false
'                                       for an insensitive one
'
' DESCRIPTION:
' User friendly entry point for InsertSortStrings
'*******************************************************************************
Public Sub InsertSortStringsStart(ListArray, _
                                  Optional ByVal bAscending As Boolean = True, _
                                  Optional ByVal bCaseSensitive As Boolean = False)

Dim lMin         As Long
Dim lMax         As Long
Dim lOrder       As Long
Dim lCompareType As Long

    lMin = LBound(ListArray)
    lMax = UBound(ListArray)
    
    If lMin = lMax Then
        Exit Sub
    End If
    
    ' Order Ascending or Descending?
    lOrder = IIf(bAscending, 1, -1)
    
    ' Case sensitive search or not?
    lCompareType = IIf(bCaseSensitive, vbBinaryCompare, vbTextCompare)
    
    InsertSortStrings ListArray, lMin, lMax, lOrder, lCompareType
End Sub


'*******************************************************************************
' InsertSortStrings (SUB)
'
' PARAMETERS:
' (In/Out) - ListArray()  - String - Array to sort
' (In)     - lMin         - Long   - Start of sorting region within array
' (In)     - lMax         - Long   - End of sorting region within array
' (In)     - lOrder       - Long   - Ascending is -1, Descending is +1, used
'                                    for comparison in StrComp
' (In)     - lCompareType - Long   - Either vbBinaryCompare or vbTextCompare,
'                                    used in StrComp function
'
' DESCRIPTION:
' Simple Insert Sort routine for strings, fast for small arrays as there is no
' recursion (say less than 60 values)
'*******************************************************************************
Private Sub InsertSortStrings(ListArray, _
                              ByVal lMin As Long, _
                              ByVal lMax As Long, _
                              ByVal lOrder As Long, _
                              ByVal lCompareType As Long)
    
Dim sValue  As String
Dim lCount1 As Long
Dim lCount2 As Long
    
    ' Loop through array shifting elements down to their correct place
    For lCount1 = lMin + 1 To lMax
        sValue = CStr(ListArray(lCount1))
        
        ' Find the place to put it
        For lCount2 = lCount1 - 1 To lMin Step -1
            If StrComp(CStr(ListArray(lCount2)), sValue, lCompareType) <> lOrder Then
                Exit For
            End If
            ListArray(lCount2 + 1) = ListArray(lCount2)
        Next lCount2
        
        ' Insert it
        ListArray(lCount2 + 1) = sValue
    Next
End Sub

