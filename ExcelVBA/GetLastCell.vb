Public Function GetLastCell(InRange As Range, SearchOrder As XlSearchOrder, _
                        Optional ProhibitEmptyFormula As Boolean = False) As Range
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GetLastCell
' By Chip Pearson, chip@cpearson.com, www.cpearson.com
'
' This returns the last used cell in a worksheet or range. If InRange
' is a single cell, the last cell of the entire worksheet if found. If
' InRange contains two or more cells, the last cell in that range is
' returned.
' If SearchOrder is xlByRows (= 1), the last cell is the last
' (right-most) non-blank cell on the last row of data in the
' worksheet's UsedRange. If SearchOrder is xlByColumns
' (= 2), the last cell is the last (bottom-most) non-blank cell in the
' last (right-most) column of the worksheet's UsedRange. If SearchOrder
' is xlByColumns + xlByRows (= 3), the last cell is the intersection of
' the last row and the last column. Note that this cell may not contain
' any value.
' If SearchOrder is anything other than xlByRows, xlByColumns, or
' xlByRows+xlByColumns, an error 5 is raised.
'
' ProhibitEmptyFormula indicates how to handle the case in which the
' last cell is a formula that evaluates to an empty string. If this setting
' is omitted for False, the last cell is allowed to be a formula that
' evaluates to an empty string. If this setting is True, the last cell
' must be either a static value or a formula that evaluates to a non-empty
' string. The default is False, allowing the last cell to be a formula
' that evaluates to an empty string.
'''''''''''''''''''''''''
' Example:
'       a   b   c
'               d   e
'       f   g
'
' If SearchOrder is xlByRows, the last cell is 'g'. If SearchOrder is
' xlByColumns, the last cell is 'e'. If SearchOrder is xlByRows+xlByColumns,
' the last cell is the intersection of the row containing 'g' and the column
' containing 'e'. This cell has no value in this example.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim WS As Worksheet
Dim R As Range
Dim LastCell As Range
Dim LastR As Range
Dim LastC As Range
Dim SearchRange As Range
Dim LookIn As XlFindLookIn
Dim RR As Range

Set WS = InRange.Worksheet

If ProhibitEmptyFormula = False Then
    LookIn = xlFormulas
Else
    LookIn = xlValues
End If

Select Case SearchOrder
    Case XlSearchOrder.xlByColumns, XlSearchOrder.xlByRows, _
            XlSearchOrder.xlByColumns + XlSearchOrder.xlByRows
        ' OK
    Case Else
        Err.Raise 5
        Exit Function
End Select

With WS
    If InRange.Cells.Count = 1 Then
        Set RR = .UsedRange
    Else
       Set RR = InRange
    End If
    Set R = RR(RR.Cells.Count)
    
    If SearchOrder = xlByColumns Then
        Set LastCell = RR.Find(what:="*", after:=R, LookIn:=LookIn, _
                LookAt:=xlPart, SearchOrder:=xlByColumns, _
                SearchDirection:=xlPrevious, MatchCase:=False)
    ElseIf SearchOrder = xlByRows Then
        Set LastCell = RR.Find(what:="*", after:=R, LookIn:=LookIn, _
                LookAt:=xlPart, SearchOrder:=xlByRows, _
                SearchDirection:=xlPrevious, MatchCase:=False)
    ElseIf SearchOrder = xlByColumns + xlByRows Then
        Set LastC = RR.Find(what:="*", after:=R, LookIn:=LookIn, _
                LookAt:=xlPart, SearchOrder:=xlByColumns, _
                SearchDirection:=xlPrevious, MatchCase:=False)
        Set LastR = RR.Find(what:="*", after:=R, LookIn:=LookIn, _
                LookAt:=xlPart, SearchOrder:=xlByRows, _
                SearchDirection:=xlPrevious, MatchCase:=False)
        Set LastCell = Application.Intersect(LastR.EntireRow, LastC.EntireColumn)
    Else
        Err.Raise 5
        Exit Function
    End If
End With

Set GetLastCell = LastCell

End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' END CODE GetLastCell
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
