Option Explicit
Option Compare Text

''''''''''''''''''''''''''''''''''''''
' This enum lists the possible types
' of data to count with the CountType
' function.
''''''''''''''''''''''''''''''''''''''
Public Enum cstCountType
    cstTypeNonBlank = 1
    cstTypeNumbers = 2
    cstTypeText = 4
    cstTypeFormulas = 8
    cstTypeNonFormulas = 16
    cstTypeErrors = 32
    cstTypeBlanks = 64
    cstTypeBoolean = 128
    cstTypeDateTime = 256
    cstTypeAll = 4096
End Enum

Public Function CountType(InputRange As Range, CountTypeOf As cstCountType) As Variant
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' CountType
' By Chip Pearson, 22-August-2007.
' Documentation at  http://www.cpearson.com/Excel/CountType.aspx
' Granted to the Public Domain.
' This function returns the number of cells in InputRange that are of the type
' specified by CountTypeOf.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim R As Range
Dim D As Date
Dim CellCount As Long
Application.Volatile True
On Error GoTo ErrH

''''''''''''''''''''''''''''''''
' If the InputRange is Nothing
' get out immediately.
''''''''''''''''''''''''''''''''
If InputRange Is Nothing Then
    CountType = 0
    Exit Function
End If

'''''''''''''''''''''''''''''''''''''
' cstTypeAll returns the total
' cell count, regardless of content.
'''''''''''''''''''''''''''''''''''''
If CountTypeOf = cstTypeAll Then
    CountType = InputRange.Cells.Count
    Exit Function
End If

'''''''''''''''''''''''''''''''''''''
' Loop through the cells in the
' InputRange, testing each condition
' and keeping a running total of each
' data type.
''''''''''''''''''''''''''''''''''''
For Each R In InputRange.Cells
    
    '''''''''''''''''''''''''''''''''''''
    ' cstTypeBlanks
    '''''''''''''''''''''''''''''''''''''
    If (CountTypeOf And cstTypeBlanks) Then
        If R.HasFormula = False Then
            If R.Text = vbNullString Then
                CellCount = CellCount + 1
                GoTo EndOfLoop
            End If
        Else
            If R.Text = vbNullString Then
                CellCount = CellCount + 1
                GoTo EndOfLoop
            End If
        End If
    End If
    
    ''''''''''''''''''''''''''''''''''''''
    ' cstTypeBoolean
    ''''''''''''''''''''''''''''''''''''''
    If (CountTypeOf And cstTypeBoolean) Then
        If (StrComp(CStr(R.Text), "TRUE") = 0) Or (StrComp(CStr(R.Text), "FALSE") = 0) Then
            CellCount = CellCount + 1
            GoTo EndOfLoop
        End If
    End If
    
    ''''''''''''''''''''''''''''''''''''''
    ' cstTypeNonBlank
    ''''''''''''''''''''''''''''''''''''''
    If (CountTypeOf And cstTypeNonBlank) Then
        If R.Text <> vbNullString Then
            CellCount = CellCount + 1
            GoTo EndOfLoop
        End If
    End If
    
    ''''''''''''''''''''''''''''''''''''''
    ' cstTypeNumbers
    ''''''''''''''''''''''''''''''''''''''
    If (CountTypeOf And cstTypeNumbers) Then
        If R.Text <> vbNullString Then
            If IsNumeric(R.Value) Then
                On Error Resume Next
                Err.Clear
                D = DateValue(R.Text)
                If Err.Number <> 0 Then
                    If ((StrComp(CStr(R.Value), "TRUE", vbTextCompare) <> 0) And _
                        ((StrComp(CStr(R.Value), "FALSE", vbTextCompare) <> 0))) Then
                        CellCount = CellCount + 1
                        GoTo EndOfLoop
                    End If
                End If
            End If
        End If
    End If
    
    ''''''''''''''''''''''''''''''''''''''
    ' cstTypeText
    ''''''''''''''''''''''''''''''''''''''
    If (CountTypeOf And cstTypeText) Then
        On Error Resume Next
        If R.Text <> vbNullString Then
            If IsNumeric(R.Value) = False Then
                If IsError(R.Value) = False Then
                    Err.Clear
                    D = DateValue(R.Text)
                    If Err.Number <> 0 Then
                        CellCount = CellCount + 1
                        GoTo EndOfLoop
                    End If
                End If
            End If
        End If
        On Error GoTo 0
    End If
    
    ''''''''''''''''''''''''''''''''''''''
    ' cstTypeFormulas
    ''''''''''''''''''''''''''''''''''''''
    If (CountTypeOf And cstTypeFormulas) Then
        If R.HasFormula = True Then
            CellCount = CellCount + 1
            GoTo EndOfLoop
        End If
    End If
    
    ''''''''''''''''''''''''''''''''''''''
    ' cstTypeNonFormulas
    ''''''''''''''''''''''''''''''''''''''
    If (CountTypeOf And cstTypeNonFormulas) Then
        If R.HasFormula = False Then
            CellCount = CellCount + 1
            GoTo EndOfLoop
        End If
    End If
    
    ''''''''''''''''''''''''''''''''''''''
    ' cstTypeErrors
    ''''''''''''''''''''''''''''''''''''''
    If (CountTypeOf And cstTypeErrors) Then
        If IsError(R.Value) Then
            CellCount = CellCount + 1
            GoTo EndOfLoop
        End If
    End If

    ''''''''''''''''''''''''''''''''''''''
    ' cstTypeDateTime
    ''''''''''''''''''''''''''''''''''''''
    If (CountTypeOf And cstTypeDateTime) Then
        On Error Resume Next
        Err.Clear
        D = DateValue(CStr(R.Text))
        If Err.Number = 0 Then
            CellCount = CellCount + 1
            GoTo EndOfLoop
        Else
            Err.Clear
            D = TimeValue(CStr(R.Text))
            If Err.Number = 0 Then
                CellCount = CellCount + 1
                GoTo EndOfLoop
            End If
        End If
        On Error GoTo ErrH
    End If

''''''''''''''''''''''''''''''''''''''
' end of cell loop
''''''''''''''''''''''''''''''''''''''
EndOfLoop:
Next R
        
''''''''''''''''''''''''''''''''''''''
' Return the result CellCount.
''''''''''''''''''''''''''''''''''''''
CountType = CellCount

''''''''''''''''''''''''''''''''''''''
' Exit Function
''''''''''''''''''''''''''''''''''''''
Exit Function

'''''''''''''''''''''''''''''''''''''''''''''''
' ErrH Error Handler: Should never get here.
'''''''''''''''''''''''''''''''''''''''''''''''
ErrH:
CountType = "ERROR: Cell: " & R.Address(False, False)

End Function
