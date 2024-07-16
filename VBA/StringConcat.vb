Function StringConcat(Sep As String, ParamArray Args()) As Variant
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' StringConcat
' By Chip Pearson, chip@cpearson.com, www.cpearson.com
'                  www.cpearson.com/Excel/stringconcatenation.aspx
' This function concatenates all the elements in the Args array,
' delimited by the Sep character, into a single string. This function
' can be used in an array formula. There is a VBA imposed limit that
' a string in a passed in array (e.g.,  calling this function from
' an array formula in a worksheet cell) must be less than 256 characters.
' See the comments at STRING TOO LONG HANDLING for details.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim S As String
Dim N As Long
Dim M As Long
Dim R As Range
Dim NumDims As Long
Dim LB As Long
Dim IsArrayAlloc As Boolean

'''''''''''''''''''''''''''''''''''''''''''
' If no parameters were passed in, return
' vbNullString.
'''''''''''''''''''''''''''''''''''''''''''
If UBound(Args) - LBound(Args) + 1 = 0 Then
    StringConcat = vbNullString
    Exit Function
End If

For N = LBound(Args) To UBound(Args)
    ''''''''''''''''''''''''''''''''''''''''''''''''
    ' Loop through the Args
    ''''''''''''''''''''''''''''''''''''''''''''''''
    If IsObject(Args(N)) = True Then
        '''''''''''''''''''''''''''''''''''''
        ' OBJECT
        ' If we have an object, ensure it
        ' it a Range. The Range object
        ' is the only type of object we'll
        ' work with. Anything else causes
        ' a #VALUE error.
        ''''''''''''''''''''''''''''''''''''
        If TypeOf Args(N) Is Excel.Range Then
            '''''''''''''''''''''''''''''''''''''''''
            ' If it is a Range, loop through the
            ' cells and create append the elements
            ' to the string S.
            '''''''''''''''''''''''''''''''''''''''''
            For Each R In Args(N).Cells
                If Len(R.Text) > 0 Then
                    S = S & R.Text & Sep
                End If
            Next R
        Else
            '''''''''''''''''''''''''''''''''
            ' Unsupported object type. Return
            ' a #VALUE error.
            '''''''''''''''''''''''''''''''''
            StringConcat = CVErr(xlErrValue)
            Exit Function
        End If
    
    ElseIf IsArray(Args(N)) = True Then
        '''''''''''''''''''''''''''''''''''''
        ' ARRAY
        ' If Args(N) is an array, ensure it
        ' is an allocated array.
        '''''''''''''''''''''''''''''''''''''
        IsArrayAlloc = (Not IsError(LBound(Args(N))) And _
            (LBound(Args(N)) <= UBound(Args(N))))
        If IsArrayAlloc = True Then
            ''''''''''''''''''''''''''''''''''''
            ' The array is allocated. Determine
            ' the number of dimensions of the
            ' array.
            '''''''''''''''''''''''''''''''''''''
            NumDims = 1
            On Error Resume Next
            Err.Clear
            NumDims = 1
            Do Until Err.Number <> 0
                LB = LBound(Args(N), NumDims)
                If Err.Number = 0 Then
                    NumDims = NumDims + 1
                Else
                    NumDims = NumDims - 1
                End If
            Loop
            On Error GoTo 0
            Err.Clear
            ''''''''''''''''''''''''''''''''''
            ' The array must have either
            ' one or two dimensions. Greater
            ' that two caues a #VALUE error.
            ''''''''''''''''''''''''''''''''''
            If NumDims > 2 Then
                StringConcat = CVErr(xlErrValue)
                Exit Function
            End If
            If NumDims = 1 Then
                For M = LBound(Args(N)) To UBound(Args(N))
                    If Args(N)(M) <> vbNullString Then
                        S = S & Args(N)(M) & Sep
                    End If
                Next M
                
            Else
                ''''''''''''''''''''''''''''''''''''''''''''''''
                ' STRING TOO LONG HANDLING
                ' Here, the error handler must be set to either
                '   On Error GoTo ContinueLoop
                '   or
                '   On Error GoTo ErrH
                ' If you use ErrH, then any error, including
                ' a string too long error, will cause the function
                ' to return #VALUE and quit. If you use ContinueLoop,
                ' the problematic value is ignored and not included
                ' in the result, and the result is the concatenation
                ' of all non-error values in the input. This code is
                ' used in the case that an input string is longer than
                ' 255 characters.
                ''''''''''''''''''''''''''''''''''''''''''''''''
                On Error GoTo ContinueLoop
                'On Error GoTo ErrH
                Err.Clear
                For M = LBound(Args(N), 1) To UBound(Args(N), 1)
                    If Args(N)(M, 1) <> vbNullString Then
                        S = S & Args(N)(M, 1) & Sep
                    End If
                Next M
                Err.Clear
                M = LBound(Args(N), 2)
                If Err.Number = 0 Then
                    For M = LBound(Args(N), 2) To UBound(Args(N), 2)
                        If Args(N)(M, 2) <> vbNullString Then
                            S = S & Args(N)(M, 2) & Sep
                        End If
                    Next M
                End If
                On Error GoTo ErrH:
            End If
        Else
            If Args(N) <> vbNullString Then
                S = S & Args(N) & Sep
            End If
        End If
        Else
        On Error Resume Next
        If Args(N) <> vbNullString Then
            S = S & Args(N) & Sep
        End If
        On Error GoTo 0
    End If
ContinueLoop:
Next N

'''''''''''''''''''''''''''''
' Remove the trailing Sep
'''''''''''''''''''''''''''''
If Len(Sep) > 0 Then
    If Len(S) > 0 Then
        S = Left(S, Len(S) - Len(Sep))
    End If
End If

StringConcat = S
'''''''''''''''''''''''''''''
' Success. Get out.
'''''''''''''''''''''''''''''
Exit Function
ErrH:
'''''''''''''''''''''''''''''
' Error. Return #VALUE
'''''''''''''''''''''''''''''
StringConcat = CVErr(xlErrValue)
End Function
