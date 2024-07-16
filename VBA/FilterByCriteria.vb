
Function FilterByCriteria(InputValues As Range, Criteria As Range) As Variant

    Dim ResultArray() As Variant
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' This first large block of code determines whether the function
    ' is being called from a worksheet range or by another function.
    ' If it is being called from a worksheet, it must be called from
    ' a range with only one column or only one row. Two-dimensional
    ' ranges will cause a #REF error.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If IsObject(Application.Caller) = True Then
        If Application.Caller.Rows.Count > 1 And Application.Caller.Columns.Count > 1 Then
            DistinctValues = CVErr(xlErrRef)
            Exit Function
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Save the size of the region from which the
    ' function was called and save a flag indicating
    ' whether we need to transpose the result upon
    ' returning.
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    If Application.Caller.Rows.Count > 1 Then
        TransposeAtEnd = True
        ReturnSize = Application.Caller.Rows.Count
    Else
        TransposeAtEnd = False
        ReturnSize = Application.Caller.Columns.Count
    End If
    
    
    
    
    
     
    i = 1
    For Each V In InputValues
        If V = Criteria Then
            ReDim Preserve ResultArray(i)
            ResultArray(i) = V
            i = i + 1
        End If
    Next V
    
    FilterByCriteria = ResultArray
            
End Function


