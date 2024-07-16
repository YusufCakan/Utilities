Option Explicit
Option Base 1

Private Function CholeskyDecomposition(A() As Variant) As Variant
    'Returns the array L such that LL' = A
    'See https://en.wikipedia.org/?title=Cholesky_decomposition
    Dim L() As Variant
    ReDim L(UBound(A, 1), UBound(A, 2))
    Dim i As Integer
    Dim j As Integer
    Dim temp As Double
    Dim k As Integer
    For j = 1 To UBound(L, 1)
        'handle diagonal case first
        temp = 0
        For k = 1 To j - 1
            temp = temp + L(j, k) ^ 2
        Next k
        If A(j, j) - temp < 0 Then 'not positive definite
            CholeskyDecomposition = CVErr(xlErrValue)
            MsgBox ("The correlation matrix isn't positive definite.")
            Exit Function
        End If
        L(j, j) = (A(j, j) - temp) ^ 0.5
        'handle the non-diagonal elements
        For i = j To UBound(L, 2)
            temp = 0
            For k = 1 To j - 1
                temp = temp + L(i, k) * L(j, k)
            Next k
            L(i, j) = 1 / L(j, j) * (A(i, j) - temp)
        Next i
    Next j
    CholeskyDecomposition = L
End Function

Private Function CheckCorrelationMatrix(CorrelationMatrix As Variant) As Boolean
    CheckCorrelationMatrix = True
    'check that the matrix is:
    '   -square
    '   -symmetrical
    '   -1's along diagonal
    '   -less than 1 for non-diagonal terms
     
    Dim i As Integer
    Dim j As Integer
    'check square first
    If UBound(CorrelationMatrix, 1) <> UBound(CorrelationMatrix, 2) Then
        CheckCorrelationMatrix = False
        Exit Function
    End If
    For i = 1 To UBound(CorrelationMatrix, 1)
        For j = 1 To UBound(CorrelationMatrix, 2)
            If i = j Then
                'check ones on diagonal
                If CorrelationMatrix(i, j) <> 1 Then
                    CheckCorrelationMatrix = False
                    Exit Function
                End If
            Else
                'check less than 1 on non-diagonal
                If Abs(CorrelationMatrix(i, j)) >= 1 Then
                    CheckCorrelationMatrix = False
                    Exit Function
                End If
                'check symmetrical
                If CorrelationMatrix(i, j) <> CorrelationMatrix(j, i) Then
                    CheckCorrelationMatrix = False
                    Exit Function
                End If
            End If
        Next j
    Next i
End Function

Private Function SampleSingleMultiVariateNormal(A) As Variant
    'declare some variables we need before generating the samples
    Dim Z() As Variant
    ReDim Z(UBound(A, 1))
    Dim X() As Variant
    ReDim X(UBound(A, 1))
    Dim i As Integer
    Dim j As Integer
    'we want to calculate X = AZ
    'generate the vector Z first
    For i = 1 To UBound(Z)
        Z(i) = WorksheetFunction.NormInv(Rnd, 0, 1)
    Next i
    'calculate multiply A and Z
    For i = 1 To UBound(Z)
        X(i) = 0
        For j = 1 To UBound(A)
            X(i) = X(i) + A(i, j) * Z(j)
        Next j
    Next i
    'return the results
    SampleSingleMultiVariateNormal = X
End Function

Function SampleMultiVariateNormal(CorrelationMatrixRange As Range) As Variant
    'read in the correlation matrix
    Dim CorrelationMatrix() As Variant
    CorrelationMatrix = CorrelationMatrixRange
    'check that we have a valid correlation matrix
    If CheckCorrelationMatrix(CorrelationMatrix) = False Then
        MsgBox ("The correlation matrix isn't symmetrical.")
        SampleMultiVariateNormal = CVErr(xlErrValue)
        Exit Function
    End If
    'use cholesky matrix to get A such that AA' = CorrelationMatrix
    Dim A() As Variant
    A = CholeskyDecomposition(CorrelationMatrix)
    'prepare the array we want to return
    Dim CallerRows As Integer
    Dim CallerCols As Integer
    Dim Result() As Variant
    With Application.Caller
        CallerRows = .Rows.Count
        CallerCols = .Columns.Count
    End With
    ReDim Result(CallerRows, CallerCols)
    'loop through the columns and generate the random samples
    Dim i As Integer
    Dim j As Integer
    Dim X() As Variant
    For i = 1 To CallerRows
        X = SampleSingleMultiVariateNormal(A)
        For j = 1 To CallerCols
            If j <= UBound(X) Then
                Result(i, j) = X(j)
            Else
                Result(i, j) = ""
            End If
        Next j
    Next i
    SampleMultiVariateNormal = Result
End Function

' Note need to select as many rows as there are variables to simulate

Function SampleMultivariateNormalT(CorrelationMatrixRange As Range, DegreesOfFreedom As Double) As Variant
    'read in the correlation matrix
    Dim CorrelationMatrix() As Variant
    CorrelationMatrix = CorrelationMatrixRange
    'check that we have a valid correlation matrix
    If CheckCorrelationMatrix(CorrelationMatrix) = False Then
        MsgBox ("The correlation matrix isn't symmetrical.")
        'SampleMultiVariateNormal = CVErr(xlErrValue)
        Exit Function
    End If
    'use cholesky matrix to get A such that AA' = CorrelationMatrix
    Dim A() As Variant
    A = CholeskyDecomposition(CorrelationMatrix)
    'prepare the array we want to return
    Dim CallerRows As Integer
    Dim CallerCols As Integer
    Dim Result() As Variant
    With Application.Caller
        CallerRows = .Rows.Count
        CallerCols = UBound(CorrelationMatrix, 1) '.Columns.Count
    End With
    ReDim Result(CallerRows, CallerCols)
    'loop through the columns and generate the random samples
    Dim i As Integer
    Dim j As Integer
    Dim X() As Variant
    Dim ChiSquare As Double
    For i = 1 To CallerRows
        X = SampleSingleMultiVariateNormal(A) 
        ChiSquare = WorksheetFunction.ChiInv(Rnd(), DegreesOfFreedom)
        For j = 1 To CallerCols
            If j <= UBound(X) Then
                ' The expected value of Chi-Square is its degree of freedom.
                ' However, the lower the degree of freedom, the more likely is that a random draw will be below the mean (since the distribution is more skewed)
                ' and so would scale up the normal random variable to crate heavier tails
                Result(i, j) = X(j) * Sqr(DegreesOfFreedom / ChiSquare )
            Else
                Result(i, j) = ""
            End If
        Next j
    Next i
    SampleMultivariateNormalT = Result
End Function



Public Function arrayrank(vArray() As Double) As Long()

    ' Usage: Y = arrayrank(vArray)
    '
    ' Returns an array of longs representing the the rank orders of the elements
    ' of vArray. If vArray is two-dimensional, then it returns an array with
    ' the same number of rows and columns with the columns containing the ranks
    ' of the corresponding columns of "vArray".

    Dim vDims As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim columns As Long
    Dim rows As Long
    Dim vRank() As Long
    Dim tmpRank() As Long

    ' Check that dimensionality of vArray is either 1 or 2
    vDims = NumberOfArrayDimensions(vArray)
    If Not (vDims = 1 Or vDims = 2) Then
        Err.Raise Number:=vbObjectError + 7, _
                Source:="arrayrank", _
                Description:="Input argument not one or two dimensional"
        Debug.Print "Arrayrank input argument not one or two dimensional"
        Debug.Print vDims
        Exit Function
    End If

    ' Check that array indices are numbered from 1
    If vDims = 1 Then
        If Not (LBound(vArray) = 1) Then
            Err.Raise Number:=vbObjectError + 8, _
                    Source:="arrayrank", _
                    Description:="One dimensional array input argument not numbered from 1"
            Debug.Print "arrayrank input argument start index = " & LBound(vArray)
            Exit Function
        End If
        rows = UBound(vArray)
    ElseIf vDims = 2 Then
        If Not (LBound(vArray, 1) = 1 And LBound(vArray, 2) = 1) Then
            Err.Raise Number:=vbObjectError + 9, _
                    Source:="arrayrank", _
                    Description:="Two dimensional array input argument not numbered from 1"
            Debug.Print "column index starts from" & LBound(vArray, 1) & " and row index starts from" & LBound(vArray, 1)
        End If
        rows = UBound(vArray, 1) - LBound(vArray, 1) + 1
        columns = UBound(vArray, 2) - LBound(vArray, 2) + 1
    End If

    ' Create array of consecutive longs from startIndex to endIndex
    If vDims = 1 Then
        ReDim vRank(1 To rows)
        ReDim tmpRank(1 To rows)
    ElseIf vDims = 2 Then
        ReDim vRank(1 To rows, 1 To columns)
        ReDim tmpRank(1 To rows)
    End If

    ' Pass this array to QuickSort along with the array to be ranked
    If vDims = 1 Then
        For i = 1 To rows
            tmpRank(i) = i
        Next i
            Call quicksort(keyArray:=vArray, otherArray:=tmpRank)
        For i = 1 To rows
            vRank(tmpRank(i)) = i
        Next i
    ElseIf vDims = 2 Then
        For j = 1 To columns
            For i = 1 To rows
                tmpRank(i) = i
            Next i
                Call quicksort(keyArray:=vArray, Column:=j, otherArray:=tmpRank)
            For i = 1 To rows
                vRank(tmpRank(i), j) = i
            Next i
        Next j
    End If

    arrayrank = vRank

End Function


Public Function quicksort(keyArray() As Double, Optional Column As Long, Optional otherArray)

    ' Usage: quicksort(keyArray, column, otherArray)
    '
    ' Sorts keyArray in place. If keyArray is two-dimensional (i.e. a matrix) then
    ' only the column specified by the optional argument "column" will be sorted.
    '
    ' An optional "otherArray" can be sorted in parallel, also in-place. If
    ' "otherArray" is used it must be one-dimensional and have the same start and
    ' end indices as the columns of "keyArray".
    '
    ' "keyArray" must be of type Double, "column" and otherArray must be of type
    ' long

    Dim keyDims As Long
    Dim otherDims As Long
    Dim inLow As Long
    Dim inHi As Long
    Dim i As Long

    ' TESTS

    ' Check that the dimensionality of "keyArray" is either 1 or 2
    keyDims = NumberOfArrayDimensions(keyArray)
    If Not (keyDims = 1 Or keyDims = 2) Then
        Debug.Print "input argument not one or two dimensional"
        Exit Function
    End If

    If Not IsMissing(otherArray) Then
    ' Check that "otherArray" is one-dimensional
        otherDims = NumberOfArrayDimensions(otherArray)
        If Not otherDims = 1 Then
            Debug.Print "'otherArray' not one-dimensional"
            Exit Function
        End If
        If keyDims = 1 Then
    ' Check that "keyArray" and "otherArray" are conformable
            If Not ( _
                        UBound(keyArray) = UBound(otherArray) And _
                        LBound(keyArray) = LBound(otherArray) _
                    ) Then
                    Debug.Print "'keyArray' and 'otherArray' not conformable"
                    Exit Function
            End If
        ElseIf keyDims = 2 Then
    ' Check that the argument "column" has been supplied
            If IsMissing(Column) Then
                Debug.Print "'column' argument not passed"
                Exit Function
            End If
    ' Check that "otherArray" is conformable to columns of "keyArray"
            If Not ( _
                    UBound(keyArray, 1) = UBound(otherArray) And _
                    LBound(keyArray, 1) = LBound(otherArray) _
                ) Then
                Debug.Print "'keyArray' and 'otherArray' not conformable"
                Exit Function
            End If
        End If
    End If

    ' Check that the argument "column" points to one of the columns of "keyArray"
    If Not (LBound(keyArray, 2) <= Column And Column <= UBound(keyArray, 2)) Then
        ' ERROR: Argument "column" does not point to one of the columns of "keyArray"
        Exit Function
    End If

    ' END OF TESTS

    ' Calculate "inHi" and "inLow"

    If keyDims = 1 Then
        inLow = LBound(keyArray)
        inHi = UBound(keyArray)
    ElseIf keyDims = 2 Then
        inLow = LBound(keyArray, 1)
        inHi = UBound(keyArray, 1)
    End If

    ' Call appropriate sort function

    If keyDims = 1 And IsMissing(otherArray) Then
        Call quick_sort_1ds(inLow, inHi, keyArray)
            
    ElseIf keyDims = 1 And Not IsMissing(otherArray) Then
        Call quick_sort_1dd(inLow, inHi, keyArray, otherArray)

    ElseIf keyDims = 2 And IsMissing(otherArray) Then
        Call quick_sort_2ds(inLow, inHi, keyArray, Column)

    ElseIf keyDims = 2 And Not IsMissing(otherArray) Then
        Call quick_sort_2dd(inLow, inHi, keyArray, Column, otherArray)

    End If

End Function


Private Function quick_sort_1ds(inLow, inHi, keyArray)

' Usage: quicksort(inLow, inHi, keyArray)
'
' "quick_sort_1ds" = quick sort one dimensional single array
'
' Sorts keyArray in place between indices inLow and inHi, keyArray must be of
' type Double
'
' This function is based on code suggested in:
' http://stackoverflow.com/questions/152319/vba-array-sort-function

Dim tmpLow As Long
Dim tmpHi As Long
Dim pivot As Double
Dim keyTmpSwap As Double

tmpLow = inLow
tmpHi = inHi
pivot = keyArray((inLow + inHi) \ 2)

Do While (tmpLow <= tmpHi)
    Do While keyArray(tmpLow) < pivot And tmpLow < inHi
        tmpLow = tmpLow + 1
    Loop
    Do While keyArray(tmpHi) > pivot And tmpHi > inLow
        tmpHi = tmpHi - 1
    Loop
    If (tmpLow <= tmpHi) Then
        keyTmpSwap = keyArray(tmpLow)
        keyArray(tmpLow) = keyArray(tmpHi)
        keyArray(tmpHi) = keyTmpSwap
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
    End If
Loop

If (inLow < tmpHi) Then quick_sort_1ds inLow, tmpHi, keyArray
If (tmpLow < inHi) Then quick_sort_1ds tmpLow, inHi, keyArray

End Function

Private Function quick_sort_1dd(inLow, inHi, keyArray, otherArray)

' Usage: quick_sort_1dd(inLow, inHi, keyArray, otherArray)
'
' "quick_sort_1dd" = quick sort one dimensional double array
'
' Sorts keyArray in place between indices inLow and inHi.
'
' An array, "otherArray" is sorted in parallel, also in-place. "otherArray" must
' be one dimensional and have the same start and end indices as the first dimen-
' sion of "keyArray".
'
' "keyArray" must be of type Double, "otherArray" must be of type long.
'
' This function is based on code suggested in:
' http://stackoverflow.com/questions/152319/vba-array-sort-function

Dim pivot As Double
Dim tmpLow As Long
Dim tmpHi As Long
Dim keyTmpSwap As Double
Dim otherTmpSwap As Long
Dim keyDims As Long
Dim otherDims As Long

tmpLow = inLow
tmpHi = inHi
pivot = keyArray((inLow + inHi) \ 2)

Do While (tmpLow <= tmpHi)

    Do While keyArray(tmpLow) < pivot And tmpLow < inHi
        tmpLow = tmpLow + 1
    Loop
    
    Do While keyArray(tmpHi) > pivot And tmpHi > inLow
        tmpHi = tmpHi - 1
    Loop
    
    If (tmpLow <= tmpHi) Then
    
        keyTmpSwap = keyArray(tmpLow)
        otherTmpSwap = otherArray(tmpLow)
        
        keyArray(tmpLow) = keyArray(tmpHi)
        otherArray(tmpLow) = otherArray(tmpHi)
        
        keyArray(tmpHi) = keyTmpSwap
        otherArray(tmpHi) = otherTmpSwap
        
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
        
    End If
    
Loop

If (inLow < tmpHi) Then quick_sort_1dd inLow, tmpHi, keyArray, otherArray
If (tmpLow < inHi) Then quick_sort_1dd tmpLow, inHi, keyArray, otherArray

End Function

Private Function quick_sort_2ds(inLow As Long, inHi As Long, keyArray() As Double, Column As Long)

' Usage: quick_sort_2ds(inLow, inHi, keyArray, column)
'
' "quick_sort_2ds" = quick sort two dimensional single array
'
' Sorts a column of "keyArray", in place ' between indices "inLow" and "inHi".
' "keyArray" must be is two-dimensional (i.e. a matrix). Only the column
' specified by the optional argument "column" will be sorted, the other columns
' are left untouched.
'
' "keyArray" must be of type Double, column and "column" must be of type
' long
'
' This function is based on code suggested in:
' http://stackoverflow.com/questions/152319/vba-array-sort-function

Dim pivot As Double
Dim tmpLow As Long
Dim tmpHi As Long
Dim keyTmpSwap As Double
Dim keyDims As Long

tmpLow = inLow
tmpHi = inHi
pivot = keyArray((inLow + inHi) \ 2, Column)

Do While (tmpLow <= tmpHi)
    Do While keyArray(tmpLow, Column) < pivot And tmpLow < inHi
        tmpLow = tmpLow + 1
    Loop
    Do While keyArray(tmpHi, Column) > pivot And tmpHi > inLow
        tmpHi = tmpHi - 1
    Loop
    If (tmpLow <= tmpHi) Then
        keyTmpSwap = keyArray(tmpLow, Column)
        keyArray(tmpLow, Column) = keyArray(tmpHi, Column)
        keyArray(tmpHi, Column) = keyTmpSwap
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
    End If
Loop

If (inLow < tmpHi) Then quick_sort_2ds inLow, tmpHi, keyArray, Column
If (tmpLow < inHi) Then quick_sort_2ds tmpLow, inHi, keyArray, Column

End Function

Private Function quick_sort_2dd(inLow, inHi, keyArray, Column, otherArray)

' Usage: quick_sort_2dd(inLow, inHi, keyArray, column, otherArray)
'
' "quick_sort_1dd" = quick sort two dimensional double array
'
' Sorts "keyArray" and "otherArray" in place between indices inLow and inHi.
'
' An array, "otherArray" is sorted in parallel, also in-place. "otherArray" must
' be one dimensional and have the same start and end indices as the first dimen-
' sion of "keyArray".
'
' "keyArray" must be of type Double, "column" and "otherArray" must be of type
' long.
'
' This function is based on code suggested in:
' http://stackoverflow.com/questions/152319/vba-array-sort-function

Dim pivot As Double
Dim tmpLow As Long
Dim tmpHi As Long
Dim keyTmpSwap As Double
Dim otherTmpSwap As Long
Dim keyDims As Long
Dim otherDims As Long

tmpLow = inLow
tmpHi = inHi
pivot = keyArray((inLow + inHi) \ 2, Column)

Do While (tmpLow <= tmpHi)

    Do While keyArray(tmpLow, Column) < pivot And tmpLow < inHi
        tmpLow = tmpLow + 1
    Loop
    
    Do While keyArray(tmpHi, Column) > pivot And tmpHi > inLow
        tmpHi = tmpHi - 1
    Loop
    
    If (tmpLow <= tmpHi) Then
    
        keyTmpSwap = keyArray(tmpLow, Column)
        otherTmpSwap = otherArray(tmpLow)
        
        keyArray(tmpLow, Column) = keyArray(tmpHi, Column)
        otherArray(tmpLow) = otherArray(tmpHi)
        
        keyArray(tmpHi, Column) = keyTmpSwap
        otherArray(tmpHi) = otherTmpSwap
        
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
        
    End If
    
Loop

If (inLow < tmpHi) Then quick_sort_2dd inLow, tmpHi, keyArray, Column, otherArray
If (tmpLow < inHi) Then quick_sort_2dd tmpLow, inHi, keyArray, Column, otherArray

End Function
