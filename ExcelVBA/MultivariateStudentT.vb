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


Function SampleMultivariateNormalT(CorrelationMatrixRange As Range, DegreesOfFreedom As Double) As Variant
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
    ChiSquare = WorksheetFunction.ChiInv(Rnd, DegreesOfFreedom)
    For i = 1 To CallerRows
        X = SampleSingleMultiVariateNormal(A) / Sqr(ChiSquare / DegreesOfFreedom)
        For j = 1 To CallerCols
            If j <= UBound(X) Then
                Result(i, j) = X(j)
            Else
                Result(i, j) = ""
            End If
        Next j
    Next i
    SampleMultivariateNormalT = Result
End Function

