Function ShuffleArray(InArray() As Variant) As Variant()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ShuffleArray
' This function returns the values of InArray in random order. The original
' InArray is not modified.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim N As Long
    Dim Temp As Variant
    Dim J As Long
    Dim Arr() As Variant
    
    
    Randomize
    L = UBound(InArray) - LBound(InArray) + 1
    ReDim Arr(LBound(InArray) To UBound(InArray))
    For N = LBound(InArray) To UBound(InArray)
        Arr(N) = InArray(N)
    Next N
    For N = LBound(InArray) To UBound(InArray)
        J = CLng(((UBound(InArray) - N) * Rnd) + N)
        Temp = Arr(N)
        Arr(N) = ARr(J)
        Arr(J) = Temp
    Next N
    ShuffleArray = Arr
End Function

Sub ShuffleArrayInPlace(InArray() As Variant)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ShuffleArrayInPlace
' This shuffles InArray to random order, randomized in place.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim N As Long
    Dim Temp As Variant
    Dim J As Long
   
    Randomize
    For N = LBound(InArray) To UBound(InArray)
        J = CLng(((UBound(InArray) - N) * Rnd) + N)
        If N <> J Then
            Temp = InArray(N)
            InArray(N) = InArray(J)
            InArray(J) = Temp
        End If
    Next N
End Sub
