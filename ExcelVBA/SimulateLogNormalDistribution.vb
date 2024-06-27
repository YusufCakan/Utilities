Function GenerateLognormalSimulations(ByVal mean As Double, ByVal coefficient_of_variation As Double, ByVal numSimulations As Long) As Variant
    Dim i As Long
    Dim stdDev As Double
    Dim lognormalSamples() As Double
    
    ' Calculate the standard deviation from the coefficient of variation (CV)
    stdDev = mean * coefficient_of_variation
    
    ' Resize the array to store the lognormal samples
    ReDim lognormalSamples(1 To numSimulations)
    
    ' Generate lognormal samples
    For i = 1 To numSimulations
        lognormalSamples(i) = Exp(stdDev * WorksheetFunction.NormInv(Rnd(), mean, stdDev))
    Next i
    
    ' Return the array of lognormal samples
    GenerateLognormalSimulations = lognormalSamples
End Function
