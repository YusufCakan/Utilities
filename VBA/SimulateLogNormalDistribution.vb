Function SampleLogNormal(Mean As Double, StDev As Double)
    Dim ScaledMean As Double
    Dim ScaledStDev As Double
    
    ScaledMean = Log(Mean ^ 2 / Sqr(Mean ^ 2 + StDev ^ 2))
    ScaledStDev = Sqr(Log((Mean ^ 2 + StDev ^ 2) / Mean ^ 2))
    SampleLogNormal = WorksheetFunction.LogNorm_Inv(Rnd(), ScaledMean, ScaledStDev)
End Function

