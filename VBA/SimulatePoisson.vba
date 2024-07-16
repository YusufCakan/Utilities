Function PoissonInv(dP As Double, dMu As Double) As Variant
 'POISSONINV is the inverse of the Poisson cumulative distribution function (cdf)
 '  with parameter mu. X = PoissonInv(P, MU) returns the smallest value of X,
 '  such that the Poisson cdf evaluated at X, equals or exceeds P.
 '
 '  This function will use a simply iterative method when MU is less than 10. When
 '  MU is greater than 10, an initial guess for X is first obtained from a normal
 '  approximation, after which the initial guess iterates up and down to ensure it
 '  is the correct answer.
 '
 'SYNTAX
 '  X = PoissonInv(P, MU) where P is the probability between 0 and 1 and MU is the
 '  expected number of events.
 '
 'EXAMPLE
 '  P = POISSON(8, 20, TRUE)        = 0.002087259
 '  X = POISSONINV(0.002087259, 20) = 8
 '
 '
 '  Author:    Andrew O'Connor <andrew.oconnor@relken.com>
 '  Date:      10 Jul 2014
 '  Copyright: 2014 Relken Engineering
 
 ' These variables are used to simplify this summation:
 ' dCDF = dCDF + Exp(-dMu) * dMu ^ iX / .Fact(iX)
 Dim iX As Long               ' number of events
 Dim dCDF As Double           ' cumulative distribution function of iX
 Dim dExpMu As Double         ' Exp(-dMu)
 Dim dTerm As Double          ' incremental term
 
 ' These terms are used to conduct a normal approximation of
 ' the Poisson Distribution
 Dim dX As Double             ' normal approximation for iX
 Dim dSigma As Double         ' Signma for normal approximation
 Dim dMuThreshold As Double   ' Threshold for u above which the normal approximation us used
 
 'Set Threshold after which a normal distribution is used
 dMuThreshold = 10
 
 'Determine method of calculation
 If (dP < 0 Or dP >= 1) Or dMu < 0 Then  'Raise error
     PoissonInv = CVErr(xlErrValue)
 
 ElseIf dMu > dMuThreshold Then  'Use normal approximation
     'Obtain initial estimate
     dSigma = Sqr(dMu)
     dX = WorksheetFunction.NormInv(dP, dMu, dSigma)
     iX = WorksheetFunction.Max(WorksheetFunction.RoundUp(dX, 0), 0)
     dCDF = WorksheetFunction.Poisson(iX, dMu, True)
 
     'If the approximation was lower than dP increase iX
     If dCDF < dP Then
        Do While dCDF < dP
           iX = iX + 1
           dCDF = dCDF + WorksheetFunction.Poisson(iX, dMu, False)
        Loop
     Else
     'If approximation was higher than dP, find smallest iX
        Do While dCDF >= dP
           dCDF = dCDF - WorksheetFunction.Poisson(iX, dMu, False)
           iX = iX - 1
        Loop
        'Take back last subtraction, plus add one to get dCDF < dP
        iX = iX + 1
     End If
 
 Else 'Use iterative approach
     'Prepare calculation variables
     dExpMu = Exp(-dMu)
     dTerm = dExpMu
 
     'Loop through each iteration until the required probability
     'is obtained. The number of loops is the answer
     Do While dCDF < dP
        'Update cumulative function
        dCDF = dCDF + dTerm
        'Add iteration
        iX = iX + 1
        'Update addition term
        dTerm = dTerm * dMu / iX
     Loop
     'Take back the last addition
     iX = iX - 1
 
 End If
 
 PoissonInv = iX
 
 End Function
