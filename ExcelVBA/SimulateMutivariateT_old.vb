' Need to fix this up so that it is more generic.
' It comes from Credit Risk Modelling with Excel VBA.pdf

Sub simVBAt()
Dim M As Long, N As Long, i As Long, j As Long, df As Long
M = Range("c3") ’Number of simulations
N = Application.Count(Range("B10:B65536")) ’Number of loans
df = Range("C4")
Dim d(), LGD() As Double, EAD() As Double, w() As Double, w2() As Double
Dimloss(),factorAsDouble,loss_jAsDouble, tadjustAsDouble
ReDim d(1 To N), LGD(1 To N), EAD(1 To N), w(1 To N), w2(1 To N), _ loss(1 To M)
’Write loan characteristics into arrays For i = 1 To N
d(i) = -Application.WorksheetFunction.TInv(Range("B" & i + 9) * 2, df) LGD(i) = Range("C" & i + 9)
EAD(i) = Range("D" & i + 9)
w(i) = Range("E" & i + 9)
w2(i) = ((1 − w(i) * w(i))) ˆ 0.5 Next i
’Conduct M Monte Carlo trials For j = 1 To M
factor = nrnd()
tadjust = (Application.WorksheetFunction.ChiInv(Rnd, df) / df) ˆ 0.5
√
             ’Compute portfolio loss for one trial loss_j = 0
For i = 1 To N
If (w(i) * factor + w2(i) * nrnd()) loss_j = loss_j + LGD(i) * EAD(i)
End If Next i
loss(j) = loss_j Next j
Sort loss For i = 3 To 7
/ tadjust
< d(i) Then
 Range("h" & i) = loss(Int((M+1) * Range("g" & i))) Next i
End Sub
