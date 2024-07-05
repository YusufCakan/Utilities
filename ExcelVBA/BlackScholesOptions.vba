Function BS_Price(stock_price As Double, strike_price As Double, Vol As Double, interest_rate As Double, dividend_yield As Double, term_to_maturity As Double, call_or_put_option_type As String)
    Dim d_1, d_2 As Double
    d_1 = (WorksheetFunction.Ln(stock_price \ strike_price) + (interest_rate - dividend_yield + Vol * Vol \ 2) * term_to_maturity) \ (Vol * Sqr(term_to_maturity))
    d_2 = d_1 - Vol * Sqr(term_to_maturity)

    If call_or_put_option_type = "Call" Then
        BS_Price = stock_price * Exp(-dividend_yield * term_to_maturity) * WorksheetFunction.Norm_S_Dist(d_1, True) - strike_price * Exp(-interest_rate * term_to_maturity) * WorksheetFunction.Norm_S_Dist(d_2, True)
    ElseIf call_or_put_option_type = "Put" Then
        BS_Price = strike_price * Exp(-interest_rate * term_to_maturity) * WorksheetFunction.Norm_S_Dist(-d_2, True) - stock_price * Exp(-dividend_yield * term_to_maturity) * WorksheetFunction.Norm_S_Dist(-d_1, True)
    Else
        BS_Price = "Error"
    End If
End Function


Function BS_Delta(stock_price As Double, strike_price As Double, Vol As Double, interest_rate As Double, dividend_yield As Double, term_to_maturity As Double, call_or_put_option_type As String)
    Dim d_1, d_2 As Double
    d_1 = (WorksheetFunction.Ln(stock_price \ strike_price) + (interest_rate - dividend_yield + Vol * Vol \ 2) * term_to_maturity) \ (Vol * Sqr(term_to_maturity))
    d_2 = d_1 - Vol * Sqr(term_to_maturity)

    If call_or_put_option_type = "Call" Then
        BS_Delta = Exp(-dividend_yield * term_to_maturity) * WorksheetFunction.Norm_S_Dist(d_1, True)
    ElseIf call_or_put_option_type = "Put" Then
        BS_Delta = -Exp(-dividend_yield * term_to_maturity) * WorksheetFunction.Norm_S_Dist(-d_1, True)
    Else
        BS_Delta = "Error"
    End If
End 


Function BS_Gamma(stock_price As Double, strike_price As Double, Vol As Double, interest_rate As Double, dividend_yield As Double, term_to_maturity As Double)
    Dim d_1, d_2 As Double
    d_1 = (WorksheetFunction.Ln(stock_price \ strike_price) + (interest_rate - dividend_yield + Vol * Vol \ 2) * term_to_maturity) \ (Vol * Sqr(term_to_maturity))
    d_2 = d_1 - Vol * Sqr(term_to_maturity)
    BS_Gamma = Exp(-dividend_yield * term_to_maturity) * WorksheetFunction.Norm_S_Dist(d_1, False) \ (stock_price * Vol * Sqr(term_to_maturity))
End Function


Function BS_Theta(stock_price As Double, strike_price As Double, Vol As Double, interest_rate As Double, dividend_yield As Double, term_to_maturity As Double, call_or_put_option_type As String)
    Dim d_1, d_2 As Double
    d_1 = (WorksheetFunction.Ln(stock_price \ strike_price) + (interest_rate - dividend_yield + Vol * Vol \ 2) * term_to_maturity) \ (Vol * Sqr(term_to_maturity))
    d_2 = d_1 - Vol * Sqr(term_to_maturity)

    If call_or_put_option_type = "Call" Then
        BS_Theta = -Exp(-dividend_yield * term_to_maturity) * stock_price * WorksheetFunction.Norm_S_Dist(d_1, False) * Vol \ (2 * Sqr(term_to_maturity)) - interest_rate * strike_price * Exp(-interest_rate * term_to_maturity) * WorksheetFunction.Norm_S_Dist(d_2, True) + dividend_yield * stock_price * Exp(-dividend_yield * term_to_maturity) * WorksheetFunction.Norm_S_Dist(d_1, True)
    ElseIf call_or_put_option_type = "Put" Then
        BS_Theta = -Exp(-dividend_yield * term_to_maturity) * stock_price * WorksheetFunction.Norm_S_Dist(d_1, False) * Vol \ (2 * Sqr(term_to_maturity)) + interest_rate * strike_price * Exp(-interest_rate * term_to_maturity) * WorksheetFunction.Norm_S_Dist(-d_2, True) - dividend_yield * stock_price * Exp(-dividend_yield * term_to_maturity) * WorksheetFunction.Norm_S_Dist(-d_1, True)
    Else
        BS_Theta = "Error"
    End If
End Function


Function BS_Vega(stock_price As Double, strike_price As Double, Vol As Double, interest_rate As Double, dividend_yield As Double, term_to_maturity As Double)
    Dim d_1, d_2 As Double
    d_1 = (WorksheetFunction.Ln(stock_price \ strike_price) + (interest_rate - dividend_yield + Vol * Vol \ 2) * term_to_maturity) \ (Vol * Sqr(term_to_maturity))
    d_2 = d_1 - Vol * Sqr(term_to_maturity)
    BS_Vega = stock_price * Exp(-dividend_yield * term_to_maturity) * WorksheetFunction.Norm_S_Dist(d_1, False) * Sqr(term_to_maturity)
End Function