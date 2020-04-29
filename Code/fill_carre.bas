Attribute VB_Name = "fill"
Sub fill_carre()
    Dim mu_arr() As Variant
    Dim sig_arr() As Variant
    Dim temp As Range
    
    mu_arr = Range([mu_first], [mu_first].End(xlToRight))
    mu_arr = WorksheetFunction.Transpose(mu_arr)
    sig_arr = Range([sig_first], [sig_first].End(xlDown))
    
    Dim the_fill() As Double
    ReDim the_fill(LBound(mu_arr) To UBound(mu_arr), LBound(sig_arr) To UBound(sig_arr))
    
    For i = LBound(mu_arr) To UBound(mu_arr)
        For j = LBound(sig_arr) To UBound(sig_arr)
            the_fill(i, j) = Module1.moindre_carre(CDbl(mu_arr(i, 1)), CDbl(sig_arr(j, 1)))
        Next j
    Next i
    Range([mu_first].Offset(1, 0), [mu_first].Offset(UBound(sig_arr), UBound(mu_arr) - LBound(mu_arr))) = WorksheetFunction.Transpose(the_fill)
End Sub
