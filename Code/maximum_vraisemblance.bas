Attribute VB_Name = "maxi_de_vraisemblance"
Sub fill_carre_vraisemblance() 'initialiser mu_first2 et sig_first2
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Dim mu_arr() As Variant
    Dim sig_arr() As Variant
    Dim temp As Range
    
    mu_arr = Range([mu_first2], [mu_first2].End(xlToRight))
    mu_arr = WorksheetFunction.Transpose(mu_arr)
    sig_arr = Range([sig_first2], [sig_first2].End(xlDown))
    
    Dim the_fill() As Double
    ReDim the_fill(LBound(mu_arr) To UBound(mu_arr), LBound(sig_arr) To UBound(sig_arr))
    
    For i = LBound(mu_arr) To UBound(mu_arr)
        For j = LBound(sig_arr) To UBound(sig_arr)
            the_fill(i, j) = maximum_vraisemblance(CDbl(mu_arr(i, 1)), CDbl(sig_arr(j, 1)))
        Next j
    Next i
    Range([mu_first2].Offset(1, 0), [mu_first2].Offset(UBound(sig_arr), UBound(mu_arr) - LBound(mu_arr))) = WorksheetFunction.Transpose(the_fill)
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub


Function maximum_vraisemblance(mu As Double, sig As Double)
    Dim temp() As Variant
    temp = Range([first_value3].Offset(0, -1), [first_value3].Offset(0, -1).End(xlDown))
    
    Dim milieu_abscisse() As Double
    ReDim milieu_abscisse(LBound(temp) To UBound(temp))
    Dim frequence() As Double
    ReDim frequence(LBound(temp) To UBound(temp))
    Dim proba_loi_normale() As Double
    ReDim proba_loi_normale(LBound(temp) To UBound(temp))
    Dim vraisemb() As Double
    ReDim vraisemb(LBound(temp) To UBound(temp))
    
    
    
    milieu_abscisse = arr_to_arr(temp, milieu_abscisse)
    Erase temp
    temp = Range([first_value3], [first_value3].End(xlDown).Offset(-1, 0))
    frequence = arr_to_arr(temp, frequence)
    
    Call loi_normale(milieu_abscisse, proba_loi_normale, mu, sig)
    
    maximum_vraisemblance = populate_vraisemb(frequence, proba_loi_normale, vraisemb)
    
End Function

Function arr_to_arr(ByRef temp() As Variant, ByRef ret() As Double) As Double()
    Dim i As Integer
    For i = LBound(temp) To UBound(temp)
        ret(i) = temp(i, 1)
    Next i
    arr_to_arr = ret
End Function

Sub loi_normale(milieu_abs() As Double, proba_loi_norm() As Double, mu As Double, sig As Double)
    Dim i As Integer
    Dim moy As Double
    For i = LBound(milieu_abs) To UBound(milieu_abs)
        proba_loi_norm(i) = WorksheetFunction.Norm_Dist(milieu_abs(i), mu, sig, False)
        moy = moy + proba_loi_norm(i)
    Next i
    
    For i = LBound(milieu_abs) To UBound(milieu_abs)
        proba_loi_norm(i) = proba_loi_norm(i) / moy
    Next i
End Sub

Function populate_vraisemb(freq() As Double, proba() As Double, ret() As Double) As Double
    Dim i As Integer
    Dim somme As Double
    For i = LBound(freq) To UBound(freq)
        ret(i) = freq(i) * WorksheetFunction.Ln(proba(i))
        somme = somme + ret(i)
    Next i
    populate_vraisemb = somme
End Function

