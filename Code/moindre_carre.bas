Attribute VB_Name = "ols"
Option Explicit
Option Base 1

Function moindre_carre(mu As Double, sig As Double) As Double ' il faut intialiser [first_value] en input
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    Dim temp() As Variant
    Dim arr_frequence() As Double
    Dim arr_proba_calcule() As Double
    Dim arr_proba_connu() As Double
    
    temp = Range([first_value].Offset(0, -2), [first_value].Offset(0, -2).End(xlDown))
    arr_frequence = arr_to_arr(temp)
    Erase temp
    temp = Range([first_value], [first_value].End(xlDown))
    arr_proba_connu = arr_to_arr(temp)
    Erase temp
    
    arr_proba_calcule = populate_arr_calcul(arr_frequence, mu, sig)
    
    moindre_carre = ols(arr_proba_connu, arr_proba_calcule)
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
End Function

Function arr_to_arr(arr() As Variant) As Double()
    Dim ret() As Double
    ReDim ret(LBound(arr) To UBound(arr))
    Dim i As Integer
    
    For i = LBound(arr) To UBound(arr)
        ret(i) = arr(i, 1)
    Next i
    arr_to_arr = ret
End Function

Function populate_arr_calcul(arr() As Double, mu As Double, sig As Double) As Double()
    Dim ret() As Double ' contient les probas des valeurs de la loi normale
    ReDim ret(LBound(arr) To UBound(arr))
    Dim temp() As Double ' contient les valeurs de la loi normale
    ReDim temp(LBound(arr) To UBound(arr))
    Dim i As Integer
    Dim moy As Double
    
    For i = LBound(arr) To UBound(arr)
        temp(i) = WorksheetFunction.Norm_Dist(arr(i), mu, sig, False)
        moy = moy + temp(i)
    Next i
    'moy = moy / (UBound(arr) - LBound(arr))
    
    For i = LBound(arr) To UBound(arr)
        ret(i) = temp(i) / moy
    Next i
    populate_arr_calcul = ret
    
End Function

Function ols(arr_connu() As Double, arr_calcul() As Double) As Double
    Dim i As Integer
    Dim temp() As Double
    ReDim temp(LBound(arr_calcul) To UBound(arr_calcul))
    Dim somme As Double
    
    For i = LBound(temp) To UBound(temp)
        temp(i) = (arr_calcul(i) - arr_connu(i)) ^ 2
        somme = somme + temp(i)
    Next i
    somme = somme ^ (1 / 2)
    ols = somme
End Function

