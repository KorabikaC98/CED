Module InteractionCurve
    'these console application calucate the points of the Interaction Curve
    '#Functions:
    '_Calucating Omiga
    Function OMG(ByVal y As Double, Optional Nu As Double = 0, Optional h As Double = 0, Optional b As Double = 0, Optional fc As Double = 0) As Double
        Dim omg_val As Double
        If Nu Then
            Try
                omg_val = 0.9 - 0.5 * (Nu) / (0.85 * fc * b * h)
            Catch ex As Exception
                omg_val = 0.65
            End Try
        Else
            Try
                omg_val = 0.9 - 0.5 * y / h
            Catch ex As Exception
                omg_val = 0.65
            End Try
        End If
        If omg_val < 0.65 Then omg_val = 0.65
        If omg_val > 0.9 Then omg_val = 0.9
        Return omg_val
    End Function

    '_Calucating yb
    Function y_balance(ByVal d As Double, ByVal fy As Double, Optional Beta As Double = 0.85) As Double
        Return (630 * d / (630 + fy))
    End Function

    '_Calucating Nu/Omg
    Function NuOnOmg(ByVal fc As Double, ByVal Ac As Double, ByVal fst As Double, ByVal Ast As Double, ByVal Asc As Double, ByVal fsc As Double) As Double
        Return (0.85 * fc * Ac + Asc * fsc + Ast * fst)
    End Function

    '_Calucating Mu/Omg
    Function MuOnOmg(ByVal fc As Double, ByVal Ac As Double, ByVal fst As Double, ByVal Ast As Double, ByVal Asc As Double, ByVal fsc As Double, ByVal y As Double, ByVal h As Double, ByVal a As Double, ByVal dd As Double) As Double
        Return (0.85 * fc * Ac * 0.5 * (h - y) + fsc * Asc * (0.5 * h - dd) + fst * Ast * (0.5 * h - a))
    End Function

    '_Calucating fst
    Function Fs_t(ByVal fy_t As Double, ByVal y As Double, ByVal d As Double) As Double
        Dim fst As Double : fst = 630 * ((1 - 0.85 * d) / y)
        If fst > fy_t Then
            fst = fy_t
        ElseIf fst < (-1 * fy_t) Then
            fst = -1 * fy_t
        End If
        Return fst
    End Function

    '_Calucating fsc
    Function Fc_t(ByVal fy_c As Double, ByVal y As Double, ByVal dd As Double) As Double
        Dim fsc As Double : fsc = 630 * ((1 - 0.85 * dd) / y)
        If fsc > fy_c Then
            fsc = fy_c
        ElseIf fsc < (-1 * fy_c) Then
            fsc = -1 * fy_c
        End If
        Return fsc
    End Function

    'Calucation Of y


    Sub Main()
        '#Varaibles:
        '_Input
        Dim b, h, Ast, Asc, fyt, fyc, fc, a, dd As Double
        '_Calucation Varaibles
        Dim d, Ac, fst, fsc, y, a1, b1, c1, a2, b2, c2 As Double
        '_Output Variables
        Dim Mu, Nu As Double
        '#intialization values:
        b = 10 : h = 20 : a = 50 : dd = 50 'mm
        fc = 20 : fyt = 240 : fyc = 240 'Mpa
        Asc = 1800 : Ast = 1800 'mm2

        'a algorithm to calucate the Nu/Omg and Mu/Omg




    End Sub
End Module



