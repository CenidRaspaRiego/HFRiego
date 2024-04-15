Attribute VB_Name = "Hidraulicas"
Option Private Module
'ULTIMA MODIFICACIÓN FEBRERO DE 2022



' *********************************************************************
    ' Module  : HFriego
    ' Purpose : Calculo hidráulico en tuberias ciegas y con salidas multiples.
'
' Notes   : 
' *********************************************************************

' ---------------------------------------------------------------------
' Date        Developer                   Action
    ' 2012  Sergio Jiménez          Created

Option Explicit



Function PerdidaX2(gasto, diametro, longitud As Double) As Double
'Determina la perdida por fricción en la tuberia con PVC
Dim Coefp, DIp, Qp, Longp, Hfp As Double
    Qp = gasto / 1000
    Longp = longitud
    DIp = diametro / 1000
    Coefp = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("E1").Value
    If Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B1").Value = 1 Then
        Hfp = 10.648 * 1 / (Coefp) ^ (1.852) * (Qp) ^ (1.852) / (DIp) ^ (4.871) * Longp
    ElseIf Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B1").Value = 2 Then
        Hfp = 10.3 * (Coefp) ^ (2) * (Qp) ^ (2) / (DIp) ^ (16 / 3) * Longp
    ElseIf Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B1").Value = 3 Then
        Hfp = 0.004098 * Coefp * (Qp) ^ (1.9) / (DIp) ^ (4.9) * Longp
    End If

    PerdidaX = 1
    
End Function

Function FChristiansen(NS As Integer) As Double
'Determina el Factor para salidas multiples
Rem REVISAR Y CORREGIR 03 DE
Dim N As Double

    If Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B1").Value = 1 Then
    N = 1.852
    ElseIf Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B1").Value = 2 Then
    N = 2
    ElseIf Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B1").Value = 3 Then
    N = 1.9
    Else
    N = 2
    End If
    If NS > 0 Then
        FChristiansen = 1 / (N + 1) + 1 / (2 * NS) + (Sqr(N - 1)) / (6 * (NS) ^ 2)
    Else
        FChristiansen = CVErr(xlErrDiv0)
    End If
End Function
Function FJensen(NS As Integer) As Double
'Determina el Factor para salidas multiples con jensen

Dim N As Double

    If Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B1").Value = 1 Then
    N = 1.852
    ElseIf Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B1").Value = 2 Then
    N = 2
    ElseIf Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B1").Value = 3 Then
    N = 1.9
    Else
    N = 2
    End If
    If NS > 0 Then
        FJensen = (2 * NS / (2 * NS - 1)) * (1 / (N + 1) + (Sqr(N - 1)) / (6 * (NS) ^ 2))
    Else
        FJensen = CVErr(xlErrDiv0)
    End If
End Function
Function FScaloppi(NS As Integer, s, So As Double) As Double
'Determina el Factor para salidas multiples con jensen

Dim N, f1, Rs As Double

    If Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B1").Value = 1 Then
    N = 1.852
    ElseIf Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B1").Value = 2 Then
    N = 2
    ElseIf Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B1").Value = 3 Then
    N = 1.9
    Else
    N = 2
    End If
    If NS > 0 Then
        f1 = 1 / (N + 1) + 1 / (2 * NS) + (Sqr(N - 1)) / (6 * (NS) ^ 2)
        Rs = So / s
        FScaloppi = (NS * f1 + Rs - 1) / (NS + Rs - 1)
    Else
        FScaloppi = CVErr(xlErrDiv0)
    End If
End Function
Function Qminimoxseccion(area As Double, ETc As Double, Trd As Double)
'Calcula el caudal minimo necesario para regar una sección de riego en un area total
area = area * 10000
ETc = ETc / 1000
Qminimoxseccion = (area * ETc) / (Trd * 3.6)
End Function

Function LaminaHoraria(Qemisor As Double, SepEmisor As Double, SepRegantes As Double)
'Estima la lamina horaria aplicada por el emisor
Dim areamojada As Double
    areamojada = SepEmisor * SepRegantes
    LaminaHoraria = Qemisor / (areamojada)
End Function
Function LaminaHorariaGoteo(Qemisor As Double, SepEmisor As Double, SepRegantes As Double, areaMojado As Double)
'Estima la lamina horaria aplicada por el emisor
Dim areamojada As Double
    areamojada = SepEmisor * SepRegantes
    LaminaHorariaGoteo = Qemisor / (areamojada) / (areaMojado / 100)
End Function
Public Function LaminaHorariaMicro(diametro, gasto As Double)
Dim area, q As Double
area = WorksheetFunction.pi * (diametro) ^ 2 / 4 'm2
q = gasto / 1000 ' m3/hora
LaminaHorariaMicro = q / area * 1000
End Function
Function Qtotalreq(lh As Double, area As Double)
    lh = lh / 1000
    area = area * 10000
    Qtotalreq = (lh * area) / 3.6
End Function
Function dinterno(diametro As Double) As Double
'Estima el diametro interno en función del diametro nominal y redondea
'al diametro siguiente en función del valor
Dim DN(16) As Double
    DN(1) = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A4").Value
    DN(2) = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A5").Value
    DN(3) = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A6").Value
    DN(4) = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A7").Value
    DN(5) = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A8").Value
    DN(6) = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A9").Value
    DN(7) = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A10").Value
    DN(8) = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A11").Value
    DN(9) = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A12").Value
    DN(10) = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A13").Value
    DN(11) = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A14").Value
    DN(12) = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A15").Value
    DN(13) = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A16").Value
    DN(14) = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A17").Value
    DN(15) = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A18").Value
    DN(16) = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A19").Value

Select Case diametro
   Case 1 To DN(1)
        'dinterno = 18.1
        dinterno = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B4").Value
   Case DN(1) + 0.0000001 To DN(2)
        dinterno = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B5").Value
   Case DN(2) + 0.0000000001 To DN(3)
        dinterno = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B6").Value
   Case DN(3) + 0.0000000001 To DN(4)
        dinterno = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B7").Value
   Case DN(4) + 0.0000000001 To DN(5)
        dinterno = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B8").Value
   Case DN(5) + 0.0000000001 To DN(6)
        dinterno = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B9").Value
   Case DN(6) + 0.0000000001 To DN(7)
        dinterno = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B10").Value
   Case DN(7) + 0.0000000001 To DN(8)
        dinterno = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B11").Value
   Case DN(8) + 0.0000000001 To DN(9)
        dinterno = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B12").Value
   Case DN(9) + 0.0000000001 To DN(10)
        dinterno = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B13").Value
   Case DN(10) + 0.0000000001 To DN(11)
        dinterno = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B14").Value
   Case DN(11) + 0.0000000001 To DN(12)
        dinterno = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B15").Value
   Case DN(12) + 0.0000000001 To DN(13)
        dinterno = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B16").Value
   Case DN(13) + 0.0000000001 To DN(14)
        dinterno = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B17").Value
   Case DN(14) + 0.0000000001 To DN(15)
        dinterno = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B18").Value
   Case DN(15) + 0.0000000001 To DN(16)
        dinterno = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B19").Value
End Select
End Function
Function Dcalculado(caudal As Double)
'Sugiere el valor del diametro de un tubo en función del caudal a pasar por el
'no toma en cuenta para nada el valor de la fricción
Dim DiaMat As Double
caudal = caudal / 1000
DiaMat = (Sqr(caudal)) * 0.9213 * 1000
Select Case DiaMat
   Case 1 To 18.1
        Dcalculado = 13
   Case 18.1 To 22.7
      Dcalculado = 19
   Case 22.7 To 30.4
      Dcalculado = 25
   Case 30.4 To 39
      Dcalculado = 32
   Case 39 To 55.7
      Dcalculado = 50
   Case 55.7 To 82.1
      Dcalculado = 75
    Case 82.1 To 105.5
      Dcalculado = 100
    Case 105.5 To 154.4
      Dcalculado = 160
    Case 154.4 To 193
      Dcalculado = 200
    Case 193 To 241.2
      Dcalculado = 250
    Case 241.2 To 303.8
        Dcalculado = 315
    Case 303.8 To 342.6
        Dcalculado = 355
    Case 342.6 To 386
        Dcalculado = 400
    Case 386 To 434.2
        Dcalculado = 450
    Case 434.2 To 482.4
        Dcalculado = 500
    Case 482.4 To 607.8
        Dcalculado = 630
End Select

End Function

Function TirNormal(gasto, Ancho, Pendiente, Talud, Manning As Double)
'Sugiere el valor del diametro de un tubo en función del caudal a pasar por el
'no toma en cuenta para nada el valor de la fricción
        Dim yn, pn1, an, Pn, RH, pn2, pn3, pn4, ptn, Theta, So As Double
        Theta = Math.Atn(Pendiente)
        So = Pendiente
        If Ancho = 0 Then
        yn = ((gasto * Manning * (2 * (1 + Talud * Talud) ^ 0.5) ^ (2 / 3)) / ((So) ^ 0.5 * (Talud) ^ (5 / 3))) ^ (3 / 8)
        Else
        If So >= 0.001 Then
        yn = 1
        Else
        yn = 50
        End If
        pn1 = (gasto * Manning) / (So) ^ (0.5)
        an = (Ancho + Talud * yn) * yn
        Pn = Ancho + 2 * yn * ((Talud) ^ (2) + 1) ^ (0.5)
        RH = (an) / (Pn)
        pn2 = an * (RH) ^ (2 / 3)
        pn3 = 2 * (RH) ^ (2 / 3) - 2 * (RH) ^ (5 / 3) * (1 + (Talud) ^ (2)) ^ (0.5)
        Do While Abs(pn1 - pn2) >= 0.00011
            pn1 = (gasto * Manning) / (So) ^ (0.5)
            an = (Ancho + Talud * yn) * yn
            Pn = Ancho + 2 * yn * ((Talud) ^ (2) + 1) ^ (0.5)
            RH = (an) / (Pn)
            pn2 = an * ((RH) ^ (2 / 3))
            pn3 = 2 * Ancho + 2 * Talud * yn * (RH) ^ (2 / 3) - 2 * (RH) ^ (5 / 3) * (1 + (Talud) ^ (2)) ^ (0.5)
            ptn = (pn2 - pn1) / pn3
            yn = yn - ptn
        Loop
        End If
TirNormal = yn
End Function

Function LongMaxRegante(GastoEmisor, s, hf, diametro As Double)
Dim Coeficiente, DI, nmi, a, b, C, Sem, q As Double
Dim L, Qt, a1, B1, C1, F, HP, ah, res, Rey, fdw As Double
    
    Coeficiente = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("E1").Value
    DI = diametro / 1000
    Sem = s
    q = GastoEmisor
    HP = hf
    If Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B1").Value = 1 Then
        nmi = 1.852
        a = 0
        b = 10000
        ah = 10.648 / ((Coeficiente) ^ 1.852 * (DI) ^ 4.871 * (3600000) ^ 1.852)
        res = 1
        Do While Abs(res) > 0.0000001
            'calculamos C
            C = (a + b) / 2
            'Evaluamos en C
            L = Sem * C 'calculo de la longitud por cada salida
            Qt = (q * C) ^ (nmi) ' calculo del gasto por cada  salida
            a1 = 1 / (nmi + 1)
            B1 = 1 / (2 * C)
            C1 = (nmi - 1) ^ 0.5 / (6 * (C) ^ 2)
            F = a1 + B1 + C1 'calculo del Factor de salidas Multiples
            res = ah * F * L * Qt - HP 'calculo de la perdida de carga
                If res > 0 Then
                    b = C
                Else
                    a = C
                End If
        Loop
    
    ElseIf Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B1").Value = 2 Then
    
        nmi = 2
        ah = 10.3 * (Coeficiente) ^ 2 / ((DI) ^ (16 / 3) * (3600000) ^ 2)
        a = 0
        b = 10000
        res = 1
        Do While Abs(res) > 0.0000001
            C = (a + b) / 2
            L = Sem * C ' calculo de la longitud por cada salida
            Qt = (q * C) ^ (nmi) ' calculo del gasto por cada  salida
            a1 = 1 / (nmi + 1)
            B1 = 1 / (2 * C)
            C1 = (nmi - 1) ^ 0.5 / (6 * (C) ^ 2)
            F = a1 + B1 + C1 'calculo del Factor de salidas Multiples
            res = ah * F * L * Qt - HP 'calculo de la perdida de carga
                If res > 0 Then
                    b = C
                Else
                    a = C
                End If
        Loop
    
    ElseIf Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B1").Value = 3 Then
        nmi = 1.9
        ah = 4.098 * (10) ^ (-3) * Coeficiente / ((DI) ^ (4.9) * (3600000) ^ nmi)
        a = 0
        b = 10000
        res = 1
        Do While Abs(res) > 0.0000001
            C = (a + b) / 2
            L = Sem * C ' calculo de la longitud por cada salida
            Qt = (q * C) ^ (nmi) ' calculo del gasto por cada  salida
            a1 = 1 / (nmi + 1)
            B1 = 1 / (2 * C)
            C1 = (nmi - 1) ^ 0.5 / (6 * (C) ^ 2)
            F = a1 + B1 + C1 'calculo del Factor de salidas Multiples
            res = ah * F * L * Qt - HP 'calculo de la perdida de carga
                If res > 0 Then
                    b = C
                Else
                    a = C
                End If
        Loop
    ElseIf Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B1").Value = 4 Then
        nmi = 2
        ah = 0.0827 / ((DI) ^ (5) * (3600000) ^ 2)
        a = 0
        b = 10000
        res = 1
        Do While Abs(res) > 0.00000001
            C = (a + b) / 2
            L = Sem * C ' calculo de la longitud por cada salida
            Qt = (q * C) ^ (nmi) ' calculo del gasto por cada  salida
            a1 = 1 / (nmi + 1)
            B1 = 1 / (2 * C)
            C1 = (nmi - 1) ^ 0.5 / (6 * (C) ^ 2)
            F = a1 + B1 + C1 'calculo del Factor de salidas Multiples
            Rey = Workbooks("RegisterU2DF7.xlam").NReynoldsP((q * C) / 3600, DI * 1000) * 1
                If Rey <= 2000 Then
                        fdw = 64 / Rey
                Else
                    If Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("E2").Value = 0 Then
                        fdw = Workbooks("RegisterU2DF7.xlam").CoeFriccionDWP(Rey, Coeficiente, DI * 1000)
                    Else
                        fdw = Workbooks("RegisterU2DF7.xlam").CoeFriccionSJ(Rey, Coeficiente, DI * 1000)
                    End If
                End If
            res = ah * fdw * F * L * Qt - HP 'calculo de la perdida de carga
                If res > 0 Then
                    b = C
                Else
                    a = C
                End If
        Loop
    
    
    End If
    LongMaxRegante = L
    

End Function

Function PrecipitacionEfectiva(Precipitacion As Double)
Dim P, PE As Double
    P = Precipitacion
    If Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("B3").Value = "Porcentaje fijo" Then
        PE = P * (Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("B9").Value) / 100
    
    ElseIf Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("B3").Value = "Precipitacion Confiable" Then
        If P <= 70 Then
            PE = 0.6 * P - 10
        Else
            PE = 0.8 * P - 24
        End If
    ElseIf Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("B3").Value = "Formula empirica" Then
        If P <= (Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("H6").Value) Then
            PE = (Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("B6").Value) * P + ((Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("E6").Value)) * 1
        Else
            PE = (Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("B7").Value) * P + ((Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("E7").Value)) * 1
        End If
    ElseIf Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("B3").Value = "USDA" Then
        If P <= 250 Then
            PE = (P * (125 - 0.2 * P)) / 125
        Else
            PE = 0.1 * P + 125
        End If
    Else
        PE = 0
    End If
    If PE <= 0 Then
    PrecipitacionEfectiva = 0
    Else
    PrecipitacionEfectiva = PE
    End If
End Function

Function EvapotranspiracionA(Evaporacion, VelocidadViento, HumedadRelativa, CoberturaTanque As Double)

    Dim Ev, U2, HR, d, Kt, Eto As Double
    Ev = Evaporacion
    U2 = VelocidadViento
    HR = HumedadRelativa
    d = CoberturaTanque
    If Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B63").Value = 1 Then
        U2 = U2 * 86400 / 1000
        Kt = 0.475 - 0.00024 * U2 + 0.00516 * HR + 0.00118 * d - 0.000016 * (HR) ^ 2 - 0.101 * (10) ^ -5 * (d) ^ 2 - 0.8 * (10) ^ -8 * (HR) ^ 2 * U2 - 1 * (10) ^ -8 * (HR) ^ 2 * d
    ElseIf Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B63").Value = 2 Then
        Kt = 0.108 - 0.0286 * U2 + 0.0422 * WorksheetFunction.Ln(d) + 0.1434 * WorksheetFunction.Ln(HR) - 0.000631 * (WorksheetFunction.Ln(d)) ^ 2 * WorksheetFunction.Ln(HR)
    Else
        Kt = 0.61 + 0.00341 * HR - 0.000162 * U2 * HR - 0.00000959 * U2 * d + 0.00327 * U2 * WorksheetFunction.Ln(d) - 0.00289 * U2 * WorksheetFunction.Ln(86.4 * d) - 0.0106 * WorksheetFunction.Ln(86.4 * U2) * WorksheetFunction.Ln(d) + 0.00063 * (WorksheetFunction.Ln(d)) ^ 2 * WorksheetFunction.Ln(86.4 * U2)
    End If
    Eto = Kt * Ev
    EvapotranspiracionA = Eto
End Function

Function PotenciaBomba(gasto, Presion, EfiBomba, EfiMotor As Double)
    Dim Qa, Pr, EB, EM, pot As Double
    Qa = gasto
    Pr = Presion
    EB = EfiBomba / 100
    EM = EfiMotor / 100
    pot = Qa * Pr / (76 * EB * EM)
    PotenciaBomba = pot
End Function

Function perdida(gasto, diametro, longitud As Double)
'Determina la perdida por fricción en la tuberia con PVC
Dim Coefp, DIp, Qp, Longp, Hfp, Rey, fdw As Double
    Qp = gasto / 1000
    Longp = longitud
    DIp = diametro / 1000
    Coefp = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("E1").Value
    If Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B1").Value = 1 Then
        Hfp = 10.648 * 1 / (Coefp) ^ (1.852) * (Qp) ^ (1.852) / (DIp) ^ (4.871) * Longp
    ElseIf Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B1").Value = 2 Then
        Hfp = 10.3 * (Coefp) ^ (2) * (Qp) ^ (2) / (DIp) ^ (16 / 3) * Longp
    ElseIf Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B1").Value = 3 Then
        Hfp = 0.004098 * Coefp * (Qp) ^ (1.9) / (DIp) ^ (4.9) * Longp
    ElseIf Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B1").Value = 4 Then
        Rey = Workbooks("RegisterU2DF7.xlam").NReynoldsP(Qp * 1000, DIp * 1000) * 1
        If Rey <= 2000 Then
            fdw = 64 / Rey
        Else
                If Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("E2").Value = 0 Then
                    fdw = Workbooks("RegisterU2DF7.xlam").CoeFriccionDWP(Rey, Coefp, DIp * 1000) * 1
                Else
                    fdw = Workbooks("RegisterU2DF7.xlam").CoeFriccionSJ(Rey, Coefp, DIp * 1000) * 1
                End If
        End If
        
        Hfp = 0.0827 * fdw * (Qp) ^ (2) / (DIp) ^ (5) * Longp
    End If
    
    


    perdida = Hfp
    
End Function
'ULTIMA MODIFICACIÓN 12 DE septiembre DE 2014 by sergio ivan jimenez jimenez
Public Function GalonphLitroph(GPH As Double) As Double
'CONVIERTE LAS UNIDADES DE GALONES POR HORA A LITROS POR HORA
If GPH = 0 Then
GalonphLitroph = CVErr(xlErrDiv0)
End If
If IsNumeric(GPH > 0) = True Then
GalonphLitroph = 3.7854118 * GPH
End If
End Function

Public Function LitrophGalonph(LPH As Double) As Double
'CONVIERTE LAS UNIDADES DE LITROS POR HORA A GALONES POR HORA
If LPH = 0 Then
LitrophGalonph = CVErr(xlErrDiv0)
End If
If IsNumeric(LPH > 0) = True Then
LitrophGalonph = LPH / 3.7854118
End If
End Function

'Eto Priestley Taylor model
Public Function EToPriestleTaylor(Juliano, Tmax, Tmin, Tmean, RH, Rs, elevation, Latitud) As Double
Dim LE, a, cs, P, Es, Ea, Ra, Rnl, RT, Rn, AT, Rso, Rns As Double
'Tmean = (Tmax + Tmin) / 2
    If Juliano > 366 Or Juliano <= 0 Or elevation < 0 Then
        EToPriestleTaylor = 0 / 0
    ElseIf Tmax < Tmin Or RH > 100 Then
        EToPriestleTaylor = 0 / 0
    Else
        LE = LatentHVaporization(Tmean)
        a = SlopeSaturationVP(Tmean)
        P = AtmosphericP(elevation)
        cs = PsychrometricC(Tmean, P)
        Ea = ActualVaporP(RH, Tmean)
        Ra = RadiacionExtraterrestres(Juliano, Latitud)
        Rso = ClearSkySR(elevation, Ra)
        Rns = NetShortwaveR(0.23, Rs)
        Rnl = Rnlarga(Tmax, Tmin, Ea, Rs, Rso)
            If Rnl <= 0 Then
                Rnl = 0
            Else
            End If
        Rn = Rns - Rnl
        EToPriestleTaylor = 1 / LE * 1.26 * (a) / (a + cs) * (Rn - 0)
    End If
End Function
Public Function EToPM(Juliano, Tmax, Tmin, Tmean, v, RH, Rs, elevation, Latitud) As Double
Dim a, cs, P, Es, Ea, Ra, Rnl, RT, Rn, AT, Rso, Rns As Double
'Tmean = (Tmax + Tmin) / 2
If Juliano > 366 Or Juliano <= 0 Or elevation < 0 Then
    EToPM = 0 / 0
ElseIf Tmax < Tmin Or RH > 100 Then
    EToPM = 0 / 0
Else
    a = SlopeSaturationVP(Tmean)
    P = AtmosphericP(elevation)
    cs = PsychrometricC(Tmean, P)
    Es = SaturationVaporP(Tmax, Tmin)
    Ea = ActualVaporP(RH, Tmean)
    Ra = RadiacionExtraterrestres(Juliano, Latitud)
    Rso = ClearSkySR(elevation, Ra)
    Rns = NetShortwaveR(0.23, Rs)
    Rnl = Rnlarga(Tmax, Tmin, Ea, Rs, Rso)
    If Rnl <= 0 Then
        Rnl = 0
    Else
    End If
    
    Rn = Rns - Rnl
    RT = RadiationTerm(a, cs, v, Rn)
    AT = AdvectionTerm(a, cs, v, Tmean, Es, Ea)
    EToPM = RT + AT
End If
End Function
Public Function Windspeed(Velocidad, altura As Double)
    Windspeed = Velocidad * 4.87 / Math.Log(67.8 * altura - 5.42)
End Function
Function aDiaJulianoo(Fecha As Long) As Integer
' verificar cuantos dias julianos tiene el año
Dim año As Integer
Dim dia As Integer
Dim DiaJ As String

año = Year(Fecha)
dia = DateDiff("d", DateSerial(año, 1, 0), Fecha)

If Fecha = DateSerial(año, 2, 29) Then
    dia = 59
        DiaJ = Format(dia, "000")
    aDiaJulianoo = DiaJ

End If

If Fecha = DateSerial(año, 3, 1) Then

    dia = 60
    DiaJ = Format(dia, "000")
    aDiaJulianoo = DiaJ

Else
    If dia > 60 Then
        dia = dia
        DiaJ = Format(dia, "000")
        aDiaJulianoo = DiaJ
    End If
End If
        DiaJ = Format(dia, "000")
        aDiaJulianoo = DiaJ
End Function



