Attribute VB_Name = "modComplex"
'==========================
' Modul: modComplex
'==========================
Option Explicit

' Vytvorenie komplexného èísla
Public Function CCreate(Re As Double, Im As Double) As Complex
    CCreate.Re = Re
    CCreate.Im = Im
End Function

' Vytvorenie z polárneho tvaru (uhol v stupòoch)
Public Function CFromPolar(mag As Double, angDeg As Double) As Complex
    Dim angRad As Double
    angRad = angDeg * DEG2RAD
    CFromPolar.Re = mag * Cos(angRad)
    CFromPolar.Im = mag * Sin(angRad)
End Function

' Konverzia do polárneho tvaru (uhol v stupòoch)
' UDT (Complex) NESMIE by ByVal v Public procedúrach
Public Sub CToPolar(Z As Complex, ByRef mag As Double, ByRef angDeg As Double)
    mag = Sqr(Z.Re * Z.Re + Z.Im * Z.Im)
    If mag = 0 Then
        angDeg = 0
    Else
        angDeg = Atn2(Z.Im, Z.Re) * RAD2DEG
    End If
End Sub

' Pomocná funkcia pre atan2 (VBA nemá natívne)
Private Function Atn2(Y As Double, X As Double) As Double
    If X > 0 Then
        Atn2 = Atn(Y / X)
    ElseIf X < 0 And Y >= 0 Then
        Atn2 = Atn(Y / X) + PI
    ElseIf X < 0 And Y < 0 Then
        Atn2 = Atn(Y / X) - PI
    ElseIf X = 0 And Y > 0 Then
        Atn2 = PI / 2#
    ElseIf X = 0 And Y < 0 Then
        Atn2 = -PI / 2#
    Else
        Atn2 = 0
    End If
End Function

' Sèítanie
Public Function CAdd(a As Complex, B As Complex) As Complex
    CAdd.Re = a.Re + B.Re
    CAdd.Im = a.Im + B.Im
End Function

' Odèítanie
Public Function CSub(a As Complex, B As Complex) As Complex
    CSub.Re = a.Re - B.Re
    CSub.Im = a.Im - B.Im
End Function

' Násobenie
Public Function CMul(a As Complex, B As Complex) As Complex
    CMul.Re = a.Re * B.Re - a.Im * B.Im
    CMul.Im = a.Re * B.Im + a.Im * B.Re
End Function

' Delenie
Public Function CDiv(a As Complex, B As Complex) As Complex
    Dim denom As Double
    denom = B.Re * B.Re + B.Im * B.Im
    If denom = 0 Then
        ' delenie nulou – bezpeèný fallback
        CDiv.Re = 0#
        CDiv.Im = 0#
    Else
        CDiv.Re = (a.Re * B.Re + a.Im * B.Im) / denom
        CDiv.Im = (a.Im * B.Re - a.Re * B.Im) / denom
    End If
End Function

' Umocnenie na celé èíslo n
Public Function CPow(a As Complex, n As Long) As Complex
    Dim i As Long
    Dim res As Complex
    Dim base As Complex   ' lokálna kópia, aby sme nemenili argument

    res = CCreate(1#, 0#)
    base = a

    If n < 0 Then
        base = CDiv(CCreate(1#, 0#), base)
        n = -n
    End If

    For i = 1 To n
        res = CMul(res, base)
    Next i

    CPow = res
End Function

' Absolútna hodnota
Public Function CAbs(Z As Complex) As Double
    CAbs = Sqr(Z.Re * Z.Re + Z.Im * Z.Im)
End Function

' Argument v stupòoch
Public Function CArgDeg(Z As Complex) As Double
    CArgDeg = Atn2(Z.Im, Z.Re) * RAD2DEG
End Function

' Komplexne združené
Public Function CConj(Z As Complex) As Complex
    CConj.Re = Z.Re
    CConj.Im = -Z.Im
End Function

