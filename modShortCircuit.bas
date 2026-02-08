Attribute VB_Name = "modShortCircuit"
'==========================
' Modul: modShortCircuit
'==========================
Option Explicit

' V˝poËet Ik3 (argumenty poæa s˙ Variant pre stabilitu)
Public Sub CalculateShortCircuit( _
    ByVal nBuses As Long, ByVal nBranches As Long, ByRef FromBus As Variant, ByRef ToBus As Variant, _
    ByRef R As Variant, ByRef X As Variant, ByRef BranchStatus As Variant, _
    ByVal nTrafo As Long, ByRef TrFrom As Variant, ByRef TrTo As Variant, _
    ByRef TrR As Variant, ByRef TrX As Variant, ByRef TrRatio As Variant, _
    ByVal nReaktory As Long, ByRef ReaktorFrom As Variant, ByRef ReaktorTo As Variant, _
    ByRef ReaktorR As Variant, ByRef ReaktorX As Variant, _
    ByVal nDifReaktory As Long, ByRef DifReaktorFrom As Variant, ByRef DifReaktorTo As Variant, _
    ByRef DifReaktorR As Variant, ByRef DifReaktorX As Variant, _
    ByVal nMotors As Long, ByRef MotorBus As Variant, ByRef MotorXk As Variant, ByRef MotorStatus As Variant, _
    ByRef BusNames As Variant, ByRef BusTypes As Variant, ByRef BusBaseKV As Variant, _
    ByRef Ik_input As Variant, ByRef Ik_result As Variant, ByVal SBase_MVA As Double, _
    ByRef IsBusIsolated As Variant, ByRef IsBranchIsolated As Variant, ByRef IsTrafoIsolated As Variant, ByRef IsReaktorIsolated As Variant, ByRef IsDifReaktorIsolated As Variant)

    Dim i As Long, J As Long, k As Long
    Dim Ysc() As Complex, MatBig() As Double, InvBig As Variant
    Dim Z As Complex, Ys As Complex, t1 As Complex, t2 As Complex, a As Double
    Dim slackIdx As Long, Ik_slack As Double, Z_grid_abs As Double, Un As Double
    Dim R_th As Double, X_th As Double, Z_th As Double
    
    ReDim Ysc(1 To nBuses, 1 To nBuses), Ik_result(1 To nBuses)
    
    ' Inicializ·cia Ysc (izolovanÈ = 1.0 na diagon·le)
    For i = 1 To nBuses
        For J = 1 To nBuses: Ysc(i, J) = CCreate(0, 0): Next J
        If IsBusIsolated(i) Then Ysc(i, i) = CCreate(1, 0)
    Next i
    
    ' Vedenia
    For k = 1 To nBranches
        If BranchStatus(k) > 0 Then
            If Not IsBranchIsolated(k) And Not (R(k) = 0 And X(k) = 0) Then
                Z = CCreate(CDbl(R(k)), CDbl(X(k))): Ys = CDiv(CCreate(1, 0), Z)
                i = FromBus(k): J = ToBus(k)
                Ysc(i, i) = CAdd(Ysc(i, i), Ys): Ysc(J, J) = CAdd(Ysc(J, J), Ys)
                Ysc(i, J) = CSub(Ysc(i, J), Ys): Ysc(J, i) = CSub(Ysc(J, i), Ys)
            End If
        End If
    Next k
    
    ' Traf·
    For k = 1 To nTrafo
        If Not IsTrafoIsolated(k) And Not (TrR(k) = 0 And TrX(k) = 0) Then
            Z = CCreate(CDbl(TrR(k)), CDbl(TrX(k))): Ys = CDiv(CCreate(1, 0), Z)
            i = TrFrom(k): J = TrTo(k): a = TrRatio(k)
            t1 = CCreate(Ys.Re / (a * a), Ys.Im / (a * a))
            Ysc(i, i) = CAdd(Ysc(i, i), t1): Ysc(J, J) = CAdd(Ysc(J, J), Ys)
            t2 = CCreate(Ys.Re / a, Ys.Im / a)
            Ysc(i, J) = CSub(Ysc(i, J), t2): Ysc(J, i) = CSub(Ysc(J, i), t2)
        End If
    Next k
    
    ' Reaktory
    For k = 1 To nReaktory
        If Not IsReaktorIsolated(k) And Not (ReaktorR(k) = 0 And ReaktorX(k) = 0) Then
            Z = CCreate(CDbl(ReaktorR(k)), CDbl(ReaktorX(k))): Ys = CDiv(CCreate(1, 0), Z)
            i = ReaktorFrom(k): J = ReaktorTo(k)
            Ysc(i, i) = CAdd(Ysc(i, i), Ys): Ysc(J, J) = CAdd(Ysc(J, J), Ys)
            Ysc(i, J) = CSub(Ysc(i, J), Ys): Ysc(J, i) = CSub(Ysc(J, i), Ys)
        End If
    Next k
    
    ' Dif. Reaktory
    For k = 1 To nDifReaktory
        If Not IsDifReaktorIsolated(k) And Not (DifReaktorR(k) = 0 And DifReaktorX(k) = 0) Then
            Z = CCreate(CDbl(DifReaktorR(k)), CDbl(DifReaktorX(k))): Ys = CDiv(CCreate(1, 0), Z)
            i = DifReaktorFrom(k): J = DifReaktorTo(k)
            Ysc(i, i) = CAdd(Ysc(i, i), Ys): Ysc(J, J) = CAdd(Ysc(J, J), Ys)
            Ysc(i, J) = CSub(Ysc(i, J), Ys): Ysc(J, i) = CSub(Ysc(J, i), Ys)
        End If
    Next k
    
    ' Motory VN - prÌspevok do skratu
    ' Model: impedancia voËi zemi Z = j*Xk (R sa zanedb·va alebo je v Xk zahrnutÈ ako impedancia)
    ' Y = 1 / (j*Xk) = -j / Xk
    For k = 1 To nMotors
        If MotorStatus(k) = 1 Then
            ' Kontrola na nenulov˙ reaktanciu
            If Abs(CDbl(MotorXk(k))) > 0.0000001 Then
                ' Ys = 1 / (j * Xk) = -j * (1/Xk)
                Ys = CCreate(0, -1# / CDbl(MotorXk(k)))
                i = MotorBus(k)
                ' Pridanie k diagon·le
                Ysc(i, i) = CAdd(Ysc(i, i), Ys)
            End If
        End If
    Next k
    
    ' Slack impedancia
    For i = 1 To nBuses
        If BusTypes(i) = 0 Then ' btSlack
            Ik_slack = Ik_input(i)
            If Ik_slack > 0 Then
                Un = BusBaseKV(i)
                Z_grid_abs = (1.1 * SBase_MVA) / (Sqr(3) * Un * Ik_slack)
                Ysc(i, i) = CAdd(Ysc(i, i), CDiv(CCreate(1, 0), CCreate(0, Z_grid_abs)))
            End If
            slackIdx = i: Exit For
        End If
    Next i
    
    ' Debug v˝pis skratovej matice
    Call WriteSCMatrix(Ysc, BusNames)
    
    ' Inverzia (Double matica pre Excel MInverse)
    ReDim MatBig(1 To 2 * nBuses, 1 To 2 * nBuses)
    For i = 1 To nBuses
        For J = 1 To nBuses
            MatBig(i, J) = Ysc(i, J).Re: MatBig(nBuses + i, nBuses + J) = Ysc(i, J).Re
            MatBig(i, nBuses + J) = -Ysc(i, J).Im: MatBig(nBuses + i, J) = Ysc(i, J).Im
        Next J
    Next i
    
    On Error Resume Next
    InvBig = Application.WorksheetFunction.MInverse(MatBig)
    On Error GoTo 0
    If Not IsArray(InvBig) Then Err.Raise vbObjectError + 101, , "Matica je singul·rna (skontrolujte izolovanÈ Ëasti)."
    
    ' V˝poËet Ik
    For i = 1 To nBuses
        If IsBusIsolated(i) Then
            Ik_result(i) = 0
        Else
            R_th = InvBig(i, i): X_th = InvBig(nBuses + i, i)
            Z_th = Sqr(R_th * R_th + X_th * X_th)
            Un = BusBaseKV(i)
            If Z_th > 0.0000001 Then
                Ik_result(i) = (1.1 / Z_th) * (SBase_MVA / (Sqr(3) * Un))
            Else
                Ik_result(i) = 0
            End If
        End If
    Next i
End Sub

Public Sub WriteShortCircuitResults(ByRef Ik_result As Variant, ByVal nBuses As Long)
    Dim ws As Worksheet, i As Long
    Set ws = ThisWorkbook.Worksheets("uzly")
    ws.Cells(2, 10).Value = "Ik3'' [kA]"
    For i = 1 To nBuses
        ws.Cells(2 + i, 10).Value = Round(Ik_result(i), 2)
    Next i
End Sub

' Z·pis skratovej admitanËnej matice pre kontrolu
Private Sub WriteSCMatrix(ByRef Ysc() As Complex, ByRef BusNames As Variant)
    Dim ws As Worksheet
    Dim n As Long
    Dim i As Long, J As Long
    Dim row0 As Long, col0 As Long
    Dim txt As String
    
    Set ws = GetOrCreateSheet("SC_matica")
    ws.Cells.Clear
    
    n = UBound(Ysc, 1)
    
    row0 = 1
    col0 = 1
    
    ' hlaviËky stÂpcov
    For J = 1 To n
        ws.Cells(row0, col0 + J).Value = BusNames(J)
    Next J
    
    ' hlaviËky riadkov + samotn· matica v textovom tvare
    For i = 1 To n
        ws.Cells(row0 + i, col0).Value = BusNames(i)
        For J = 1 To n
            txt = Format(Ysc(i, J).Re, "0.000000") & IIf(Ysc(i, J).Im >= 0, "+j" & Format(Ysc(i, J).Im, "0.000000"), "-j" & Format(Abs(Ysc(i, J).Im), "0.000000"))
            ws.Cells(row0 + i, col0 + J).Value = txt
        Next J
    Next i
    
    ' Re·lna Ëasù
    Dim startRowR As Long
    startRowR = row0 + n + 2
    ws.Cells(startRowR, col0).Value = "Re(Ysc)"
    For J = 1 To n
        ws.Cells(startRowR + 1, col0 + J).Value = BusNames(J)
    Next J
    For i = 1 To n
        ws.Cells(startRowR + 1 + i, col0).Value = BusNames(i)
        For J = 1 To n
            ws.Cells(startRowR + 1 + i, col0 + J).Value = Ysc(i, J).Re
        Next J
    Next i
    
    ' Imagin·rna Ëasù
    Dim startRowX As Long
    startRowX = startRowR + n + 4
    ws.Cells(startRowX, col0).Value = "Im(Ysc)"
    For J = 1 To n
        ws.Cells(startRowX + 1, col0 + J).Value = BusNames(J)
    Next J
    For i = 1 To n
        ws.Cells(startRowX + 1 + i, col0).Value = BusNames(i)
        For J = 1 To n
            ws.Cells(startRowX + 1 + i, col0 + J).Value = Ysc(i, J).Im
        Next J
    Next i
End Sub


