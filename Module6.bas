Attribute VB_Name = "Module6"
'==========================
' Modul: modNR
' Newton-Raphson Load Flow
'==========================
Option Explicit

'--------------------------------------
' Výpoèet èinných a jalových výkonov P, Q
'--------------------------------------
Private Sub CalcPower(ByVal nBuses As Long, _
                      ByRef G() As Double, _
                      ByRef B() As Double, _
                      ByRef Vmag() As Double, _
                      ByRef Vang() As Double, _
                      ByRef Pcalc() As Double, _
                      ByRef Qcalc() As Double)

    Dim i As Long, k As Long
    Dim theta As Double
    
    For i = 1 To nBuses
        Pcalc(i) = 0#
        Qcalc(i) = 0#
        For k = 1 To nBuses
            theta = Vang(i) - Vang(k)
            Pcalc(i) = Pcalc(i) + Vmag(i) * Vmag(k) * (G(i, k) * Cos(theta) + B(i, k) * Sin(theta))
            Qcalc(i) = Qcalc(i) + Vmag(i) * Vmag(k) * (G(i, k) * Sin(theta) - B(i, k) * Cos(theta))
        Next k
    Next i
End Sub

'--------------------------------------
' Zostavenie vektora nesúladu ?P, ?Q pre PQ uzly
' ?P = Pspec - Pcalc
' ?Q = Qspec - Qcalc
'--------------------------------------
Private Sub BuildMismatchVectors(ByVal nBuses As Long, _
                                 ByRef BusTypes() As BusType, _
                                 ByRef Pspec() As Double, _
                                 ByRef Qspec() As Double, _
                                 ByRef Pcalc() As Double, _
                                 ByRef Qcalc() As Double, _
                                 ByRef PQIndex() As Long, _
                                 ByVal nPQ As Long, _
                                 ByRef mismatch() As Double, _
                                 ByRef maxDP As Double, _
                                 ByRef maxDQ As Double, _
                                 ByRef epsilon As Double)

    Dim i As Long, idx As Long
    Dim dP As Double, dQ As Double
    
    maxDP = 0#
    maxDQ = 0#
    
    ' prvá polovica vektora – ?P
    For i = 1 To nPQ
        idx = PQIndex(i)
        dP = Pspec(idx) - Pcalc(idx)
        mismatch(i) = dP
        If Abs(dP) > maxDP Then maxDP = Abs(dP)
    Next i
    
    ' druhá polovica – ?Q
    For i = 1 To nPQ
        idx = PQIndex(i)
        dQ = Qspec(idx) - Qcalc(idx)
        mismatch(nPQ + i) = dQ
        If Abs(dQ) > maxDQ Then maxDQ = Abs(dQ)
    Next i
    
    epsilon = IIf(maxDP > maxDQ, maxDP, maxDQ)
End Sub

'--------------------------------------
' Zostavenie Jakobiho matice pre PQ uzly
' J má rozmery (2*nPQ) x (2*nPQ) a skladá sa z blokov H, N, M, L:
' [?P]   [ H  N ] [??]
' [?Q] = [ M  L ] [?|V|]
'--------------------------------------
Private Sub BuildJacobian(ByVal nBuses As Long, _
                          ByRef G() As Double, _
                          ByRef B() As Double, _
                          ByRef Vmag() As Double, _
                          ByRef Vang() As Double, _
                          ByRef Pcalc() As Double, _
                          ByRef Qcalc() As Double, _
                          ByRef PQIndex() As Long, _
                          ByVal nPQ As Long, _
                          ByRef J() As Double)

    Dim rowPQ As Long, colPQ As Long   ' indexy v rámci množiny PQ uzlov
    Dim i As Long, k As Long           ' indexy uzlov v celej sieti
    Dim theta As Double
    Dim H As Double, n As Double, M As Double, L As Double
    Dim Vi As Double
    
    For rowPQ = 1 To nPQ
        i = PQIndex(rowPQ)
        For colPQ = 1 To nPQ
            k = PQIndex(colPQ)
            
            If i = k Then
                ' diagonálne prvky (i = k)
                Vi = Vmag(i)
                If Abs(Vi) < 0.000000001 Then Vi = 0.000000001 ' ochrana proti deleniu nulou
                
                ' vzahy v polárnych súradniciach:
                ' H_ii = -Q_i - B_ii * V_i^2
                H = -Qcalc(i) - B(i, i) * Vi * Vi
                ' N_ii = P_i / V_i + G_ii * V_i
                n = Pcalc(i) / Vi + G(i, i) * Vi
                ' M_ii = P_i - G_ii * V_i^2
                M = Pcalc(i) - G(i, i) * Vi * Vi
                ' L_ii = Q_i / V_i - B_ii * V_i
                L = Qcalc(i) / Vi - B(i, i) * Vi
            Else
                ' mimo diagonály (i ? k)
                theta = Vang(i) - Vang(k)
                ' H_ik = V_i V_k (G_ik sin?ik - B_ik cos?ik)
                H = Vmag(i) * Vmag(k) * (G(i, k) * Sin(theta) - B(i, k) * Cos(theta))
                ' N_ik = V_i (G_ik cos?ik + B_ik sin?ik)
                n = Vmag(i) * (G(i, k) * Cos(theta) + B(i, k) * Sin(theta))
                ' M_ik = -V_i V_k (G_ik cos?ik + B_ik sin?ik)
                M = -Vmag(i) * Vmag(k) * (G(i, k) * Cos(theta) + B(i, k) * Sin(theta))
                ' L_ik = V_i (G_ik sin?ik - B_ik cos?ik)
                L = Vmag(i) * (G(i, k) * Sin(theta) - B(i, k) * Cos(theta))
            End If
            
            ' zápis do globálnej Jakobiho matice
            ' poradie neznámych: [??_PQ(1..nPQ), ?|V|_PQ(1..nPQ)]
            J(rowPQ, colPQ) = H                            ' H blok
            J(rowPQ, nPQ + colPQ) = n                      ' N blok
            J(nPQ + rowPQ, colPQ) = M                      ' M blok
            J(nPQ + rowPQ, nPQ + colPQ) = L                ' L blok
        Next colPQ
    Next rowPQ
End Sub

'--------------------------------------
' Riešenie lineárneho systému J * x = rhs
' Riešenie J * x = rhs pomocou MINVERSE/MMULT
Private Sub SolveLinearSystem_JInverse(ByRef J() As Double, _
                                       ByRef rhs() As Double, _
                                       ByRef solution() As Double)

    Dim n As Long
    Dim i As Long, col As Long
    Dim Jvar As Variant
    Dim rhsVar As Variant
    Dim Jinv As Variant
    Dim deltaVar As Variant
    
    On Error GoTo ErrHandler
    
    n = UBound(J, 1)
    
    ReDim Jvar(1 To n, 1 To n)
    ReDim rhsVar(1 To n, 1 To 1)
    
    ' kopírovanie Jakobiho matice a pravej strany do Variant 2D polí
    For i = 1 To n
        For col = 1 To n
            Jvar(i, col) = J(i, col)
        Next col
        rhsVar(i, 1) = rhs(i)
    Next i
    
    ' J^-1 a ?x = J^-1 * rhs
    Jinv = Application.WorksheetFunction.MInverse(Jvar)
    deltaVar = Application.WorksheetFunction.MMult(Jinv, rhsVar)
    
    ReDim solution(1 To n)
    For i = 1 To n
        solution(i) = deltaVar(i, 1)
    Next i
    
    Exit Sub

ErrHandler:
    Err.Raise vbObjectError + 10, , "Chyba pri riešení sústavy (MINVERSE/MMULT) – Jakobiho matica môže by singulárna."
End Sub


'--------------------------------------
' Aktualizácia stavu napätí
'--------------------------------------
Private Sub UpdateState(ByRef Vmag() As Double, _
                        ByRef Vang() As Double, _
                        ByRef PQIndex() As Long, _
                        ByVal nPQ As Long, _
                        ByRef deltaX() As Double)

    Dim i As Long, idx As Long
    
    ' prvá polovica vektora – ??
    For i = 1 To nPQ
        idx = PQIndex(i)
        Vang(idx) = Vang(idx) + deltaX(i)
    Next i
    
    ' druhá polovica – ?|V|
    For i = 1 To nPQ
        idx = PQIndex(i)
        Vmag(idx) = Vmag(idx) + deltaX(nPQ + i)
    Next i
End Sub

'--------------------------------------
' Hlavný Newton-Raphson load-flow
'--------------------------------------
Public Sub NewtonRaphsonLoadFlow()
    Dim SBase_MVA As Double, UBase_kV As Double
    Dim nBuses As Long, nBranches As Long
    Dim BusNames() As String
    Dim BusTypes() As BusType
    Dim Vmag() As Double, Vang() As Double
    Dim Pspec() As Double, Qspec() As Double
    Dim FromBus() As Long, ToBus() As Long
    Dim R() As Double, X() As Double
    Dim Y() As Complex
    Dim G() As Double, B() As Double
    
    Dim Pcalc() As Double, Qcalc() As Double
    Dim PQIndex() As Long
    Dim nPQ As Long
    Dim slackIndex As Long
    Dim i As Long, k As Long
    
    Dim maxIter As Long
    Dim epsLimit As Double
    Dim iter As Long, iterUsed As Long
    Dim mismatch() As Double
    Dim J() As Double
    Dim deltaX() As Double
    Dim maxDP As Double, maxDQ As Double, eps As Double
    Dim startTime As Double, totalTime As Double
    Dim converged As Boolean
    
    On Error GoTo ErrHandler
    
    '--------------------------
    ' Naèítanie parametrov z listu index
    ' B3 – max. poèet iterácií
    ' B4 – epsilon limit
    '--------------------------
    With ThisWorkbook.Worksheets("index")
        maxIter = CLng(ParseDouble(.Range("B3").Value))
        If maxIter <= 0 Then maxIter = 20
        
        epsLimit = ParseDouble(.Range("B4").Value)
        If epsLimit <= 0 Then epsLimit = 0.000001
    End With
    
    '--------------------------
    ' Naèítanie uzlov a vedení
    '--------------------------
    ' naèítanie bázových hodnôt
    Call GetBaseValues(SBase_MVA, UBase_kV)
    
    ' naèítanie uzlov (skutoèné -> p.u.)
    Call LoadBusData(nBuses, BusNames, BusTypes, Vmag, Vang, Pspec, Qspec, SBase_MVA, UBase_kV)
    
    ' naèítanie vedení (ohm -> p.u.)
    Call LoadBranchData(nBranches, FromBus, ToBus, R, X, BusNames, SBase_MVA, UBase_kV)
    
    ' tvorba Y-matice v p.u.
    Call BuildYBus(nBuses, nBranches, FromBus, ToBus, R, X, BusNames, Y, G, B)
    
    '--------------------------
    ' Identifikácia slack a PQ uzlov
    '--------------------------
    slackIndex = 0
    nPQ = 0
    For i = 1 To nBuses
        If BusTypes(i) = btSlack Then
            If slackIndex <> 0 Then
                Err.Raise vbObjectError + 11, , "Viac ako jeden slack uzol."
            End If
            slackIndex = i
        ElseIf BusTypes(i) = btPQ Then
            nPQ = nPQ + 1
        End If
    Next i
    
    If slackIndex = 0 Then
        Err.Raise vbObjectError + 12, , "Nenájdený slack uzol."
    End If
    If nPQ = 0 Then
        Err.Raise vbObjectError + 13, , "Nie sú žiadne PQ uzly."
    End If
    
    ReDim PQIndex(1 To nPQ)
    k = 0
    For i = 1 To nBuses
        If BusTypes(i) = btPQ Then
            k = k + 1
            PQIndex(k) = i
        End If
    Next i
    
    ReDim Pcalc(1 To nBuses)
    ReDim Qcalc(1 To nBuses)
    ReDim mismatch(1 To 2 * nPQ)
    ReDim J(1 To 2 * nPQ, 1 To 2 * nPQ)
    
    '--------------------------
    ' Príprava výsledkových listov
    '--------------------------
    Call ClearResultsSheets
    
    startTime = Timer
    converged = False
    iterUsed = 0
    
    '--------------------------
    ' Hlavný NR iteraèný cyklus
    '--------------------------
    For iter = 1 To maxIter
        iterUsed = iter
        
        ' výpoèet P, Q
        Call CalcPower(nBuses, G, B, Vmag, Vang, Pcalc, Qcalc)
        
        ' vektor nesúladu a epsilon
        Call BuildMismatchVectors(nBuses, BusTypes, Pspec, Qspec, Pcalc, Qcalc, PQIndex, nPQ, mismatch, maxDP, maxDQ, eps)
        
        ' logovanie napätí a epsilon
        Call LogVoltages(iter, BusNames, Vmag, Vang)
        Call LogEpsilon(iter, maxDP, maxDQ, eps)
        
        ' kontrola konvergencie
        If eps < epsLimit Then
            converged = True
            Exit For
        End If
        
        ' Jakobiho matica
        Call BuildJacobian(nBuses, G, B, Vmag, Vang, Pcalc, Qcalc, PQIndex, nPQ, J)
        
        ' riešenie J * ?x = mismatch
        Call SolveLinearSystem_JInverse(J, mismatch, deltaX)
        
        ' aktualizácia napätí
        Call UpdateState(Vmag, Vang, PQIndex, nPQ, deltaX)
    Next iter
    
       totalTime = Timer - startTime
    Call WriteSummaryToIndex(totalTime, iterUsed, eps, converged)
    
    ' zapíš výsledné napätia na list "uzly" v skutoèných hodnotách [kV]
    Call WriteFinalVoltagesToUzly(Vmag, Vang, UBase_kV)
    
    ' vypoèítaj a zapíš prúdy vo vedeniach v reálnych hodnotách
    Call WriteBranchCurrents(nBranches, FromBus, ToBus, R, X, Vmag, Vang, SBase_MVA, UBase_kV)

    Exit Sub

ErrHandler:
    MsgBox "Chyba v Newton-Raphson výpoète: " & Err.Description, vbCritical
End Sub
