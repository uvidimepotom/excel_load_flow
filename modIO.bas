Attribute VB_Name = "modIO"
'==========================
' Modul: modIO
'==========================
' 25.01.2026 21:30:03 (CET)
Option Explicit

' NaËÌtanie d·t uzlov z listu "uzly"
' SkutoËnÈ hodnoty -> prepoËet do pomern˝ch (p.u.)
'
' Form·t listu "uzly":
'   riadok 2: hlaviËka
'   riadky 3.. : d·ta
'   B: N·zov uzla
'   C: Typ (Slack / PQ)
'   D: |V| [kV]
'   E: ? [deg]
'   F: P [MW]
'   G: Q [Mvar]
'
' SBase_MVA, UBase_kV s˙ b·zovÈ hodnoty zo "index"
' UBase_VN, UBase_NN - novÈ b·zy pre rozlÌöenie hladÌn
Public Sub LoadBusData( _
    ByRef nBuses As Long, _
    ByRef BusNames() As String, _
    ByRef BusTypes() As BusType, _
    ByRef Vmag() As Double, _
    ByRef Vang() As Double, _
    ByRef Pspec() As Double, _
    ByRef Qspec() As Double, _
    ByRef BusBaseKV() As Double, _
    ByVal SBase_MVA As Double, _
    ByVal UBase_VN As Double, _
    ByVal UBase_NN As Double)

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim t As String
    Dim V_kV As Double, P_MW As Double, Q_mvar As Double
    Dim busBase As Double
    
    Set ws = ThisWorkbook.Worksheets("uzly")
    
    ' posledn˝ riadok podæa stÂpca B (N·zov uzla)
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If lastRow < 3 Then
        Err.Raise vbObjectError + 1, , "V liste 'uzly' nie s˙ ûiadne uzly (oËak·vam d·ta od riadku 3)."
    End If
    
    ' prvÈ d·ta s˙ v riadku 3 => poËet uzlov
    nBuses = lastRow - 2
    
    ReDim BusNames(1 To nBuses)
    ReDim BusTypes(1 To nBuses)
    ReDim Vmag(1 To nBuses)
    ReDim Vang(1 To nBuses)
    ReDim Pspec(1 To nBuses)
    ReDim Qspec(1 To nBuses)
    ReDim BusBaseKV(1 To nBuses)
    
    For i = 1 To nBuses
        ' riadok s d·tami = 2 + i (3,4,...)
        BusNames(i) = CStr(ws.Cells(2 + i, 2).Value)   ' B: N·zov uzla
        
        t = CStr(ws.Cells(2 + i, 3).Value)             ' C: Typ
        Select Case UCase$(Trim$(t))
            Case "SLACK"
                BusTypes(i) = btSlack
            Case "PQ"
                BusTypes(i) = btPQ
            Case "PV"
                BusTypes(i) = btPV
            Case Else
                BusTypes(i) = btPQ
        End Select
        
        ' naËÌtanie skutoËn˝ch hodnÙt
        V_kV = ParseDouble(ws.Cells(2 + i, 4).Value)   ' D: |V| [kV]
        P_MW = ParseDouble(ws.Cells(2 + i, 6).Value)   ' F: P [MW]
        Q_mvar = ParseDouble(ws.Cells(2 + i, 7).Value) ' G: Q [Mvar]
        
        ' UrËenie b·zy pre uzol
        busBase = GetBaseVoltageForBus(V_kV, UBase_VN, UBase_NN)
        BusBaseKV(i) = busBase
        
        ' prepoËet do p.u.
        If busBase <> 0# Then
            Vmag(i) = V_kV / busBase
        Else
            Vmag(i) = 1#
        End If
        
        ' PÙvodne naËÌtanie uhla zo stÂpca E:
        ' Vang(i) = ParseDouble(ws.Cells(2 + i, 5).Value) * DEG2RAD
        ' Zmena (27.12.2025): Uhol pre prv˝ krok je vûdy 0
        Vang(i) = 0#
        
        If SBase_MVA <> 0# Then
            Pspec(i) = P_MW / SBase_MVA
            Qspec(i) = Q_mvar / SBase_MVA
        Else
            Pspec(i) = 0#
            Qspec(i) = 0#
        End If
    Next i
End Sub

' NaËÌtanie d·t transform·torov z listu "transformatory"
' C: Uzol od, D: Uzol do
' Detekcia stÂpcov pre Rk, Xk, G0, B0, Prevod
' Ak je prÌtomn˝ stÂpec Zk (napr. na pozÌcii 18), posun +1.
Public Sub LoadTransformerData( _
    ByRef nTrafo As Long, _
    ByRef TrFrom() As Long, _
    ByRef TrTo() As Long, _
    ByRef TrR() As Double, _
    ByRef TrX() As Double, _
    ByRef TrG() As Double, _
    ByRef TrB() As Double, _
    ByRef TrRatio() As Double, _
    ByRef BusNames() As String, _
    ByRef BusBaseKV() As Double, _
    ByVal SBase_MVA As Double)

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim fromName As String, toName As String
    Dim idxFrom As Long, idxTo As Long
    Dim Zbase_prim As Double
    Dim R_ohm As Double, X_ohm As Double
    Dim G_siemens As Double, B_siemens As Double
    Dim Ubase1 As Double
    Dim colOffset As Long
    
    Set ws = ThisWorkbook.Worksheets("transformatory")
    
    lastRow = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
    If lastRow < 3 Then
        nTrafo = 0
        Exit Sub
    End If
    
    nTrafo = lastRow - 2
    
    ReDim TrFrom(1 To nTrafo)
    ReDim TrTo(1 To nTrafo)
    ReDim TrR(1 To nTrafo)
    ReDim TrX(1 To nTrafo)
    ReDim TrG(1 To nTrafo)
    ReDim TrB(1 To nTrafo)
    ReDim TrRatio(1 To nTrafo)
    
    ' Detekcia posunu stÂpcov:
    If InStr(1, LCase(CStr(ws.Cells(2, 18).Value)), "zk") > 0 Then
        colOffset = 1
    Else
        colOffset = 0
    End If
    
    For i = 1 To nTrafo
        fromName = CStr(ws.Cells(i + 2, 3).Value) ' C: Uzol od
        toName = CStr(ws.Cells(i + 2, 4).Value)   ' D: Uzol do
        
        idxFrom = GetBusIndex(fromName, BusNames)
        If idxFrom = 0 Then
            Err.Raise vbObjectError + 4, , "Uzol '" & fromName & "' v liste 'transformatory', riadok " & (i + 2) & " neexistuje."
        End If
        TrFrom(i) = idxFrom
        
        idxTo = GetBusIndex(toName, BusNames)
        If idxTo = 0 Then
            Err.Raise vbObjectError + 5, , "Uzol '" & toName & "' v liste 'transformatory', riadok " & (i + 2) & " neexistuje."
        End If
        TrTo(i) = idxTo
        
        ' B·za impedancie na strane prim·ru (uzol od)
        Ubase1 = BusBaseKV(idxFrom)
        
        If SBase_MVA <> 0# Then
            Zbase_prim = (Ubase1 * Ubase1) / SBase_MVA
        Else
            Zbase_prim = 1#
        End If
        
        ' R, X v Ohmoch -> p.u.
        R_ohm = ParseDouble(ws.Cells(i + 2, 18 + colOffset).Value)
        X_ohm = ParseDouble(ws.Cells(i + 2, 19 + colOffset).Value)
        
        If Zbase_prim <> 0# Then
            TrR(i) = R_ohm / Zbase_prim
            TrX(i) = X_ohm / Zbase_prim
        Else
            TrR(i) = 0#
            TrX(i) = 0#
        End If
        
        ' G, B v Siemensoch -> p.u.
        G_siemens = ParseDouble(ws.Cells(i + 2, 20 + colOffset).Value)
        B_siemens = ParseDouble(ws.Cells(i + 2, 21 + colOffset).Value)
        
        TrG(i) = G_siemens * Zbase_prim
        TrB(i) = B_siemens * Zbase_prim
        
        ' Prevod a
        TrRatio(i) = ParseDouble(ws.Cells(i + 2, 22 + colOffset).Value)
        If TrRatio(i) <= 0# Then TrRatio(i) = 1#
    Next i
End Sub

' NaËÌtanie d·t reaktorov z listu "reaktory"
' B: OznaËenie, C: Uzol od, D: Uzol do, H: R[ohm], I: X[ohm]
Public Sub LoadReactorData( _
    ByRef nReaktory As Long, _
    ByRef ReaktorName() As String, _
    ByRef ReaktorFrom() As Long, _
    ByRef ReaktorTo() As Long, _
    ByRef ReaktorR() As Double, _
    ByRef ReaktorX() As Double, _
    ByRef BusNames() As String, _
    ByRef BusBaseKV() As Double, _
    ByVal SBase_MVA As Double)

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim fromName As String, toName As String
    Dim idxFrom As Long, idxTo As Long
    Dim Zbase As Double
    Dim R_ohm As Double, X_ohm As Double
    Dim Ubase As Double
    
    Set ws = GetOrCreateSheet("reaktory")
    
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If lastRow < 4 Then
        nReaktory = 0
        Exit Sub
    End If
    
    nReaktory = lastRow - 3
    
    ReDim ReaktorName(1 To nReaktory)
    ReDim ReaktorFrom(1 To nReaktory)
    ReDim ReaktorTo(1 To nReaktory)
    ReDim ReaktorR(1 To nReaktory)
    ReDim ReaktorX(1 To nReaktory)
    
    For i = 1 To nReaktory
        ReaktorName(i) = CStr(ws.Cells(i + 3, 2).Value) ' B
        fromName = CStr(ws.Cells(i + 3, 3).Value)       ' C
        toName = CStr(ws.Cells(i + 3, 4).Value)         ' D
        
        idxFrom = GetBusIndex(fromName, BusNames)
        If idxFrom = 0 Then
            Err.Raise vbObjectError + 6, , "Uzol '" & fromName & "' v liste 'reaktory', riadok " & (i + 3) & " neexistuje."
        End If
        ReaktorFrom(i) = idxFrom
        
        idxTo = GetBusIndex(toName, BusNames)
        If idxTo = 0 Then
            Err.Raise vbObjectError + 7, , "Uzol '" & toName & "' v liste 'reaktory', riadok " & (i + 3) & " neexistuje."
        End If
        ReaktorTo(i) = idxTo
        
        ' ImpedanËn· b·za podæa uzla "od"
        Ubase = BusBaseKV(idxFrom)
        If SBase_MVA <> 0# Then
            Zbase = (Ubase * Ubase) / SBase_MVA
        Else
            Zbase = 1#
        End If
        
        R_ohm = ParseDouble(ws.Cells(i + 3, 8).Value) ' H (8)
        X_ohm = ParseDouble(ws.Cells(i + 3, 9).Value) ' I (9)
        
        If Zbase <> 0# Then
            ReaktorR(i) = R_ohm / Zbase
            ReaktorX(i) = X_ohm / Zbase
        Else
            ReaktorR(i) = 0#
            ReaktorX(i) = 0#
        End If
    Next i
End Sub

' NaËÌtanie d·t dif. reaktorov z listu "dif_reaktory"
' C: Uzol od, D: Uzol do, I: R[ohm], J: X[ohm]
' HlaviËka: riadok 3, d·ta od 4
Public Sub LoadDifReactorData( _
    ByRef nDifReaktory As Long, _
    ByRef DifReaktorName() As String, _
    ByRef DifReaktorFrom() As Long, _
    ByRef DifReaktorTo() As Long, _
    ByRef DifReaktorR() As Double, _
    ByRef DifReaktorX() As Double, _
    ByRef BusNames() As String, _
    ByRef BusBaseKV() As Double, _
    ByVal SBase_MVA As Double)

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim fromName As String, toName As String
    Dim idxFrom As Long, idxTo As Long
    Dim Zbase As Double
    Dim R_ohm As Double, X_ohm As Double
    Dim Ubase As Double
    
    Set ws = GetOrCreateSheet("dif_reaktory")
    
    lastRow = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
    If lastRow < 4 Then
        nDifReaktory = 0
        Exit Sub
    End If
    
    nDifReaktory = lastRow - 3
    
    ReDim DifReaktorName(1 To nDifReaktory)
    ReDim DifReaktorFrom(1 To nDifReaktory)
    ReDim DifReaktorTo(1 To nDifReaktory)
    ReDim DifReaktorR(1 To nDifReaktory)
    ReDim DifReaktorX(1 To nDifReaktory)
    
    For i = 1 To nDifReaktory
        ' Name nie je v öpecifik·cii, ale pole existuje. D·me tam "DR" + index alebo pr·zdny string?
        DifReaktorName(i) = "DR" & i
        fromName = CStr(ws.Cells(i + 3, 3).Value)       ' C
        toName = CStr(ws.Cells(i + 3, 4).Value)         ' D
        
        idxFrom = GetBusIndex(fromName, BusNames)
        If idxFrom = 0 Then
            Err.Raise vbObjectError + 8, , "Uzol '" & fromName & "' v liste 'dif_reaktory', riadok " & (i + 3) & " neexistuje."
        End If
        DifReaktorFrom(i) = idxFrom
        
        idxTo = GetBusIndex(toName, BusNames)
        If idxTo = 0 Then
            Err.Raise vbObjectError + 9, , "Uzol '" & toName & "' v liste 'dif_reaktory', riadok " & (i + 3) & " neexistuje."
        End If
        DifReaktorTo(i) = idxTo
        
        ' ImpedanËn· b·za podæa uzla "od"
        Ubase = BusBaseKV(idxFrom)
        If SBase_MVA <> 0# Then
            Zbase = (Ubase * Ubase) / SBase_MVA
        Else
            Zbase = 1#
        End If
        
        R_ohm = ParseDouble(ws.Cells(i + 3, 9).Value)  ' I (9)
        X_ohm = ParseDouble(ws.Cells(i + 3, 10).Value) ' J (10)
        
        If Zbase <> 0# Then
            DifReaktorR(i) = R_ohm / Zbase
            DifReaktorX(i) = X_ohm / Zbase
        Else
            DifReaktorR(i) = 0#
            DifReaktorX(i) = 0#
        End If
    Next i
End Sub

' NaËÌtanie kompenz·cie z listu "kompenz·cia"
' B: N·zov, C: Uzol, N: Status (1/0), P: X_L[ohm], Q: X_C[ohm]
' HlaviËka: riadok 3, d·ta od 4
Public Sub LoadCompData( _
    ByRef nComp As Long, _
    ByRef CompName() As String, _
    ByRef CompBus() As Long, _
    ByRef CompB() As Double, _
    ByRef CompStatus() As Integer, _
    ByRef BusNames() As String, _
    ByRef BusBaseKV() As Double, _
    ByVal SBase_MVA As Double)

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim busName As String, idxBus As Long
    Dim XL_ohm As Double, XC_ohm As Double
    Dim Zbase As Double, Ubase As Double
    Dim BL_pu As Double, BC_pu As Double
    Dim statusVal As Variant
    
    Set ws = GetOrCreateSheet("kompenz·cia")
    
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If lastRow < 4 Then
        nComp = 0
        Exit Sub
    End If
    
    nComp = lastRow - 3
    
    ReDim CompName(1 To nComp)
    ReDim CompBus(1 To nComp)
    ReDim CompB(1 To nComp)
    ReDim CompStatus(1 To nComp)
    
    For i = 1 To nComp
        CompName(i) = CStr(ws.Cells(i + 3, 2).Value) ' B
        busName = CStr(ws.Cells(i + 3, 3).Value)     ' C
        
        idxBus = GetBusIndex(busName, BusNames)
        If idxBus = 0 Then
            Err.Raise vbObjectError + 10, , "Uzol '" & busName & "' v liste 'kompenz·cia', riadok " & (i + 3) & " neexistuje."
        End If
        CompBus(i) = idxBus
        
        ' Status N (14)
        statusVal = ws.Cells(i + 3, 14).Value
        If IsNumeric(statusVal) Then
            CompStatus(i) = CInt(statusVal)
        Else
            CompStatus(i) = 0
        End If
        
        ' Ak nie je aktÌvna (Status=0), susceptancia je 0
        If CompStatus(i) = 1 Then
            Ubase = BusBaseKV(idxBus)
            If SBase_MVA <> 0# Then
                Zbase = (Ubase * Ubase) / SBase_MVA
            Else
                Zbase = 1#
            End If
            
            XL_ohm = ParseDouble(ws.Cells(i + 3, 16).Value) ' P (16)
            XC_ohm = ParseDouble(ws.Cells(i + 3, 17).Value) ' Q (17)
            
            ' SÈriov· kombin·cia / V˝sledn· reaktancia:
            ' UûÌvateæ poûaduje 1/(XC - XL)
            ' Ak XL=0 a XC>0 -> 1/XC (Kapacitn˝, B > 0)
            ' Ak XL>0 a XC=0 -> 1/-XL (InduktÌvny, B < 0)
            ' Ak s˙ zadanÈ obe, rozhoduje rozdiel.
            
            Dim X_net_ohm As Double
            X_net_ohm = XC_ohm - XL_ohm
            
            If Zbase <> 0# And Abs(X_net_ohm) > 0.000001 Then
                ' B [S] = 1 / X_net [Ohm]
                ' B [p.u.] = B [S] * Zbase
                CompB(i) = (1# / X_net_ohm) * Zbase
            Else
                CompB(i) = 0#
            End If
        Else
            CompB(i) = 0#
        End If
    Next i
End Sub

' NaËÌtanie d·t gener·torov z listu "generatory"
' B: N·zov, C: PQ_Uzol, D: PV_Uzol, E: Status, L: Ra, M: P_gen, O: Xs, P: Xd''
Public Sub LoadGeneratorData( _
    ByRef nGens As Long, _
    ByRef GenName() As String, _
    ByRef GenBusPQ() As Long, _
    ByRef GenBusPV() As Long, _
    ByRef GenStatus() As Integer, _
    ByRef GenRa() As Double, _
    ByRef GenP() As Double, _
    ByRef GenXs() As Double, _
    ByRef GenXd() As Double, _
    ByRef BusNames() As String, _
    ByRef BusBaseKV() As Double, _
    ByVal SBase_MVA As Double)

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim pqName As String, pvName As String
    Dim idxPQ As Long, idxPV As Long
    Dim Ubase As Double, Zbase As Double
    Dim stVal As Variant
    
    Set ws = GetOrCreateSheet("generatory")
    
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If lastRow < 3 Then
        nGens = 0
        Exit Sub
    End If
    
    nGens = lastRow - 2
    
    ReDim GenName(1 To nGens)
    ReDim GenBusPQ(1 To nGens)
    ReDim GenBusPV(1 To nGens)
    ReDim GenStatus(1 To nGens)
    ReDim GenRa(1 To nGens)
    ReDim GenP(1 To nGens)
    ReDim GenXs(1 To nGens)
    ReDim GenXd(1 To nGens)
    
    For i = 1 To nGens
        GenName(i) = CStr(ws.Cells(i + 2, 2).Value) ' B
        pqName = CStr(ws.Cells(i + 2, 3).Value)     ' C
        pvName = CStr(ws.Cells(i + 2, 4).Value)     ' D
        
        ' Status E (5)
        stVal = ws.Cells(i + 2, 5).Value
        If IsNumeric(stVal) Then
            GenStatus(i) = CInt(stVal)
        Else
            GenStatus(i) = 0
        End If
        
        idxPQ = GetBusIndex(pqName, BusNames)
        If idxPQ = 0 Then
            Err.Raise vbObjectError + 12, , "PQ Uzol '" & pqName & "' v liste 'generatory' neexistuje."
        End If
        GenBusPQ(i) = idxPQ
        
        idxPV = GetBusIndex(pvName, BusNames)
        If idxPV = 0 Then
            Err.Raise vbObjectError + 13, , "PV Uzol '" & pvName & "' v liste 'generatory' neexistuje."
        End If
        GenBusPV(i) = idxPV
        
        ' P_gen [MW] -> M (13)
        If SBase_MVA <> 0 Then
            GenP(i) = ParseDouble(ws.Cells(i + 2, 13).Value) / SBase_MVA
        Else
            GenP(i) = 0
        End If
        
        ' Ra [ohm] -> L (12), Xs [ohm] -> O (15), Xd'' [ohm] -> P (16)
        ' B·za impedancie (podæa PQ uzla, lebo tam je pripojen˝)
        Ubase = BusBaseKV(idxPQ)
        If SBase_MVA <> 0 Then
            Zbase = (Ubase * Ubase) / SBase_MVA
        Else
            Zbase = 1
        End If
        
        If Zbase <> 0 Then
            GenRa(i) = ParseDouble(ws.Cells(i + 2, 12).Value) / Zbase
            GenXs(i) = ParseDouble(ws.Cells(i + 2, 15).Value) / Zbase
            GenXd(i) = ParseDouble(ws.Cells(i + 2, 16).Value) / Zbase
        End If
    Next i
End Sub

' Z·pis v˝sledkov gener·torov (Load Flow) - Q do stÂpca N
Public Sub WriteGeneratorResults( _
    ByVal nGens As Long, _
    ByRef GenBusPQ() As Long, _
    ByRef GenBusPV() As Long, _
    ByRef GenXs() As Double, _
    ByRef GenRa() As Double, _
    ByRef GenStatus() As Integer, _
    ByRef Vmag() As Double, _
    ByRef Vang() As Double, _
    ByVal SBase_MVA As Double)

    Dim ws As Worksheet
    Dim i As Long
    Dim idxPQ As Long, idxPV As Long
    Dim Vi As Complex, Vj As Complex
    Dim Zs As Complex, I_pu As Complex, S_pu As Complex
    Dim Q_mvar As Double
    
    Set ws = GetOrCreateSheet("generatory")
    ' HlaviËka
    ws.Cells(2, 14).Value = "Q_gen [Mvar]" ' N (14)
    
    For i = 1 To nGens
        If GenStatus(i) = 1 Then
            idxPQ = GenBusPQ(i)
            idxPV = GenBusPV(i)
            
            ' Tok z PV do PQ? Nie, Q_gen je injekcia do PQ uzla (alebo v˝roba gener·tora).
            ' Q_gen sa meria na svork·ch (PQ uzol).
            ' Pr˙d I = (V_pv - V_pq) / (Ra + jXs)
            ' S_pq = V_pq * conj(I) ... to je prÌkon do PQ uzla z vetvy.
            ' Ak je S_pq kladnÈ, tak to "teËie do uzla".
            
            Vi = CFromPolar(Vmag(idxPV), Vang(idxPV) * RAD2DEG)
            Vj = CFromPolar(Vmag(idxPQ), Vang(idxPQ) * RAD2DEG)
            Zs = CCreate(GenRa(i), GenXs(i))
            
            ' I = (Vi - Vj) / Zs
            I_pu = CDiv(CSub(Vi, Vj), Zs)
            
            ' V˝kon dod·van˝ do siete (na svork·ch j)
            ' S_gen = Vj * conj(I)
            ' (Pozor na znamienko: I teËie z PV do PQ. Takûe I vstupuje do PQ. S = U*I*)
            S_pu = CMul(Vj, CConj(I_pu))
            
            Q_mvar = S_pu.Im * SBase_MVA
            
            ws.Cells(i + 2, 14).Value = Round(Q_mvar, 2)
        Else
            ws.Cells(i + 2, 14).Value = 0
        End If
    Next i
End Sub

' NaËÌtanie motorov z listu "motoryVN"
' B: N·zov, C: Uzol, L: R[ohm], P: Xk[ohm], Q: G[S], R: B[S], S: Status
' PrepoËet Xk na p.u. (Zbase), G a B na p.u.
Public Sub LoadMotorData( _
    ByRef nMotors As Long, _
    ByRef MotorName() As String, _
    ByRef MotorBus() As Long, _
    ByRef MotorR() As Double, _
    ByRef MotorXk() As Double, _
    ByRef MotorG() As Double, _
    ByRef MotorB() As Double, _
    ByRef MotorStatus() As Integer, _
    ByRef BusNames() As String, _
    ByRef BusBaseKV() As Double, _
    ByVal SBase_MVA As Double)

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim busName As String, idxBus As Long
    Dim Ubase As Double, Zbase As Double
    Dim R_ohm As Double, Xk_ohm As Double
    Dim G_siemens As Double, B_siemens As Double
    Dim statusVal As Variant
    
    Set ws = GetOrCreateSheet("motoryVN")
    
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If lastRow < 3 Then
        nMotors = 0
        Exit Sub
    End If
    
    nMotors = lastRow - 2
    
    ReDim MotorName(1 To nMotors)
    ReDim MotorBus(1 To nMotors)
    ReDim MotorR(1 To nMotors)
    ReDim MotorXk(1 To nMotors)
    ReDim MotorG(1 To nMotors)
    ReDim MotorB(1 To nMotors)
    ReDim MotorStatus(1 To nMotors)
    
    For i = 1 To nMotors
        MotorName(i) = CStr(ws.Cells(i + 2, 2).Value) ' B
        busName = CStr(ws.Cells(i + 2, 3).Value)      ' C
        
        idxBus = GetBusIndex(busName, BusNames)
        If idxBus = 0 Then
            Err.Raise vbObjectError + 11, , "Uzol '" & busName & "' v liste 'motoryVN', riadok " & (i + 2) & " neexistuje."
        End If
        MotorBus(i) = idxBus
        
        ' Status S (19)
        statusVal = ws.Cells(i + 2, 19).Value
        If IsNumeric(statusVal) Then
            MotorStatus(i) = CInt(statusVal)
        Else
            MotorStatus(i) = 0
        End If
        
        ' R [ohm] L (12) - iba pre straty, neprepoËÌtavame na p.u. (alebo hej? Vzorec je 3*I^2*R. Ak I je v A a R v Ohm, tak je to OK.)
        ' NaËÌtame ako Ohmy.
        MotorR(i) = ParseDouble(ws.Cells(i + 2, 12).Value)
        
        ' B·za pre Xk, G, B
        Ubase = BusBaseKV(idxBus)
        If SBase_MVA <> 0# Then
            Zbase = (Ubase * Ubase) / SBase_MVA
        Else
            Zbase = 1#
        End If
        
        If MotorStatus(i) = 1 Then
             ' Xk [ohm] P (16)
            Xk_ohm = ParseDouble(ws.Cells(i + 2, 16).Value)
            If Zbase <> 0# Then
                MotorXk(i) = Xk_ohm / Zbase
            Else
                MotorXk(i) = 0#
            End If
            
            ' G [S] Q (17)
            G_siemens = ParseDouble(ws.Cells(i + 2, 17).Value)
            MotorG(i) = G_siemens * Zbase
            
            ' B [S] R (18)
            B_siemens = ParseDouble(ws.Cells(i + 2, 18).Value)
            MotorB(i) = B_siemens * Zbase
        Else
            MotorXk(i) = 0#
            MotorG(i) = 0#
            MotorB(i) = 0#
        End If
    Next i
End Sub

' Z·pis v˝sledkov motorov (Load Flow) - I [A] do AD, Ploss [kW] do AE
Public Sub WriteMotorResults( _
    ByVal nMotors As Long, _
    ByRef MotorBus() As Long, _
    ByRef MotorR() As Double, _
    ByRef MotorG() As Double, _
    ByRef MotorB() As Double, _
    ByRef MotorStatus() As Integer, _
    ByRef Vmag() As Double, _
    ByRef BusBaseKV() As Double, _
    ByVal SBase_MVA As Double)

    Dim ws As Worksheet
    Dim i As Long
    Dim idxBus As Long
    Dim V_pu As Double
    Dim Y_mod_pu As Double
    Dim I_pu As Double, I_abs As Double, Ibase As Double
    Dim Ploss_kW As Double
    
    Set ws = GetOrCreateSheet("motoryVN")
    
    ' HlaviËka
    ws.Cells(2, 30).Value = "I [A]"     ' AD (30)
    ws.Cells(2, 31).Value = "Ploss [kW]" ' AE (31)
    
    For i = 1 To nMotors
        If MotorStatus(i) = 1 Then
            idxBus = MotorBus(i)
            V_pu = Vmag(idxBus)
            
            ' I_pu = V_pu * |Y_motor_pu|
            ' Y_motor = G + jB
            Y_mod_pu = Sqr(MotorG(i) * MotorG(i) + MotorB(i) * MotorB(i))
            I_pu = V_pu * Y_mod_pu
            
            ' Ibase [kA] = Sbase / (sqrt(3)*Ubase) -> *1000 = [A]
            If BusBaseKV(idxBus) <> 0 Then
                Ibase = (SBase_MVA * 1000#) / (Sqr(3) * BusBaseKV(idxBus))
            Else
                Ibase = 0
            End If
            
            I_abs = I_pu * Ibase
            
            ' Ploss = 3 * R * I^2  (I v A, R v Ohm -> W -> /1000 -> kW)
            Ploss_kW = (3# * MotorR(i) * I_abs * I_abs) / 1000#
            
            ws.Cells(i + 2, 30).Value = Round(I_abs, 2)
            ws.Cells(i + 2, 31).Value = Round(Ploss_kW, 2)
        Else
            ws.Cells(i + 2, 30).Value = 0
            ws.Cells(i + 2, 31).Value = 0
        End If
    Next i
End Sub

' Z·pis v˝sledkov kompenz·cie (Load Flow) - Nap‰tie do stÂpca O
Public Sub WriteCompResults( _
    ByVal nComp As Long, _
    ByRef CompBus() As Long, _
    ByRef Vmag() As Double, _
    ByRef BusBaseKV() As Double)

    Dim ws As Worksheet
    Dim i As Long
    Dim U_kV As Double
    Dim idxBus As Long
    
    Set ws = GetOrCreateSheet("kompenz·cia")
    
    ' HlaviËka
    ws.Cells(3, 15).Value = "U [kV]" ' O (15)
    
    For i = 1 To nComp
        idxBus = CompBus(i)
        U_kV = Vmag(idxBus) * BusBaseKV(idxBus)
        ws.Cells(i + 3, 15).Value = Round(U_kV, 2)
    Next i
End Sub

' Z·pis v˝sledkov reaktorov (Load Flow)
' Z: I[A], AA: dU[%], AB: P[MW], AC: Q[MVAr], AD: Ploss[kW]
Public Sub WriteReactorResults( _
    ByVal nReaktory As Long, _
    ByRef ReaktorFrom() As Long, _
    ByRef ReaktorTo() As Long, _
    ByRef ReaktorR() As Double, _
    ByRef ReaktorX() As Double, _
    ByRef Vmag() As Double, _
    ByRef Vang() As Double, _
    ByRef BusBaseKV() As Double, _
    ByVal SBase_MVA As Double)

    Dim ws As Worksheet
    Dim k As Long
    Dim iBus As Long, jBus As Long
    Dim Vi As Complex, Vj As Complex
    Dim Zpu As Complex, Iij_pu As Complex
    Dim Iabs_pu As Double, Iabs_A As Double
    Dim Ibase_A As Double
    Dim dU_percent As Double
    Dim Sij_pu As Complex
    Dim Pij_MW As Double, Qij_MVAr As Double
    Dim Ploss_kW As Double
    Dim Ubase As Double
    
    Set ws = GetOrCreateSheet("reaktory")
    
    ' HlaviËky (riadok 3)
    ws.Cells(3, 26).Value = "I [A]"
    ws.Cells(3, 27).Value = "dU [%]"
    ws.Cells(3, 28).Value = "P [MW]"
    ws.Cells(3, 29).Value = "Q [MVAr]"
    ws.Cells(3, 30).Value = "Ploss [kW]"
    
    For k = 1 To nReaktory
        iBus = ReaktorFrom(k)
        jBus = ReaktorTo(k)
        
        Vi = CFromPolar(Vmag(iBus), Vang(iBus) * RAD2DEG)
        Vj = CFromPolar(Vmag(jBus), Vang(jBus) * RAD2DEG)
        
        Zpu = CCreate(ReaktorR(k), ReaktorX(k))
        
        If Abs(Zpu.Re) < 0.000000001 And Abs(Zpu.Im) < 0.000000001 Then
            Iabs_pu = 0#
            dU_percent = 0#
            Pij_MW = 0#
            Qij_MVAr = 0#
            Ploss_kW = 0#
        Else
            ' I = (Vi - Vj) / Z
            Iij_pu = CDiv(CSub(Vi, Vj), Zpu)
            Iabs_pu = CAbs(Iij_pu)
            
            ' dU = (|Vi| - |Vj|) * 100
            dU_percent = (Vmag(iBus) - Vmag(jBus)) * 100#
            
            ' S = Vi * conj(I)
            Sij_pu = CMul(Vi, CConj(Iij_pu))
            Pij_MW = Sij_pu.Re * SBase_MVA
            Qij_MVAr = Sij_pu.Im * SBase_MVA
            
            ' Ploss = |I|^2 * R * Sbase * 1000
            Ploss_kW = Iabs_pu * Iabs_pu * ReaktorR(k) * SBase_MVA * 1000#
        End If
        
        ' I [A]
        Ubase = BusBaseKV(iBus)
        If SBase_MVA <> 0# And Ubase <> 0# Then
            Ibase_A = (SBase_MVA * 1000#) / (Sqr(3#) * Ubase)
        Else
            Ibase_A = 1#
        End If
        Iabs_A = Iabs_pu * Ibase_A
        
        ws.Cells(k + 3, 26).Value = Round(Iabs_A, 2)
        ws.Cells(k + 3, 27).Value = Round(dU_percent, 2)
        ws.Cells(k + 3, 28).Value = Round(Pij_MW, 2)
        ws.Cells(k + 3, 29).Value = Round(Qij_MVAr, 2)
        ws.Cells(k + 3, 30).Value = Round(Ploss_kW, 2)
    Next k
End Sub

' Z·pis v˝sledkov dif. reaktorov (Load Flow)
' X: I[A], Y: dU[%], Z: P[MW], AA: Q[MVAr], AB: Ploss[kW]
Public Sub WriteDifReactorResults( _
    ByVal nDifReaktory As Long, _
    ByRef DifReaktorFrom() As Long, _
    ByRef DifReaktorTo() As Long, _
    ByRef DifReaktorR() As Double, _
    ByRef DifReaktorX() As Double, _
    ByRef Vmag() As Double, _
    ByRef Vang() As Double, _
    ByRef BusBaseKV() As Double, _
    ByVal SBase_MVA As Double)

    Dim ws As Worksheet
    Dim k As Long
    Dim iBus As Long, jBus As Long
    Dim Vi As Complex, Vj As Complex
    Dim Zpu As Complex, Iij_pu As Complex
    Dim Iabs_pu As Double, Iabs_A As Double
    Dim Ibase_A As Double
    Dim dU_percent As Double
    Dim Sij_pu As Complex
    Dim Pij_MW As Double, Qij_MVAr As Double
    Dim Ploss_kW As Double
    Dim Ubase As Double
    
    Set ws = GetOrCreateSheet("dif_reaktory")
    
    ' HlaviËky (riadok 3) od stÂpca X (24)
    ws.Cells(3, 24).Value = "I [A]"
    ws.Cells(3, 25).Value = "dU [%]"
    ws.Cells(3, 26).Value = "P [MW]"
    ws.Cells(3, 27).Value = "Q [MVAr]"
    ws.Cells(3, 28).Value = "Ploss [kW]"
    
    For k = 1 To nDifReaktory
        iBus = DifReaktorFrom(k)
        jBus = DifReaktorTo(k)
        
        Vi = CFromPolar(Vmag(iBus), Vang(iBus) * RAD2DEG)
        Vj = CFromPolar(Vmag(jBus), Vang(jBus) * RAD2DEG)
        
        Zpu = CCreate(DifReaktorR(k), DifReaktorX(k))
        
        If Abs(Zpu.Re) < 0.000000001 And Abs(Zpu.Im) < 0.000000001 Then
            Iabs_pu = 0#
            dU_percent = 0#
            Pij_MW = 0#
            Qij_MVAr = 0#
            Ploss_kW = 0#
        Else
            Iij_pu = CDiv(CSub(Vi, Vj), Zpu)
            Iabs_pu = CAbs(Iij_pu)
            
            dU_percent = (Vmag(iBus) - Vmag(jBus)) * 100#
            
            Sij_pu = CMul(Vi, CConj(Iij_pu))
            Pij_MW = Sij_pu.Re * SBase_MVA
            Qij_MVAr = Sij_pu.Im * SBase_MVA
            
            Ploss_kW = Iabs_pu * Iabs_pu * DifReaktorR(k) * SBase_MVA * 1000#
        End If
        
        ' I [A]
        Ubase = BusBaseKV(iBus)
        If SBase_MVA <> 0# And Ubase <> 0# Then
            Ibase_A = (SBase_MVA * 1000#) / (Sqr(3#) * Ubase)
        Else
            Ibase_A = 1#
        End If
        Iabs_A = Iabs_pu * Ibase_A
        
        ws.Cells(k + 3, 24).Value = Round(Iabs_A, 2)
        ws.Cells(k + 3, 25).Value = Round(dU_percent, 2)
        ws.Cells(k + 3, 26).Value = Round(Pij_MW, 2)
        ws.Cells(k + 3, 27).Value = Round(Qij_MVAr, 2)
        ws.Cells(k + 3, 28).Value = Round(Ploss_kW, 2)
    Next k
End Sub

' Z·pis admitanËnej matice a pomocn˝ch blokov G a B
Public Sub WriteYMatrix( _
    ByRef Y() As Complex, _
    ByRef G() As Double, _
    ByRef B() As Double, _
    ByRef BusNames() As String)

    Dim ws As Worksheet
    Dim n As Long
    Dim i As Long, J As Long
    Dim row0 As Long, col0 As Long
    Dim txt As String
    
    Set ws = GetOrCreateSheet("Y_matica")
    ws.Cells.Clear
    
    n = UBound(BusNames)
    
    row0 = 1
    col0 = 1
    
    ' hlaviËky stÂpcov
    For J = 1 To n
        ws.Cells(row0, col0 + J).Value = BusNames(J)
    Next J
    
    ' hlaviËky riadkov + samotn· matica G+jB
    For i = 1 To n
        ws.Cells(row0 + i, col0).Value = BusNames(i)
        For J = 1 To n
            txt = Format(G(i, J), "0.000000") & IIf(B(i, J) >= 0, "+j" & Format(B(i, J), "0.000000"), "-j" & Format(Abs(B(i, J)), "0.000000"))
            ws.Cells(row0 + i, col0 + J).Value = txt
        Next J
    Next i
    
    ' pomocn˝ blok G
    Dim startRowG As Long
    startRowG = row0 + n + 2
    ws.Cells(startRowG, col0).Value = "G = Re(Y)"
    For J = 1 To n
        ws.Cells(startRowG + 1, col0 + J).Value = BusNames(J)
    Next J
    For i = 1 To n
        ws.Cells(startRowG + 1 + i, col0).Value = BusNames(i)
        For J = 1 To n
            ws.Cells(startRowG + 1 + i, col0 + J).Value = G(i, J)
        Next J
    Next i
    
    ' pomocn˝ blok B
    Dim startRowB As Long
    startRowB = startRowG + n + 4
    ws.Cells(startRowB, col0).Value = "B = Im(Y)"
    For J = 1 To n
        ws.Cells(startRowB + 1, col0 + J).Value = BusNames(J)
    Next J
    For i = 1 To n
        ws.Cells(startRowB + 1 + i, col0).Value = BusNames(i)
        For J = 1 To n
            ws.Cells(startRowB + 1 + i, col0 + J).Value = B(i, J)
        Next J
    Next i
End Sub

' Vymazanie / prÌprava v˝sledkov˝ch listov "napatia" a "epsilon"
Public Sub ClearResultsSheets()
    Dim wsV As Worksheet, wsE As Worksheet
    
    Set wsV = GetOrCreateSheet("napatia")
    Set wsE = GetOrCreateSheet("epsilon")
    
    wsV.Cells.Clear
    wsE.Cells.Clear
    
    ' hlaviËky
    wsV.Cells(1, 1).Value = "Iter·cia"
    wsV.Cells(1, 2).Value = "Uzol"
    wsV.Cells(1, 3).Value = "|V| [p.u.]"
    wsV.Cells(1, 4).Value = "? [deg]"
    
    wsE.Cells(1, 1).Value = "Iter·cia"
    wsE.Cells(1, 2).Value = "max|?P|"
    wsE.Cells(1, 3).Value = "max|?Q|"
    wsE.Cells(1, 4).Value = "epsilon"
End Sub

' Logovanie nap‰tÌ v jednej iter·cii
Public Sub LogVoltages(ByVal iter As Long, _
                       ByRef BusNames() As String, _
                       ByRef Vmag() As Double, _
                       ByRef Vang() As Double)

    Dim ws As Worksheet
    Dim i As Long
    Dim rowStart As Long
    
    Set ws = GetOrCreateSheet("napatia")
    rowStart = FirstFreeRow(ws, 1)
    
    For i = LBound(BusNames) To UBound(BusNames)
        ws.Cells(rowStart, 1).Value = iter
        ws.Cells(rowStart, 2).Value = BusNames(i)
        ws.Cells(rowStart, 3).Value = Vmag(i)
        ws.Cells(rowStart, 4).Value = Vang(i) * RAD2DEG
        rowStart = rowStart + 1
    Next i
End Sub

' Logovanie epsilon v jednej iter·cii
Public Sub LogEpsilon(ByVal iter As Long, _
                      ByVal maxDP As Double, _
                      ByVal maxDQ As Double, _
                      ByVal eps As Double)

    Dim ws As Worksheet
    Dim R As Long
    
    Set ws = GetOrCreateSheet("epsilon")
    R = FirstFreeRow(ws, 1)
    
    ws.Cells(R, 1).Value = iter
    ws.Cells(R, 2).Value = maxDP
    ws.Cells(R, 3).Value = maxDQ
    ws.Cells(R, 4).Value = eps
End Sub

' Z·pis s˙hrnn˝ch v˝sledkov do listu "index"
Public Sub WriteSummaryToIndex(ByVal totalTime As Double, _
                               ByVal iterCount As Long, _
                               ByVal epsFinal As Double, _
                               ByVal converged As Boolean)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("index")
    
    ws.Range("B6").Value = totalTime
    ws.Range("B7").Value = iterCount
    ws.Range("B8").Value = epsFinal
    ws.Range("B9").Value = IIf(converged, "Konvergovalo", "Nekonvergovalo")
End Sub

' Z·pis v˝sledn˝ch nap‰tÌ z NR v˝poËtu na list "uzly"
' H (stÂpec 8): |V| v˝p. [kV]
' I (stÂpec 9): ? v˝p. [deg]
'
' Vmag(i) je v p.u., prepoËet na kV cez BusBaseKV(i)
Public Sub WriteFinalVoltagesToUzly( _
    ByRef Vmag() As Double, _
    ByRef Vang() As Double, _
    ByRef BusBaseKV() As Double)

    Dim ws As Worksheet
    Dim i As Long
    Dim nBuses As Long
    
    Set ws = ThisWorkbook.Worksheets("uzly")
    
    nBuses = UBound(Vmag)
    
    ' hlaviËky v˝sledkov (riadok 2)
    ws.Cells(2, 8).Value = "|V| v˝p. [kV]"
    ws.Cells(2, 9).Value = "? v˝p. [deg]"
    
    ' d·ta od riadku 3
    For i = 1 To nBuses
        ws.Cells(2 + i, 8).Value = Round(Vmag(i) * BusBaseKV(i), 2)
        ws.Cells(2 + i, 9).Value = Round(Vang(i) * RAD2DEG, 2)
    Next i
End Sub

' V˝poËet pr˙dov a tokov v˝konu vo vedeniach po NR v˝poËte
' Vn˙torn˝ v˝poËet: v p.u.
' V˝stupy na list "vedenia":
'   F: |I_ij| [A]
'   G: ?U [%]        (V_from - V_to v p.u. * 100)
'   H: P_ij [MW]     (Ëinn˝ v˝kon z "uzol-od" do "uzol-do")
'   I: Q_ij [MVAr]   (jalov˝ v˝kon z "uzol-od" do "uzol-do")
'   J: P_str [kW]    (ËinnÈ straty na vedenÌ)
'
' R(), X()  ñ v p.u. (po prepoËte v LoadBranchData)
' Vmag(), Vang() ñ nap‰tia uzlov v p.u., Vang v radianoch
' SBase_MVA ñ b·za v˝konu
' BusBaseKV() - b·zy nap‰tÌ uzlov
Public Sub WriteBranchCurrents( _
    ByVal nBranches As Long, _
    ByRef FromBus() As Long, _
    ByRef ToBus() As Long, _
    ByRef R() As Double, _
    ByRef X() As Double, _
    ByRef BranchStatus() As Integer, _
    ByRef Vmag() As Double, _
    ByRef Vang() As Double, _
    ByVal SBase_MVA As Double, _
    ByRef BusBaseKV() As Double)

    Dim ws As Worksheet
    Dim k As Long
    Dim iBus As Long, jBus As Long
    Dim Vi As Complex, Vj As Complex
    Dim Zpu As Complex, Iij_pu As Complex
    Dim Iabs_pu As Double, Iabs_A As Double
    Dim Ibase_A As Double
    Dim dU_percent As Double
    Dim Sij_pu As Complex
    Dim Pij_pu As Double, Qij_pu As Double
    Dim Pij_MW As Double, Qij_MVAr As Double
    Dim Ploss_pu As Double, Ploss_kW As Double
    Dim Ubase_line As Double
    
    Set ws = ThisWorkbook.Worksheets("vedenia")
    
    ' hlaviËky stÂpcov - riadok 2, stÂpce Q(17)..U(21)
    ws.Cells(2, 17).Value = "|I_ij| [A]"
    ws.Cells(2, 18).Value = "?U [%]"
    ws.Cells(2, 19).Value = "P_ij [MW]"
    ws.Cells(2, 20).Value = "Q_ij [MVAr]"
    ws.Cells(2, 21).Value = "P_str [kW]"
    
    ' vetvy s˙ v riadkoch 3..(nBranches+2)
    For k = 1 To nBranches
        ' Ak je vedenie vypnutÈ (Status = 0), zapÌöeme nuly
        If BranchStatus(k) = 0 Then
            ws.Cells(2 + k, 17).Value = 0
            ws.Cells(2 + k, 18).Value = 0
            ws.Cells(2 + k, 19).Value = 0
            ws.Cells(2 + k, 20).Value = 0
            ws.Cells(2 + k, 21).Value = 0
        Else
            iBus = FromBus(k)
            jBus = ToBus(k)
            
            ' komplexnÈ nap‰tia uzlov v p.u. (pol·rny -> kartezi·nsky)
            ' Vmag je v p.u., Vang v radianoch -> CFromPolar oËak·va uhol v stupÚoch
            Vi = CFromPolar(Vmag(iBus), Vang(iBus) * RAD2DEG)
            Vj = CFromPolar(Vmag(jBus), Vang(jBus) * RAD2DEG)
            
            ' impedancia vetvy v p.u.
            Zpu = CCreate(R(k), X(k))
            
            If Abs(Zpu.Re) < 0.000000001 And Abs(Zpu.Im) < 0.000000001 Then
                ' nulov· impedancia ñ pr˙d a v˝kony nedefinovanÈ; nastavÌme 0
                Iabs_pu = 0#
                dU_percent = 0#
                Pij_pu = 0#
                Qij_pu = 0#
                Ploss_pu = 0#
            Else
                ' pr˙d vetvou v p.u.: I_ij = (Vi - Vj) / Z_ij
                Iij_pu = CDiv(CSub(Vi, Vj), Zpu)
                Iabs_pu = CAbs(Iij_pu)
                
                ' ˙bytok nap‰tia v % (z pohæadu "uzol-od" -> "uzol-do"):
                ' ?U [%] = (V_from_pu - V_to_pu) * 100
                dU_percent = (Vmag(iBus) - Vmag(jBus)) * 100#
                
                ' tok v˝konu z uzla-od do uzla-do:
                ' S_ij_pu = V_i * conj(I_ij)
                Sij_pu = CMul(Vi, CConj(Iij_pu))
                Pij_pu = Sij_pu.Re
                Qij_pu = Sij_pu.Im
                
                ' ËinnÈ straty v p.u.:
                ' P_str_pu = |I|^2 * R_pu
                Ploss_pu = Iabs_pu * Iabs_pu * R(k)
            End If
            
            ' UrËenie b·zy pr˙du pre vedenie (podæa "from" uzla)
            Ubase_line = BusBaseKV(iBus)
            If SBase_MVA <> 0# And Ubase_line <> 0# Then
                Ibase_A = (SBase_MVA * 1000#) / (Sqr(3#) * Ubase_line)
            Else
                Ibase_A = 1#
            End If
            
            ' pr˙d v ampÈroch
            Iabs_A = Iabs_pu * Ibase_A
            
            ' prepoËet v˝konov do MW / MVAr
            Pij_MW = Pij_pu * SBase_MVA
            Qij_MVAr = Qij_pu * SBase_MVA
            
            ' ËinnÈ straty do kW
            Ploss_kW = Ploss_pu * SBase_MVA * 1000#
            
            ' z·pis do riadku vetvy (riadky 3..) - zaokr˙hlenie na 2 des. miesta
            ws.Cells(2 + k, 17).Value = Round(Iabs_A, 2)
            ws.Cells(2 + k, 18).Value = Round(dU_percent, 2)
            ws.Cells(2 + k, 19).Value = Round(Pij_MW, 2)
            ws.Cells(2 + k, 20).Value = Round(Qij_MVAr, 2)
            ws.Cells(2 + k, 21).Value = Round(Ploss_kW, 2)
        End If
    Next k
End Sub

' NaËÌtanie d·t vedenÌ z listu "vedenia"
' SkutoËnÈ R, X [ohm] -> prepoËet do p.u. na dan˙ z·kladÚu
'
' Form·t "vedenia":
'   riadok 2: hlaviËka
'   riadky 3.. : d·ta
'   A: (voænÈ)
'   B: N·zov vedenia
'   C: uzol-od
'   D: uzol-do
'   E: Status (0/1/2)
'   N: R [ohm] (14)
'   O: X [ohm] (15)
'
' SBase_MVA
Public Sub LoadBranchData( _
    ByRef nBranches As Long, _
    ByRef BranchName() As String, _
    ByRef FromBus() As Long, _
    ByRef ToBus() As Long, _
    ByRef R() As Double, _
    ByRef X() As Double, _
    ByRef BranchStatus() As Integer, _
    ByRef BusNames() As String, _
    ByRef BusBaseKV() As Double, _
    ByVal SBase_MVA As Double)

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim fromName As String, toName As String
    Dim idx As Long
    Dim Zbase_ohm As Double
    Dim R_ohm As Double, X_ohm As Double
    Dim Ubase1 As Double
    Dim stVal As Variant
    
    Set ws = ThisWorkbook.Worksheets("vedenia")
    
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If lastRow < 3 Then
        nBranches = 0
        Exit Sub
    End If
    
    nBranches = lastRow - 2
    
    ReDim BranchName(1 To nBranches)
    ReDim FromBus(1 To nBranches)
    ReDim ToBus(1 To nBranches)
    ReDim R(1 To nBranches)
    ReDim X(1 To nBranches)
    ReDim BranchStatus(1 To nBranches)
    
    For i = 1 To nBranches
        ' D·ta od riadku 3 -> riadok = i + 2
        BranchName(i) = CStr(ws.Cells(i + 2, 2).Value) ' B: N·zov
        fromName = CStr(ws.Cells(i + 2, 3).Value)      ' C: Uzol od
        toName = CStr(ws.Cells(i + 2, 4).Value)        ' D: Uzol do
        
        ' NaËÌtanie Statusu z E (5)
        stVal = ws.Cells(i + 2, 5).Value
        If IsEmpty(stVal) Or Trim(CStr(stVal)) = "" Then
            BranchStatus(i) = 1 ' Default ON
        ElseIf IsNumeric(stVal) Then
            BranchStatus(i) = CInt(stVal)
        Else
            BranchStatus(i) = 0
        End If
        
        idx = GetBusIndex(fromName, BusNames)
        If idx = 0 Then
            Err.Raise vbObjectError + 2, , "Uzol '" & fromName & "' v liste 'vedenia', riadok " & (i + 2) & " neexistuje v liste 'uzly'."
        End If
        FromBus(i) = idx
        
        idx = GetBusIndex(toName, BusNames)
        If idx = 0 Then
            Err.Raise vbObjectError + 3, , "Uzol '" & toName & "' v liste 'vedenia', riadok " & (i + 2) & " neexistuje v liste 'uzly'."
        End If
        ToBus(i) = idx
        
        ' UrËenie Zbase pre vedenie (podæa "from" uzla)
        Ubase1 = BusBaseKV(idx)
        
        If SBase_MVA <> 0# Then
            Zbase_ohm = (Ubase1 * Ubase1) / SBase_MVA
        Else
            Zbase_ohm = 1#
        End If
        
        ' naËÌtanie skutoËnej impedancie - R z N(14), X z O(15)
        R_ohm = ParseDouble(ws.Cells(i + 2, 14).Value)
        X_ohm = ParseDouble(ws.Cells(i + 2, 15).Value)
        
        ' prepoËet do p.u.
        If Zbase_ohm <> 0# Then
            R(i) = R_ohm / Zbase_ohm
            X(i) = X_ohm / Zbase_ohm
        Else
            R(i) = 0#
            X(i) = 0#
        End If
    Next i
End Sub

' V˝poËet pr˙dov a tokov v˝konu v transform·toroch po NR v˝poËte
' Analogicky k WriteBranchCurrents, ale zohæadÚuje prevod
' V˝stupy na list "transformatory" od stÂpca X:
'   X: |I_prim| [A]
'   Y: |I_sec| [A]
'   Z: P_prim [MW] (do trafa z prim·ru)
'   AA: Q_prim [MVAr]
'   AB: P_sec [MW] (z trafa do sekund·ru)
'   AC: Q_sec [MVAr]
'   AD: P_str [kW]
Public Sub WriteTransformerFlows( _
    ByVal nTrafo As Long, _
    ByRef TrFrom() As Long, _
    ByRef TrTo() As Long, _
    ByRef TrR() As Double, _
    ByRef TrX() As Double, _
    ByRef TrG() As Double, _
    ByRef TrB() As Double, _
    ByRef TrRatio() As Double, _
    ByRef Vmag() As Double, _
    ByRef Vang() As Double, _
    ByRef BusBaseKV() As Double, _
    ByVal SBase_MVA As Double)

    Dim ws As Worksheet
    Dim k As Long
    Dim i As Long, J As Long
    Dim Vi As Complex, Vj As Complex
    Dim Zs As Complex, Ys As Complex
    Dim Ym As Complex
    Dim a As Double
    Dim I_prim As Complex, I_sec As Complex ' Pr˙dy v p.u.
    Dim S_prim As Complex, S_sec As Complex ' V˝kony v p.u.
    Dim Ibase_prim_A As Double, Ibase_sec_A As Double
    Dim Ubase1 As Double, Ubase2 As Double
    
    Dim Iprim_A As Double, Isec_A As Double
    Dim Pprim_MW As Double, Qprim_MVAr As Double
    Dim Psec_MW As Double, Qsec_MVAr As Double
    Dim Ploss_kW As Double
    
    Set ws = ThisWorkbook.Worksheets("transformatory")
    
    ' HlaviËky
    ws.Cells(2, 24).Value = "|I_prim| [A]"
    ws.Cells(2, 25).Value = "|I_sec| [A]"
    ws.Cells(2, 26).Value = "P_prim [MW]"
    ws.Cells(2, 27).Value = "Q_prim [MVAr]"
    ws.Cells(2, 28).Value = "P_sec [MW]"
    ws.Cells(2, 29).Value = "Q_sec [MVAr]"
    ws.Cells(2, 30).Value = "P_str [kW]"
    
    For k = 1 To nTrafo
        i = TrFrom(k) ' Prim·r (s odboËkou)
        J = TrTo(k)   ' Sekund·r
        a = TrRatio(k)
        
        Vi = CFromPolar(Vmag(i), Vang(i) * RAD2DEG)
        Vj = CFromPolar(Vmag(J), Vang(J) * RAD2DEG)
        
        Zs = CCreate(TrR(k), TrX(k))
        Ys = CDiv(CCreate(1#, 0#), Zs)
        Ym = CCreate(TrG(k), TrB(k)) ' PrieËna admitancia na prim·ri
        
        ' Pr˙d do prim·ru (uzol i):
        ' I_i = (Vi/a^2 + Vi*Ym - Vj/a) * ys ... nie presne takto z modelu
        ' Z modelu PI:
        ' I_i = Vi * (ys/a^2 + Ym) - Vj * (ys/a)
        ' I_j = Vj * (ys) - Vi * (ys/a)
        
        ' PomocnÈ pre prim·r:
        ' term1 = Vi * (ys/a^2 + Ym)
        Dim term1 As Complex, term2 As Complex
        Dim Yseries_a2 As Complex
        
        Yseries_a2 = CCreate(Ys.Re / (a * a), Ys.Im / (a * a))
        term1 = CMul(Vi, CAdd(Yseries_a2, Ym))
        
        ' term2 = Vj * (ys/a)
        Dim Yseries_a As Complex
        Yseries_a = CCreate(Ys.Re / a, Ys.Im / a)
        term2 = CMul(Vj, Yseries_a)
        
        I_prim = CSub(term1, term2)
        
        ' PomocnÈ pre sekund·r (tok DO siete z uzla j):
        ' I_j = Vj * ys - Vi * (ys/a)
        term1 = CMul(Vj, Ys)
        term2 = CMul(Vi, Yseries_a)
        I_sec = CSub(term1, term2)
        
        ' V˝kony
        S_prim = CMul(Vi, CConj(I_prim))
        S_sec = CMul(Vj, CConj(I_sec))
        
        ' B·zy pr˙du pre prim·r a sekund·r
        Ubase1 = BusBaseKV(i)
        Ubase2 = BusBaseKV(J)
        
        If SBase_MVA <> 0# And Ubase1 <> 0# Then
            Ibase_prim_A = (SBase_MVA * 1000#) / (Sqr(3#) * Ubase1)
        Else
            Ibase_prim_A = 1#
        End If
        
        If SBase_MVA <> 0# And Ubase2 <> 0# Then
            Ibase_sec_A = (SBase_MVA * 1000#) / (Sqr(3#) * Ubase2)
        Else
            Ibase_sec_A = 1#
        End If
        
        ' Absol˙tne hodnoty pr˙dov
        Iprim_A = CAbs(I_prim) * Ibase_prim_A
        Isec_A = CAbs(I_sec) * Ibase_sec_A
        
        ' V˝kony v MW/MVAr
        Pprim_MW = S_prim.Re * SBase_MVA
        Qprim_MVAr = S_prim.Im * SBase_MVA
        
        ' Pozor: P_sec definujeme ako tok Z trafa DO sekund·ru?
        ' ätandardne v tabuæk·ch tokov sa pÌöe tok z i->j.
        ' Tu m·me P_prim (vstup do trafa) a P_sec (v˝stup z trafa = -S_sec, ak S_sec je injekcia do siete)
        ' Alebo jednoducho zapÌöeme injekciu do uzla j (S_sec).
        ' Dohodnime sa: P_sec bude tok smerom OD uzla j DO trafa (Ëiûe S_sec tak ako vyölo),
        ' alebo tok smerom Z trafa DO uzla j?
        ' Obvykle sa ud·va tok na zaËiatku a na konci vedenia.
        ' Takûe P_sec = P_ji = S_sec.Re * SBase_MVA.
        Psec_MW = S_sec.Re * SBase_MVA
        Qsec_MVAr = S_sec.Im * SBase_MVA
        
        ' Straty = S_prim + S_sec
        Ploss_kW = (S_prim.Re + S_sec.Re) * SBase_MVA * 1000#
        
        ' Z·pis
        ws.Cells(2 + k, 24).Value = Round(Iprim_A, 2)
        ws.Cells(2 + k, 25).Value = Round(Isec_A, 2)
        ws.Cells(2 + k, 26).Value = Round(Pprim_MW, 2)
        ws.Cells(2 + k, 27).Value = Round(Qprim_MVAr, 2)
        ws.Cells(2 + k, 28).Value = Round(Psec_MW, 2)
        ws.Cells(2 + k, 29).Value = Round(Qsec_MVAr, 2)
        ws.Cells(2 + k, 30).Value = Round(Ploss_kW, 2)
    Next k
End Sub

' Z·pis reportu o izolovan˝ch uzloch a vedeniach
Public Sub WriteIsolationReport( _
    ByVal nBuses As Long, _
    ByRef BusNames() As String, _
    ByRef IsBusIsolated() As Boolean, _
    ByVal nBranches As Long, _
    ByRef FromBus() As Long, _
    ByRef ToBus() As Long, _
    ByRef IsBranchIsolated() As Boolean, _
    ByVal nTrafo As Long, _
    ByRef TrFrom() As Long, _
    ByRef TrTo() As Long, _
    ByRef IsTrafoIsolated() As Boolean, _
    ByVal nComp As Long, _
    ByRef CompBus() As Long, _
    ByRef IsCompIsolated() As Boolean)

    Dim ws As Worksheet
    Dim R As Long
    Dim i As Long
    
    Set ws = GetOrCreateSheet("report")
    ws.Cells.Clear
    
    R = 2
    ws.Cells(R, 2).Value = "IzolovanÈ uzly"
    ws.Cells(R, 2).Font.Bold = True
    R = R + 1
    
    For i = 1 To nBuses
        If IsBusIsolated(i) Then
            ws.Cells(R, 2).Value = BusNames(i)
            R = R + 1
        End If
    Next i
    
    R = R + 1 ' Jeden voæn˝ riadok
    ws.Cells(R, 2).Value = "IzolovanÈ vedenia"
    ws.Cells(R, 2).Font.Bold = True
    R = R + 1
    
    For i = 1 To nBranches
        If IsBranchIsolated(i) Then
            ' VypÌöeme "From -> To"
            ws.Cells(R, 2).Value = BusNames(FromBus(i)) & " -> " & BusNames(ToBus(i))
            R = R + 1
        End If
    Next i
    
    ' MÙûeme pridaù aj izolovanÈ transform·tory
    For i = 1 To nTrafo
        If IsTrafoIsolated(i) Then
            ws.Cells(R, 2).Value = "Trafo: " & BusNames(TrFrom(i)) & " -> " & BusNames(TrTo(i))
            R = R + 1
        End If
    Next i
    
    ' IzolovanÈ kompenz·cie
    If nComp > 0 Then
        R = R + 1
        ws.Cells(R, 2).Value = "IzolovanÈ kompenz·cie"
        ws.Cells(R, 2).Font.Bold = True
        R = R + 1
        
        For i = 1 To nComp
            If IsCompIsolated(i) Then
                ws.Cells(R, 2).Value = "Kompenz·cia na uzle: " & BusNames(CompBus(i))
                R = R + 1
            End If
        Next i
    End If
    
    ' Motory VN nie s˙ explicitne izolovanÈ objekty (s˙ to spotrebiËe),
    ' ale mÙûu byù na izolovan˝ch uzloch.
    
    ' Z·pis pozn·mky do listu "uzly" stÂpec H (8)
    Dim wsUzly As Worksheet
    Set wsUzly = ThisWorkbook.Worksheets("uzly")
    For i = 1 To nBuses
        If IsBusIsolated(i) Then
            ' Ak je izolovan˝, do stÂpca H napÌöeme "izolovane"
            ' Pozor: WriteFinalVoltagesToUzly prepisuje stÂpec H.
            ' MusÌme zabezpeËiù poradie volania, alebo to zapÌsaù po WriteFinalVoltages.
            wsUzly.Cells(2 + i, 8).Value = "izolovane"
            wsUzly.Cells(2 + i, 9).Value = "-"
        End If
    Next i
End Sub


