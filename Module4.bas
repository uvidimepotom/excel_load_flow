Attribute VB_Name = "Module4"
'==========================
' Modul: modIO
'==========================
Option Explicit

' Naèítanie dát uzlov z listu "uzly"
' Skutoèné hodnoty -> prepoèet do pomerných (p.u.)
'
' Formát listu "uzly":
'   riadok 2: hlavièka
'   riadky 3.. : dáta
'   B: Názov uzla
'   C: Typ (Slack / PQ)
'   D: |V| [kV]
'   E: ? [deg]
'   F: P [MW]
'   G: Q [Mvar]
'
' SBase_MVA, UBase_kV sú bázové hodnoty zo "index"
Public Sub LoadBusData( _
    ByRef nBuses As Long, _
    ByRef BusNames() As String, _
    ByRef BusTypes() As BusType, _
    ByRef Vmag() As Double, _
    ByRef Vang() As Double, _
    ByRef Pspec() As Double, _
    ByRef Qspec() As Double, _
    ByVal SBase_MVA As Double, _
    ByVal UBase_kV As Double)

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim t As String
    Dim V_kV As Double, P_MW As Double, Q_Mvar As Double
    
    Set ws = ThisWorkbook.Worksheets("uzly")
    
    ' posledný riadok pod¾a ståpca B (Názov uzla)
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If lastRow < 3 Then
        Err.Raise vbObjectError + 1, , "V liste 'uzly' nie sú žiadne uzly (oèakávam dáta od riadku 3)."
    End If
    
    ' prvé dáta sú v riadku 3 => poèet uzlov
    nBuses = lastRow - 2
    
    ReDim BusNames(1 To nBuses)
    ReDim BusTypes(1 To nBuses)
    ReDim Vmag(1 To nBuses)
    ReDim Vang(1 To nBuses)
    ReDim Pspec(1 To nBuses)
    ReDim Qspec(1 To nBuses)
    
    For i = 1 To nBuses
        ' riadok s dátami = 2 + i (3,4,...)
        BusNames(i) = CStr(ws.Cells(2 + i, 2).Value)   ' B: Názov uzla
        
        t = CStr(ws.Cells(2 + i, 3).Value)             ' C: Typ
        Select Case UCase$(Trim$(t))
            Case "SLACK"
                BusTypes(i) = btSlack
            Case "PQ"
                BusTypes(i) = btPQ
            Case Else
                BusTypes(i) = btPQ
        End Select
        
        ' naèítanie skutoèných hodnôt
        V_kV = ParseDouble(ws.Cells(2 + i, 4).Value)   ' D: |V| [kV]
        P_MW = ParseDouble(ws.Cells(2 + i, 6).Value)   ' F: P [MW]
        Q_Mvar = ParseDouble(ws.Cells(2 + i, 7).Value) ' G: Q [Mvar]
        
        ' prepoèet do p.u.
        If UBase_kV <> 0# Then
            Vmag(i) = V_kV / UBase_kV
        Else
            Vmag(i) = 1#
        End If
        
        Vang(i) = ParseDouble(ws.Cells(2 + i, 5).Value) * DEG2RAD  ' E: ? [deg] -> rad
        
        If SBase_MVA <> 0# Then
            Pspec(i) = P_MW / SBase_MVA
            Qspec(i) = Q_Mvar / SBase_MVA
        Else
            Pspec(i) = 0#
            Qspec(i) = 0#
        End If
    Next i
End Sub

' Naèítanie dát vedení z listu "vedenia"
' Skutoèné R, X [ohm] -> prepoèet do p.u. na danú základòu
'
' Formát "vedenia":
'   riadok 1: hlavièka
'   riadky 2.. : dáta
'   B: uzol-od
'   C: uzol-do
'   D: R [ohm]
'   E: X [ohm]
'
' SBase_MVA, UBase_kV sú bázové hodnoty zo "index"
Public Sub LoadBranchData( _
    ByRef nBranches As Long, _
    ByRef FromBus() As Long, _
    ByRef ToBus() As Long, _
    ByRef R() As Double, _
    ByRef X() As Double, _
    ByRef BusNames() As String, _
    ByVal SBase_MVA As Double, _
    ByVal UBase_kV As Double)

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim fromName As String, toName As String
    Dim idx As Long
    Dim Zbase_ohm As Double
    Dim R_ohm As Double, X_ohm As Double
    
    Set ws = ThisWorkbook.Worksheets("vedenia")
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        nBranches = 0
        Exit Sub
    End If
    
    nBranches = lastRow - 1
    
    ReDim FromBus(1 To nBranches)
    ReDim ToBus(1 To nBranches)
    ReDim R(1 To nBranches)
    ReDim X(1 To nBranches)
    
    ' báza impedancie v ohmoch
    If SBase_MVA <> 0# Then
        Zbase_ohm = (UBase_kV * UBase_kV) / SBase_MVA
    Else
        Zbase_ohm = 1#
    End If
    
    For i = 1 To nBranches
        fromName = CStr(ws.Cells(i + 1, 2).Value)
        toName = CStr(ws.Cells(i + 1, 3).Value)
        
        idx = GetBusIndex(fromName, BusNames)
        If idx = 0 Then
            Err.Raise vbObjectError + 2, , "Uzol '" & fromName & "' v liste 'vedenia', riadok " & (i + 1) & " neexistuje v liste 'uzly'."
        End If
        FromBus(i) = idx
        
        idx = GetBusIndex(toName, BusNames)
        If idx = 0 Then
            Err.Raise vbObjectError + 3, , "Uzol '" & toName & "' v liste 'vedenia', riadok " & (i + 1) & " neexistuje v liste 'uzly'."
        End If
        ToBus(i) = idx
        
        ' naèítanie skutoènej impedancie
        R_ohm = ParseDouble(ws.Cells(i + 1, 4).Value)
        X_ohm = ParseDouble(ws.Cells(i + 1, 5).Value)
        
        ' prepoèet do p.u.
        If Zbase_ohm <> 0# Then
            R(i) = R_ohm / Zbase_ohm
            X(i) = X_ohm / Zbase_ohm
        Else
            R(i) = 0#
            X(i) = 0#
        End If
    Next i
End Sub


' Zápis admitanènej matice a pomocných blokov G a B
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
    
    ' hlavièky ståpcov
    For J = 1 To n
        ws.Cells(row0, col0 + J).Value = BusNames(J)
    Next J
    
    ' hlavièky riadkov + samotná matica G+jB
    For i = 1 To n
        ws.Cells(row0 + i, col0).Value = BusNames(i)
        For J = 1 To n
            txt = Format(G(i, J), "0.000000") & IIf(B(i, J) >= 0, "+j" & Format(B(i, J), "0.000000"), "-j" & Format(Abs(B(i, J)), "0.000000"))
            ws.Cells(row0 + i, col0 + J).Value = txt
        Next J
    Next i
    
    ' pomocný blok G
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
    
    ' pomocný blok B
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

' Vymazanie / príprava výsledkových listov "napatia" a "epsilon"
Public Sub ClearResultsSheets()
    Dim wsV As Worksheet, wsE As Worksheet
    
    Set wsV = GetOrCreateSheet("napatia")
    Set wsE = GetOrCreateSheet("epsilon")
    
    wsV.Cells.Clear
    wsE.Cells.Clear
    
    ' hlavièky
    wsV.Cells(1, 1).Value = "Iterácia"
    wsV.Cells(1, 2).Value = "Uzol"
    wsV.Cells(1, 3).Value = "|V| [p.u.]"
    wsV.Cells(1, 4).Value = "? [deg]"
    
    wsE.Cells(1, 1).Value = "Iterácia"
    wsE.Cells(1, 2).Value = "max|?P|"
    wsE.Cells(1, 3).Value = "max|?Q|"
    wsE.Cells(1, 4).Value = "epsilon"
End Sub

' Logovanie napätí v jednej iterácii
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

' Logovanie epsilon v jednej iterácii
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

' Zápis súhrnných výsledkov do listu "index"
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

' Zápis výsledných napätí z NR výpoètu na list "uzly"
' H (ståpec 8): |V| výp. [kV]
' I (ståpec 9): ? výp. [deg]
'
' Vmag(i) je v p.u., prepoèet na kV cez UBase_kV
Public Sub WriteFinalVoltagesToUzly( _
    ByRef Vmag() As Double, _
    ByRef Vang() As Double, _
    ByVal UBase_kV As Double)

    Dim ws As Worksheet
    Dim i As Long
    Dim nBuses As Long
    
    Set ws = ThisWorkbook.Worksheets("uzly")
    
    nBuses = UBound(Vmag)
    
    ' hlavièky výsledkov (riadok 2)
    ws.Cells(2, 8).Value = "|V| výp. [kV]"
    ws.Cells(2, 9).Value = "? výp. [deg]"
    
    ' dáta od riadku 3
    For i = 1 To nBuses
        ws.Cells(2 + i, 8).Value = Vmag(i) * UBase_kV
        ws.Cells(2 + i, 9).Value = Vang(i) * RAD2DEG
    Next i
End Sub

' Výpoèet prúdov a tokov výkonu vo vedeniach po NR výpoète
' Vnútorný výpoèet: v p.u.
' Výstupy na list "vedenia":
'   F: |I_ij| [A]
'   G: ?U [%]        (V_from - V_to v p.u. * 100)
'   H: P_ij [MW]     (èinný výkon z "uzol-od" do "uzol-do")
'   I: Q_ij [MVAr]   (jalový výkon z "uzol-od" do "uzol-do")
'   J: P_str [kW]    (èinné straty na vedení)
'
' R(), X()  – v p.u. (po prepoète v LoadBranchData)
' Vmag(), Vang() – napätia uzlov v p.u., Vang v radianoch
' SBase_MVA, UBase_kV – bázy zo sheetu "index"
Public Sub WriteBranchCurrents( _
    ByVal nBranches As Long, _
    ByRef FromBus() As Long, _
    ByRef ToBus() As Long, _
    ByRef R() As Double, _
    ByRef X() As Double, _
    ByRef Vmag() As Double, _
    ByRef Vang() As Double, _
    ByVal SBase_MVA As Double, _
    ByVal UBase_kV As Double)

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
    
    Set ws = ThisWorkbook.Worksheets("vedenia")
    
    ' báza prúdu v ampéroch:
    ' I_base [A] = S_base [MVA] * 1e3 / (sqrt(3) * U_base [kV])
    If SBase_MVA <> 0# And UBase_kV <> 0# Then
        Ibase_A = (SBase_MVA * 1000#) / (Sqr(3#) * UBase_kV)
    Else
        Ibase_A = 1#
    End If
    
    ' hlavièky ståpcov
    ws.Cells(1, 6).Value = "|I_ij| [A]"
    ws.Cells(1, 7).Value = "?U [%]"
    ws.Cells(1, 8).Value = "P_ij [MW]"
    ws.Cells(1, 9).Value = "Q_ij [MVAr]"
    ws.Cells(1, 10).Value = "P_str [kW]"
    
    ' vetvy sú v riadkoch 2..(nBranches+1)
    For k = 1 To nBranches
        iBus = FromBus(k)
        jBus = ToBus(k)
        
        ' komplexné napätia uzlov v p.u. (polárny -> karteziánsky)
        ' Vmag je v p.u., Vang v radianoch -> CFromPolar oèakáva uhol v stupòoch
        Vi = CFromPolar(Vmag(iBus), Vang(iBus) * RAD2DEG)
        Vj = CFromPolar(Vmag(jBus), Vang(jBus) * RAD2DEG)
        
        ' impedancia vetvy v p.u.
        Zpu = CCreate(R(k), X(k))
        
        If Abs(Zpu.Re) < 0.000000001 And Abs(Zpu.Im) < 0.000000001 Then
            ' nulová impedancia – prúd a výkony nedefinované; nastavíme 0
            Iabs_pu = 0#
            dU_percent = 0#
            Pij_pu = 0#
            Qij_pu = 0#
            Ploss_pu = 0#
        Else
            ' prúd vetvou v p.u.: I_ij = (Vi - Vj) / Z_ij
            Iij_pu = CDiv(CSub(Vi, Vj), Zpu)
            Iabs_pu = CAbs(Iij_pu)
            
            ' úbytok napätia v % (z poh¾adu "uzol-od" -> "uzol-do"):
            ' ?U [%] = (V_from_pu - V_to_pu) * 100
            dU_percent = (Vmag(iBus) - Vmag(jBus)) * 100#
            
            ' tok výkonu z uzla-od do uzla-do:
            ' S_ij_pu = V_i * conj(I_ij)
            Sij_pu = CMul(Vi, CConj(Iij_pu))
            Pij_pu = Sij_pu.Re
            Qij_pu = Sij_pu.Im
            
            ' èinné straty v p.u.:
            ' P_str_pu = |I|^2 * R_pu
            Ploss_pu = Iabs_pu * Iabs_pu * R(k)
        End If
        
        ' prúd v ampéroch
        Iabs_A = Iabs_pu * Ibase_A
        
        ' prepoèet výkonov do MW / MVAr
        Pij_MW = Pij_pu * SBase_MVA
        Qij_MVAr = Qij_pu * SBase_MVA
        
        ' èinné straty do kW
        Ploss_kW = Ploss_pu * SBase_MVA * 1000#
        
        ' zápis do riadku vetvy (riadky 2..)
        ws.Cells(1 + k, 6).Value = Iabs_A
        ws.Cells(1 + k, 7).Value = dU_percent
        ws.Cells(1 + k, 8).Value = Pij_MW
        ws.Cells(1 + k, 9).Value = Qij_MVAr
        ws.Cells(1 + k, 10).Value = Ploss_kW
    Next k
End Sub


