Attribute VB_Name = "modSLD"
'==========================
' Modul: modSLD
'==========================
Option Explicit

' Hlavn· proced˙ra pre aktualiz·ciu SLD
Public Sub UpdateSLD()
    Dim startTime As Double
    Dim wsSLD As Worksheet, wsRep As Worksheet, wsIndex As Worksheet
    Dim rng As Range, cell As Range
    Dim cellVal As String
    Dim tag As String, varType As String, direction As String
    Dim val As Double
    Dim found As Boolean
    Dim rowIdx As Long, colIdx As Long
    Dim R As Long, c As Long
    Dim reportRow As Long
    
    ' Dictionaries pre v˝sledky
    ' Kæ˙Ë: N·zov zariadenia (bez prefixu? Alebo s prefixom ak je v n·zve?
    ' Podæa zadania: TAG je "N_NazovUzla", reùazec "N_NazovUzla_P_R".
    ' Takûe v Dictionary bude kæ˙Ë "N_NazovUzla".
    Dim dictNodes As Object
    Dim dictLines As Object
    Dim dictTrafo As Object
    Dim dictReac As Object
    Dim dictDifReac As Object
    Dim dictComp As Object
    Dim dictMotor As Object
    Dim dictGen As Object
    
    On Error GoTo ErrHandler
    
    startTime = Timer
    
    Set wsSLD = GetOrCreateSheet("SLD")
    Set wsIndex = GetOrCreateSheet("index")
    
    ' Inicializ·cia reportu ch˝b
    Set wsRep = GetOrCreateSheet("SLD_Report")
    wsRep.Cells.Clear
    wsRep.Cells(1, 1).Value = "Bunka"
    wsRep.Cells(1, 2).Value = "Reùazec"
    wsRep.Cells(1, 3).Value = "Chyba"
    reportRow = 2
    
    ' NaËÌtanie v˝sledkov do pam‰te
    Set dictNodes = CreateObject("Scripting.Dictionary")
    Set dictLines = CreateObject("Scripting.Dictionary")
    Set dictTrafo = CreateObject("Scripting.Dictionary")
    Set dictReac = CreateObject("Scripting.Dictionary")
    Set dictDifReac = CreateObject("Scripting.Dictionary")
    Set dictComp = CreateObject("Scripting.Dictionary")
    Set dictMotor = CreateObject("Scripting.Dictionary")
    Set dictGen = CreateObject("Scripting.Dictionary")
    
    Call LoadResultsToDict(dictNodes, dictLines, dictTrafo, dictReac, dictDifReac, dictComp, dictMotor, dictGen)
    
    ' Iter·cia cez oblasù A1:AS200
    ' AS je stÂpec 45
    ' Prehæad·vame po bunk·ch.
    ' Optimaliz·cia: »Ìtanie vlastnostÌ Font.Color je pomalÈ.
    ' Sk˙sime ËÌtaù hodnoty do poæa, ale farbu musÌme testovaù na objekte Range.
    
    Application.ScreenUpdating = False
    
    For R = 1 To 200
        For c = 1 To 45 ' A..AS
            Set cell = wsSLD.Cells(R, c)
            
            ' Kontrola farby pÌsma (biela = 16777215 alebo vbWhite)
            If cell.Font.Color = vbWhite Then
                cellVal = Trim(CStr(cell.Value))
                If Len(cellVal) > 0 Then
                    ' Parsovanie reùazca
                    If ParseTagString(cellVal, tag, varType, direction) Then
                        ' Vyhæadanie hodnoty
                        found = False
                        val = 0
                        
                        ' Rozhodovanie podæa prefixu TAGu
                        If Left(tag, 2) = "N_" Then
                            found = GetValueFromDict(dictNodes, tag, varType, val)
                        ElseIf Left(tag, 2) = "V_" Then
                            found = GetValueFromDict(dictLines, tag, varType, val)
                        ElseIf Left(tag, 2) = "T_" Then
                            found = GetValueFromDict(dictTrafo, tag, varType, val)
                        ElseIf Left(tag, 3) = "DR_" Then
                            found = GetValueFromDict(dictDifReac, tag, varType, val)
                        ElseIf Left(tag, 2) = "R_" Then
                            found = GetValueFromDict(dictReac, tag, varType, val)
                        ElseIf Left(tag, 2) = "K_" Then
                            found = GetValueFromDict(dictComp, tag, varType, val)
                        ElseIf Left(tag, 2) = "M_" Then
                            found = GetValueFromDict(dictMotor, tag, varType, val)
                        ElseIf Left(tag, 2) = "G_" Then
                            found = GetValueFromDict(dictGen, tag, varType, val)
                        Else
                            ' Nezn·my prefix
                            wsRep.Cells(reportRow, 1).Value = cell.Address
                            wsRep.Cells(reportRow, 2).Value = cellVal
                            wsRep.Cells(reportRow, 3).Value = "Nezn·my prefix zariadenia"
                            reportRow = reportRow + 1
                            GoTo NextCell
                        End If
                        
                        If found Then
                            ' Z·pis do cieæovej bunky
                            Dim targetR As Long, targetC As Long
                            targetR = R: targetC = c
                            Select Case UCase(direction)
                                Case "R": targetC = c + 1
                                Case "L": targetC = c - 1
                                Case "U": targetR = R - 1
                                Case "D": targetR = R + 1
                            End Select
                            
                            If targetR > 0 And targetC > 0 Then
                                wsSLD.Cells(targetR, targetC).Value = Round(val, 2)
                            End If
                        Else
                            ' Hodnota nen·jden· (zlÈ meno alebo premenn·)
                            wsRep.Cells(reportRow, 1).Value = cell.Address
                            wsRep.Cells(reportRow, 2).Value = cellVal
                            wsRep.Cells(reportRow, 3).Value = "Hodnota nen·jden· (Tag: " & tag & ", Var: " & varType & ")"
                            reportRow = reportRow + 1
                        End If
                    Else
                        ' Chyba parsovania
                        wsRep.Cells(reportRow, 1).Value = cell.Address
                        wsRep.Cells(reportRow, 2).Value = cellVal
                        wsRep.Cells(reportRow, 3).Value = "Chybn˝ form·t reùazca"
                        reportRow = reportRow + 1
                    End If
                End If
            End If
NextCell:
        Next c
    Next R
    
    ' Z·pis Ëasu
    wsIndex.Range("B5").Value = Timer - startTime
    
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Chyba pri aktualiz·cii SLD: " & Err.Description, vbCritical
End Sub

' Parsovanie reùazca "TAG_X_Y"
' Y je smer (posledn˝ segment)
' X je premenn· (predposledn˝ segment)
' TAG je zvyöok
Private Function ParseTagString(ByVal s As String, ByRef tag As String, ByRef varType As String, ByRef direction As String) As Boolean
    Dim parts() As String
    Dim n As Long
    
    parts = Split(s, "_")
    n = UBound(parts)
    
    If n < 2 Then
        ParseTagString = False
        Exit Function
    End If
    
    direction = parts(n)
    varType = parts(n - 1)
    
    ' Zloûenie TAGu zo zvyön˝ch ËastÌ (0 aû n-2)
    Dim i As Long
    tag = parts(0)
    For i = 1 To n - 2
        tag = tag & "_" & parts(i)
    Next i
    
    ParseTagString = True
End Function

' Pomocn· funkcia na zÌskanie hodnoty z Dictionary
' Value je pole hodnÙt alebo objekt. Tu predpoklad·m pole Variant/Double indexovanÈ n·zvom premennej?
' Alebo Dictionary v Dictionary?
' Pre jednoduchosù: Value v hlavnom dict bude Dictionary(VarName -> Value)
Private Function GetValueFromDict(ByVal mainDict As Object, ByVal tag As String, ByVal varType As String, ByRef outVal As Double) As Boolean
    If mainDict.Exists(tag) Then
        Dim props As Object
        Set props = mainDict(tag)
        If props.Exists(varType) Then
            outVal = props(varType)
            GetValueFromDict = True
            Exit Function
        End If
    End If
    GetValueFromDict = False
End Function

' NaËÌtanie vöetk˝ch v˝sledkov
Private Sub LoadResultsToDict(ByRef dN As Object, ByRef dL As Object, ByRef dT As Object, ByRef dR As Object, ByRef dDR As Object, ByRef dC As Object, ByRef dM As Object, ByRef dG As Object)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim name As String, key As String
    Dim props As Object
    
    ' 1. Uzly ("uzly")
    ' Prefix "N_"
    ' StÂpce: B(2)=Name, J(10)=Ik3, K(11)=P_in, L(12)=Q_in, M(13)=I_in, H(8)=V_kV, I(9)=Ang
    ' Pozn·mka: Zadanie hovorÌ "P, Q, I" pre uzol. MyslÌ sa P_in (bilancia)? ¡no.
    ' Tieû Ik3.
    Set ws = ThisWorkbook.Worksheets("uzly")
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    For i = 3 To lastRow
        name = CStr(ws.Cells(i, 2).Value)
        key = "N_" & name
        Set props = CreateObject("Scripting.Dictionary")
        props("Ik3") = ParseDouble(ws.Cells(i, 10).Value)
        props("P") = ParseDouble(ws.Cells(i, 11).Value)
        props("Q") = ParseDouble(ws.Cells(i, 12).Value)
        props("I") = ParseDouble(ws.Cells(i, 13).Value)
        props("U") = ParseDouble(ws.Cells(i, 8).Value) ' Nap‰tie
        Set dN(key) = props
    Next i
    
    ' 2. Vedenia ("vedenia")
    ' Prefix "V_"
    ' StÂpce: B(2)=Name, Q(17)=I, R(18)=dU, S(19)=P, T(20)=Q, U(21)=Ploss
    Set ws = ThisWorkbook.Worksheets("vedenia")
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    For i = 3 To lastRow
        name = CStr(ws.Cells(i, 2).Value)
        key = "V_" & name
        Set props = CreateObject("Scripting.Dictionary")
        props("I") = ParseDouble(ws.Cells(i, 17).Value)
        props("dU") = ParseDouble(ws.Cells(i, 18).Value)
        props("P") = ParseDouble(ws.Cells(i, 19).Value)
        props("Q") = ParseDouble(ws.Cells(i, 20).Value)
        props("Pstr") = ParseDouble(ws.Cells(i, 21).Value)
        Set dL(key) = props
    Next i
    
    ' 3. Transform·tory ("transformatory")
    ' Prefix "T_"
    ' Predpoklad·m, ûe n·zov trafa nie je v liste "transformatory" explicitne (v LoadTransformerData nebol).
    ' Alebo je v stÂpci B? PÙvodn˝ kÛd ËÌtal C(From), D(To).
    ' AK nie je n·zov, musÌme ho vytvoriù "From-To"?
    ' Alebo uûÌvateæ prid· stÂpec B? Zadanie pre SLD hovorÌ "T_Nazov".
    ' SKONTROLOVAç: LoadTransformerData ËÌta "ws.Cells(i + 2, 3)" ako FromName.
    ' StÂpec B (2) zvyËajne b˝va Name.
    ' V pÙvodnom `LoadTransformerData` sa n·zov nenaËÌtaval.
    ' Predpokladajme, ûe v liste "transformatory" je v B n·zov.
    ' StÂpce v˝sledkov: X(24)=Iprim, Y(25)=Isec, Z(26)=Pprim, AA(27)=Qprim, AB(28)=Psec, AC(29)=Qsec, AD(30)=Pstr
    Set ws = ThisWorkbook.Worksheets("transformatory")
    lastRow = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row ' Podæa From
    For i = 3 To lastRow
        name = CStr(ws.Cells(i, 2).Value) ' Predpoklad B
        If name = "" Then name = "Trafo" & (i - 2) ' Fallback
        key = "T_" & name
        Set props = CreateObject("Scripting.Dictionary")
        props("Ip") = ParseDouble(ws.Cells(i, 24).Value)
        props("Is") = ParseDouble(ws.Cells(i, 25).Value)
        props("Pp") = ParseDouble(ws.Cells(i, 26).Value)
        props("Qp") = ParseDouble(ws.Cells(i, 27).Value)
        props("Ps") = ParseDouble(ws.Cells(i, 28).Value)
        props("Qs") = ParseDouble(ws.Cells(i, 29).Value)
        props("Pstr") = ParseDouble(ws.Cells(i, 30).Value)
        Set dT(key) = props
    Next i
    
    ' 4. Reaktory ("reaktory")
    ' Prefix "R_"
    ' B(2)=Name. V˝sledky: Z(26)=I, AA(27)=dU, AB(28)=P, AC(29)=Q, AD(30)=Pstr
    Set ws = ThisWorkbook.Worksheets("reaktory")
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    For i = 3 To lastRow
        name = CStr(ws.Cells(i, 2).Value)
        key = "R_" & name
        Set props = CreateObject("Scripting.Dictionary")
        props("I") = ParseDouble(ws.Cells(i, 26).Value)
        props("dU") = ParseDouble(ws.Cells(i, 27).Value)
        props("P") = ParseDouble(ws.Cells(i, 28).Value)
        props("Q") = ParseDouble(ws.Cells(i, 29).Value)
        props("Pstr") = ParseDouble(ws.Cells(i, 30).Value)
        Set dR(key) = props
    Next i
    
    ' 5. Dif. Reaktory ("dif_reaktory")
    ' Prefix "DR_"
    ' N·zov? LoadDifReactorData d·val "DR" & i.
    ' Ale v liste mÙûe byù stÂpec B (voæn˝/Name)?
    ' PÙvodn˝ kÛd: B nie je pouûit˝ (C=From).
    ' Predpokladajme, ûe uûÌvateæ tam d· n·zov do B.
    ' V˝sledky: X(24)=I, Y(25)=dU, Z(26)=P, AA(27)=Q, AB(28)=Pstr
    Set ws = ThisWorkbook.Worksheets("dif_reaktory")
    lastRow = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
    For i = 3 To lastRow
        name = CStr(ws.Cells(i, 2).Value)
        If name = "" Then name = "DR" & (i - 2)
        key = "DR_" & name
        Set props = CreateObject("Scripting.Dictionary")
        props("I") = ParseDouble(ws.Cells(i, 24).Value)
        props("dU") = ParseDouble(ws.Cells(i, 25).Value)
        props("P") = ParseDouble(ws.Cells(i, 26).Value)
        props("Q") = ParseDouble(ws.Cells(i, 27).Value)
        props("Pstr") = ParseDouble(ws.Cells(i, 28).Value)
        Set dDR(key) = props
    Next i
    
    ' 6. Kompenz·cia ("kompenz·cia")
    ' Prefix "K_"
    ' B(2)=Name. V˝sledok: O(15)=U [kV].
    ' »o Ôalöie? P, Q?
    ' V WriteCompResults sa pÌöe len U.
    ' V NR sa poËÌta Q_flow. Ale nezapisuje sa do riadku kompenz·cie, len do sumy uzla.
    ' Ak chceme Q kompenz·cie, musÌme dopoËÌtaù: Q = U^2 * B_comp.
    ' Zatiaæ implementujem U. Ak bude treba Q, treba doplniù v˝poËet.
    Set ws = ThisWorkbook.Worksheets("kompenz·cia")
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    For i = 3 To lastRow
        name = CStr(ws.Cells(i, 2).Value)
        key = "K_" & name
        Set props = CreateObject("Scripting.Dictionary")
        props("U") = ParseDouble(ws.Cells(i, 15).Value)
        ' DoplnÌme Q pre ˙plnosù (pribliûne)
        ' Q [Mvar] = (U[kV])^2 * B[S] ? Nie, B v liste je v p.u. alebo S?
        ' LoadCompData naËÌta XC, XL.
        Set dC(key) = props
    Next i
    
    ' 7. Motory VN ("motoryVN")
    ' Prefix "M_"
    ' B(2)=Name. V˝sledky: AD(30)=I, AE(31)=Ploss.
    ' P, Q?
    Set ws = ThisWorkbook.Worksheets("motoryVN")
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    For i = 2 To lastRow
        name = CStr(ws.Cells(i, 2).Value)
        key = "M_" & name
        Set props = CreateObject("Scripting.Dictionary")
        props("I") = ParseDouble(ws.Cells(i, 30).Value)
        props("Pstr") = ParseDouble(ws.Cells(i, 31).Value)
        Set dM(key) = props
    Next i
    
    ' 8. Gener·tory ("generatory")
    ' Prefix "G_"
    ' B(2)=Name. V˝sledok: N(14)=Q_gen. M(13)=P_gen.
    Set ws = ThisWorkbook.Worksheets("generatory")
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    For i = 3 To lastRow
        name = CStr(ws.Cells(i, 2).Value)
        key = "G_" & name
        Set props = CreateObject("Scripting.Dictionary")
        props("Q") = ParseDouble(ws.Cells(i, 14).Value)
        props("P") = ParseDouble(ws.Cells(i, 13).Value)
        Set dG(key) = props
    Next i
    
End Sub

