Attribute VB_Name = "Module3"
'==========================
' Modul: modUtils
'==========================
Option Explicit

' Univerzálne spracovanie èísel – ak obsahujú èiarku alebo bodku
Public Function ParseDouble(ByVal v As Variant) As Double
    Dim s As String
    Dim decSep As String
    Dim otherSep As String
    
    On Error GoTo FailSafe
    
    If IsEmpty(v) Or IsNull(v) Then
        ParseDouble = 0#
        Exit Function
    End If
    
    If IsNumeric(v) Then
        ParseDouble = CDbl(v)
        Exit Function
    End If
    
    s = CStr(v)
    s = Trim$(s)
    If s = "" Then
        ParseDouble = 0#
        Exit Function
    End If
    
    decSep = Application.International(xlDecimalSeparator)
    If decSep = "." Then
        otherSep = ","
    Else
        otherSep = "."
    End If
    
    s = Replace$(s, " ", "")
    s = Replace$(s, otherSep, decSep)
    
    ParseDouble = CDbl(s)
    Exit Function
    
FailSafe:
    ParseDouble = 0#
End Function

' Vyh¾adanie indexu uzla pod¾a názvu (case-insensitive)
Public Function GetBusIndex(ByVal busName As String, ByRef BusNames() As String) As Long
    Dim i As Long
    For i = LBound(BusNames) To UBound(BusNames)
        If StrComp(Trim$(BusNames(i)), Trim$(busName), vbTextCompare) = 0 Then
            GetBusIndex = i
            Exit Function
        End If
    Next i
    GetBusIndex = 0 ' nenašiel sa
End Function

' Nájde prvý vo¾ný riadok v zadanom ståpci
Public Function FirstFreeRow(ByVal ws As Worksheet, ByVal col As Long) As Long
    With ws
        If .Cells(.Rows.Count, col).End(xlUp).Row < 2 Then
            FirstFreeRow = 2
        Else
            FirstFreeRow = .Cells(.Rows.Count, col).End(xlUp).Row + 1
        End If
    End With
End Function

' Bezpeèné získanie (prípadne vytvorenie) výstupného listu
Public Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = sheetName
    End If
    
    Set GetOrCreateSheet = ws
End Function

' Naèítanie bázových hodnôt zo sheetu "index"
' B10: S_base [MVA]
' B11: U_base [kV]
Public Sub GetBaseValues(ByRef SBase_MVA As Double, ByRef UBase_kV As Double)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("index")
    
    SBase_MVA = ParseDouble(ws.Range("B10").Value)
    UBase_kV = ParseDouble(ws.Range("B11").Value)
    
    ' jednoduché defaulty, ak by bunky boli prázdne alebo <= 0
    If SBase_MVA <= 0# Then SBase_MVA = 100#   ' 100 MVA
    If UBase_kV <= 0# Then UBase_kV = 110#    ' 110 kV
End Sub

