Attribute VB_Name = "Module7"
'==========================
' Modul: modMain
'==========================
Option Explicit

' Tlaèidlo na tvorbu Y-matice
Public Sub CmdBuildYMatrix()
    On Error GoTo ErrHandler
    
    Dim nBuses As Long, nBranches As Long
    Dim BusNames() As String
    Dim BusTypes() As BusType
    Dim Vmag() As Double, Vang() As Double
    Dim Pspec() As Double, Qspec() As Double
    Dim FromBus() As Long, ToBus() As Long
    Dim R() As Double, X() As Double
    Dim Y() As Complex
    Dim G() As Double, B() As Double
    Dim SBase_MVA As Double, UBase_kV As Double
    
    ' bázy
    Call GetBaseValues(SBase_MVA, UBase_kV)
    
    ' uzly a vetvy v p.u.
    Call LoadBusData(nBuses, BusNames, BusTypes, Vmag, Vang, Pspec, Qspec, SBase_MVA, UBase_kV)
    Call LoadBranchData(nBranches, FromBus, ToBus, R, X, BusNames, SBase_MVA, UBase_kV)
    
    Call BuildYBus(nBuses, nBranches, FromBus, ToBus, R, X, BusNames, Y, G, B)
    
    MsgBox "Admitanèná matica bola vytvorená.", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Chyba pri tvorbe Y-matice: " & Err.Description, vbCritical
End Sub


' Tlaèidlo na spustenie NR load-flow
Public Sub CmdRunNR()
    Call NewtonRaphsonLoadFlow
End Sub

' Makro pre VBS – kompletný beh: Y-matica + NR
Public Sub RunFullLoadFlow()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo Cleanup
    
    Call CmdBuildYMatrix
    Call CmdRunNR

Cleanup:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
End Sub

