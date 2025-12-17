Attribute VB_Name = "Module5"
'==========================
' Modul: modYBus
'==========================
Option Explicit

' Vytvoriù Y-bus maticu z vedenÌ
Public Sub BuildYBus( _
    ByVal nBuses As Long, _
    ByVal nBranches As Long, _
    ByRef FromBus() As Long, _
    ByRef ToBus() As Long, _
    ByRef R() As Double, _
    ByRef X() As Double, _
    ByRef BusNames() As String, _
    ByRef Y() As Complex, _
    ByRef G() As Double, _
    ByRef B() As Double)

    Dim i As Long, J As Long, k As Long
    Dim Z As Complex, Yline As Complex
    
    ReDim Y(1 To nBuses, 1 To nBuses)
    ReDim G(1 To nBuses, 1 To nBuses)
    ReDim B(1 To nBuses, 1 To nBuses)
    
    For i = 1 To nBuses
        For J = 1 To nBuses
            Y(i, J) = CCreate(0#, 0#)
        Next J
    Next i
    
    For k = 1 To nBranches
        If (R(k) = 0# And X(k) = 0#) Then
            ' nulov· impedancia ñ preskoË
        Else
            Z = CCreate(R(k), X(k))
            Yline = CDiv(CCreate(1#, 0#), Z)
            
            i = FromBus(k)
            J = ToBus(k)
            
            Y(i, i) = CAdd(Y(i, i), Yline)
            Y(J, J) = CAdd(Y(J, J), Yline)
            Y(i, J) = CSub(Y(i, J), Yline)
            Y(J, i) = CSub(Y(J, i), Yline)
        End If
    Next k
    
    For i = 1 To nBuses
        For J = 1 To nBuses
            G(i, J) = Y(i, J).Re
            B(i, J) = Y(i, J).Im
        Next J
    Next i
    
    Call WriteYMatrix(Y, G, B, BusNames)
End Sub

