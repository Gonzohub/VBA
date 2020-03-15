Attribute VB_Name = "Module1"
Sub Stocks():
    
    Dim r As Long
    Dim LastRow As Long
    Dim TickRow As Long
    Dim OpenP As Double
    Dim CloseP As Double
    Dim TChange As Double
    Dim perChange As Double
    
    Range("I1").Value = "TickName"
    Range("J1").Value = "Total Volume"
    Range("K1").Value = "Yearly Change"
    Range("L1").Value = "% Change"
    
    r = 2
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    TickRow = 2
    OpenP = Cells(2, 3).Value
    
    For r = 2 To LastRow
        
        tV = tV + Cells(r, 7).Value
        
        If Cells(r + 1, 1).Value <> Cells(r, 1).Value Then
                
                TickName = Cells(r, 1).Value
                Range("I" & TickRow).Value = TickName
                Range("J" & TickRow).Value = tV
                
                tV = 0
                
                CloseP = Cells(r, 6).Value
                
                TChange = CloseP - OpenP
                
                If (OpenP = 0 And CloseP = 0) Then
                    perChange = TChange / 1
                ElseIf (OpenP = 0 And CloseP <> 0) Then
                    perChange = TChange
                Else
                    perChange = TChange / OpenP
                End If
                
                Range("K" & TickRow).Value = TChange
                Range("K" & TickRow).NumberFormat = "0.00"
                Range("L" & TickRow).Value = perChange
                Range("L" & TickRow).NumberFormat = "0.00%"
                
                OpenP = Cells(r + 1, 3).Value
                
                If perChange >= 0 Then
                    Range("L" & TickRow).Interior.Color = vbGreen
                Else
                    Range("L" & TickRow).Interior.Color = vbRed
                End If
             
                TickRow = TickRow + 1
             
            End If
    Next r
        
End Sub

