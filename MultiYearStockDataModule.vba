Attribute VB_Name = "Module1"
Sub MultiYearStockData()

'Cycle/Loop Thru All Worksheets
Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate
'Establish Last Row of Each Spreadsheet
    LRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
    
'Establish First Summary Table Header Titles
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percentage Change"
    Cells(1, 12).Value = "Total Stock Volume"
'Establish Second Summary Table Header and Row Titles
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
 'Declare my Variables
    
    Dim Tkr As String
    Dim OpenP As Double
    Dim CloseP As Double
    Dim YrCh As Double
    YrCh = 0
    Dim PerCh As Double
    PerCh = 0
    Dim Vol As Double
    Vol = 0
    Dim Row As Long
    Row = 2
    Dim i As Long
    Dim j As Integer
    j = 0
    Dim Start As Long
    Start = 2
    
'Step 1 of Calculations -- OpenP (First Row of Each Ticker & Pulling From Column 3 (<Open>)

        OpenP = Cells(Start, 3).Value
        
'Cycle/Loop Thru to Next Ticker Value & Establish CloseP
    For i = 2 To LRow
    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            Vol = Vol + Cells(i, 7).Value
            If Vol = 0 Then
                Range("I" & 2 + j).Value = Cells(i, 1).Value
                Range("J" & 2 + j).Value = 0
                Range("K" & 2 + j).Value = "%" & 0
                Range("L" & 2 + j).Value = 0
            
            Else
                If OpenP = 0 Then
                    For FindValue = Start To i
                        If Cells(FindValue, 3).Value <> 0 Then
                            Start = FindValue
                            Exit For
                        End If
                    Next FindValue
                End If
                OpenP = Cells(Start, 3).Value
                Tkr = Cells(Row, 1).Value
                'Cells(Row, 9).Value = Tkr
                CloseP = Cells(i, 6).Value
'Start Summary Table for Each Ticker Value--Starting With YrCh
                YrCh = CloseP - OpenP
                'Cells(Row, 10).Value = YrCh
'Summary Table PerCh
          If OpenP > 0 Then
             PerCh = Round(YrCh / OpenP * 100, 2)
           Else
             PerCh = 0
           End If
             
                'Cells(Row, 11).Value = PerCh
                'Cells(Row, 11).NumberFormat = "0.00%"
                
'Summary Table Vol
                Cells(Row, 12).Value = Vol
'Add One to the Summary Table Row
                Row = Row + 1
'Reset the OpenP; Reset Vol to Zero
                Start = i + 1
        
'If the Cells Immediately Following a Row Contain the Same Ticker
                Range("I" & 2 + j).Value = Cells(i, 1).Value
                Range("J" & 2 + j).Value = Round(YrCh, 2)
                Range("K" & 2 + j).Value = "%" & PerCh
                Range("L" & 2 + j).Value = Vol
            End If
            j = j + 1
            Vol = 0
        Else
'Add to the Vol
            Vol = Vol + Cells(i, 7).Value
        End If
   Next i
 
'Establish the Last Row of YrCh For Each Sheet and Begin Conditional Color Formatting
        YrCh_LRow = WS.Cells(Rows.Count, 9).End(xlUp).Row
        For j = 2 To YrCh_LRow
            If (Cells(j, 10).Value >= 0) Then
            Cells(j, 10).Interior.ColorIndex = 4
            ElseIf Cells(j, 10).Value < 0 Then
            Cells(j, 10).Interior.ColorIndex = 3
            End If
        Next j
'Build Second Summary Table

        For k = 2 To YrCh_LRow
            If Cells(k, 11).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & YrCh_LRow)) Then
                Cells(2, 16).Value = Cells(k, 9).Value
                Cells(2, 17).Value = Cells(k, 11).Value
                Cells(2, 17).NumberFormat = "0.00%"
            ElseIf Cells(k, 11).Value = Application.WorksheetFunction.Min(WS.Range("K2:K" & YrCh_LRow)) Then
                Cells(3, 16).Value = Cells(k, 9).Value
                Cells(3, 17).Value = Cells(k, 11).Value
                Cells(3, 17).NumberFormat = "0.00%"
            ElseIf Cells(k, 12).Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & YrCh_LRow)) Then
                Cells(4, 16).Value = Cells(k, 9).Value
                Cells(4, 17).Value = Cells(k, 12).Value
             End If
        Next k
    Next WS

End Sub


