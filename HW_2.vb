Sub VBA_HW()

Dim SummaryTableRow As Integer
Dim StartPx As Double
Dim EndPx As Double
Dim StockVolumn As Double

'Set initial values
SummaryTableRow = 2
StockVolumn = 0

'Create combined sheet and format headers
Sheets.Add.Name = "Combined Data"
Sheets("Combined Data").Move before:=Sheets(1)
Set combined_sheet = Worksheets("Combined Data")

combined_sheet.Range("A1").Value = "Ticker"
combined_sheet.Range("B1").Value = "Year"
combined_sheet.Range("C1").Value = "Yearly Change"
combined_sheet.Range("D1").Value = "Percent Change"
combined_sheet.Range("E1").Value = "Total Stock Volume"

'iterate through the ws
For Each ws In Worksheets
    If ws.Name <> "Combined Data" Then

        'Sort the ws by Ticker then Date
        ws.Sort.SortFields.Add Key:=Range("A1"), Order:=xlAscending
        ws.Sort.SortFields.Add Key:=Range("B1"), Order:=xlAscending
        ws.Sort.SetRange Range("A:G")
        ws.Sort.Header = xlYes
        ws.Sort.Apply

        'Add column headers
        ws.Range("H1").Value = "Ticker"
        ws.Range("I1").Value = "Year"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"

        'Set Starting Px to first value
        StartPx = ws.Range("C2").Value

        'iterate through the rows in the ws
        For i = 2 To ws.Range("A1").End(xlDown).Row

            'If the ticker is the same then add the volumn to total stock volumn
            If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
            StockVolumn = StockVolumn + ws.Cells(i, 7).Value

            'If the ticker changes, add the last volumn to total volumn, set last px, populate and format summary table
            Else
            StockVolumn = StockVolumn + ws.Cells(i, 7).Value
            EndPx = ws.Cells(i, 6).Value

            ws.Range("H" & SummaryTableRow).Value = ws.Cells(i, 1).Value
            ws.Range("I" & SummaryTableRow).Value = ws.Name
            ws.Range("J" & SummaryTableRow).Value = StartPx - EndPx
            ws.Range("L" & SummaryTableRow).Value = StockVolumn

                If StartPx = 0 Then
                    ws.Range("K" & SummaryTableRow).Value = NA
                Else
                    ws.Range("K" & SummaryTableRow).Value = (StartPx - EndPx) / StartPx
                End If

            ws.Range("K:K").Style = "Percent"
            ws.Range("L:L").NumberFormat = "0.00E+00"
                
                'Add color conditional, green for positive changes and red for negative changes
                If ws.Range("J" & SummaryTableRow).Value > 0 Then
                    ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3
                End If

            'Reset Variables for next ticker
            StockVolumn = 0
            SummaryTableRow = SummaryTableRow + 1
            StartPx = ws.Cells(i + 1, 3).Value

            End If
        Next i
    'Find number of data rows in each ws (-1 to remove header)
    ws_DataRows = ws.Range("H1").End(xlDown).Row - 1

    'Find the last row populated in combined sheet and add 1 to get first empty row
    lastrow = combined_sheet.Cells(Rows.Count, "A").End(xlUp).Row + 1

    'Copy rows to combined sheet, I had to use the copy paste option due to issues with the sheer amount of data being moved
    Set CopySource = ws.Range("H2:L" & (ws_DataRows + 1))

    Set DestinationRng = combined_sheet.Range("A" & lastrow & ":E" & (lastrow + ws_DataRows - 1))

    CopySource.Copy
    'Use xlPasteAll to keep the coloring format from the individual worksheets
    DestinationRng.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

    End If
    'Reset summary table row for each ws
    SummaryTableRow = 2
Next ws

combined_sheet.Columns("A:G").AutoFit

'Create a button on the combined data sheet that will give you the option to find the extreme values on each sheet
Sheets("Combined Data").Buttons.Delete

celLeft = Sheets("Combined Data").Range("G2").Left
celTop = Sheets("Combined Data").Range("H2").Top
celWidth = Sheets("Combined Data").Range("G2:J2").Width
celHeight = Sheets("Combined Data").Range("G2:G7").Height

Set btn = Sheets("Combined Data").Buttons.Add(celLeft, celTop, celWidth, celHeight)
With btn
    'The button will run the FindExtremeCases Macro
    .OnAction = "FindExtremeCases"
    .Caption = "Find Extremes"
    .Name = "Extremes"
End With
    
End Sub

Sub FindExtremeCases()

Dim MaxIncrease As Double
Dim MaxDecrease As Double
Dim MaxVolumn As Double

'iterate throught the ws
For Each ws In Worksheets
    
    'set up the summary table
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Totals Volumn"
    
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"
    
    'For the combined data sheet, include a column in the summary table for year
    If ws.Name = "Combined Data" Then
    
    ws_length = ws.Range("A1").End(xlDown).Row
    
    ws.Range("Q1").Value = "Year"
    
    'find the max/min percent changes and the max total volumn and insert into summary table
    MaxIncrease = ws.Application.WorksheetFunction.Max(ws.Range("D2:D" & ws_length).Value)
    MaxDecrease = ws.Application.WorksheetFunction.Min(ws.Range("D2:D" & ws_length).Value)
    MaxVolumn = ws.Application.WorksheetFunction.Max(ws.Range("E2:E" & ws_length).Value)
    
    ws.Range("P2").Value = MaxIncrease
    ws.Range("P3").Value = MaxDecrease
    ws.Range("P4").Value = MaxVolumn
    
    'Find the tickers and years associated with the values and insert into summary table
    For i = 2 To ws_length
        If ws.Cells(i, 4).Value = MaxIncrease Then
            ws.Range("O2").Value = ws.Cells(i, 1).Value
            ws.Range("Q2").Value = ws.Cells(i, 2).Value
        ElseIf ws.Cells(i, 4).Value = MaxDecrease Then
            ws.Range("O3").Value = ws.Cells(i, 1).Value
            ws.Range("Q3").Value = ws.Cells(i, 2).Value
        ElseIf ws.Cells(i, 5).Value = MaxVolumn Then
            ws.Range("O4").Value = ws.Cells(i, 1).Value
            ws.Range("Q4").Value = ws.Cells(i, 2).Value
        End If
    Next i
    
    'for all other ws, do not find the year associated with the values
    Else
    
    ws_length = ws.Range("H1").End(xlDown).Row
    
    'find the max/min percent changes and the max total volumn and insert into summary table
    MaxIncrease = ws.Application.WorksheetFunction.Max(ws.Range("K2:K" & ws_length).Value)
    MaxDecrease = ws.Application.WorksheetFunction.Min(ws.Range("K2:K" & ws_length).Value)
    MaxVolumn = ws.Application.WorksheetFunction.Max(ws.Range("L2:L" & ws_length).Value)
    
    ws.Range("P2").Value = MaxIncrease
    ws.Range("P3").Value = MaxDecrease
    ws.Range("P4").Value = MaxVolumn
    
    'Find the tickers and years associated with the values and insert into summary table
    For i = 2 To ws.Range("H1").End(xlDown).Row
        If ws.Cells(i, 11).Value = MaxIncrease Then
            ws.Range("O2").Value = ws.Cells(i, 8).Value
        ElseIf ws.Cells(i, 11).Value = MaxDecrease Then
            ws.Range("O3").Value = ws.Cells(i, 8).Value
        ElseIf ws.Cells(i, 12).Value = MaxVolumn Then
            ws.Range("O4").Value = ws.Cells(i, 8).Value
        End If
    Next i
    End If
    
    'Format the summary table
    ws.Range("P2").Style = "Percent"
    ws.Range("P3").Style = "Percent"
    ws.Range("P4").NumberFormat = "0.00E+00"
    
    ws.Columns("A:Q").AutoFit
Next ws

End Sub




