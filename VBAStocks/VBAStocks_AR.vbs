'This is the main sub. Run this macro to run the VBAStocks program
Public Sub Run_VBAStock():
    
    Dim curws As Worksheet
    
    'Loop through each worksheet
    For Each curws In Worksheets
        
        'Delete any existing summary tables, create column headers and fill the table
        curws.Activate
        curws.Columns("I:Q").Delete
        
        Fill_Summary_Tables
        
    Next

End Sub

'Create the summary table with headers and appropriate formatting
Private Sub Create_Summary_Table_Headers():
    
    Dim ws As Worksheet
    Dim rn As Range
    
    Set ws = ActiveSheet
    Set rn = ws.UsedRange
    
    'Set column headers in summary table
    ws.Cells(1, 9) = "Ticker"
    ws.Cells(1, 10) = "Yearly Change"
    ws.Cells(1, 11) = "Percentage Change"
    ws.Cells(1, 12) = "Total Stock Volume"
    
    ws.Cells(1, 16) = "Ticker"
    ws.Cells(1, 17) = "Greatest Value"
    ws.Cells(2, 15) = "Greatest % Increase"
    ws.Cells(3, 15) = "Greatest % Decrease"
    ws.Cells(4, 15) = "Greatest Total Volume"
    
    'Format columns/cells in summary table
    ws.Columns(10).NumberFormat = "0.00000000"
    ws.Columns(11).NumberFormat = "0.00%"
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
    
End Sub
'Fill summary table for each ticker
'Show yearly change, percentage change and total volume
Public Sub Fill_Summary_Tables():

    Dim ws As Worksheet
    Dim wb As Workbook
    Dim rn As Range
    
    Dim i As Long 'Total Rows in usedrange
    Dim j As Long 'Counter to loop through usedrange
    Dim k As Long 'Used to fill summary table - Ticker
    Dim l As Long 'Used to fill summary table - Yearly Change
    
    Dim locTicker As String
    Dim locOpenVal As Variant
    Dim locCloseVal As Variant
    Dim locTotStockVol As Variant
            
    Set ws = ActiveSheet
    Set rn = ws.UsedRange
    
    Create_Summary_Table_Headers
    
    i = rn.Rows.Count 'Identify the number of rows in the spreadsheet
    j = 2 'Start at the second row, ommiting the header row
    k = 1 'First row in the summary table
        
    Do While j <= i

        'If the ticker has changed from previous row - new entry needs to be made
        If ws.Range("A" & j) <> ws.Range("A" & j - 1) Then
            locTicker = ws.Range("A" & j).Value
            locTotStockVol = ws.Range("G" & j).Value
            locOpenVal = ws.Range("C" & j).Value
            locCloseVal = ws.Range("F" & j).Value
            k = k + 1
        Else
            'locTicker = ws.Range("A" & j).Value
            locTotStockVol = locTotStockVol + ws.Range("G" & j).Value
            locCloseVal = ws.Range("F" & j).Value
        End If
    
        'Fill Ticker, Yearly Change, Percentage Change and Volume
        ws.Cells(k, 9) = locTicker
        ws.Cells(k, 10) = locCloseVal - locOpenVal
        If locCloseVal <> 0 Then
            ws.Cells(k, 11) = (locCloseVal - locOpenVal) / locCloseVal
        End If
        ws.Cells(k, 12) = locTotStockVol
        
        j = j + 1
    
    Loop
         
    'Add conditional formatting after cleaning up old rules
    ws.Range("J2:J" & i).FormatConditions.Delete
    With ws.Range("J2:J" & k)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual _
        , Formula1:="=0"
        With .FormatConditions(1).Interior
            .Color = RGB(0, 255, 0)
        End With
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess _
        , Formula1:="=0"
        With .FormatConditions(2).Interior
            .Color = RGB(255, 0, 0)
        End With
    End With
    
    ws.Columns.AutoFit
        
End Sub