Attribute VB_Name = "Module1"
Sub Stock_Market():

    ' Set Variable for Ticker Name.
    Dim Ticker As String
    
    ' Set Variable For Opening Price
    Dim OpenRate As Double
    
    ' Set Variable for Closing Price
    Dim CloseRate As Double
    
    ' Set Variable for Yearly Change
    Dim YearlyChange As Double
    
    ' Set Variable for % Change
    Dim PercentChange As Double
    
    ' Set Variable for Total Volume
    Dim TotalVol As LongLong
    TotalVol = 0
    
    ' Loop Through all Sheets
    For Each ws In Worksheets
    
      ' Set OpenRate Value
      OpenRate = ws.Cells(2, 3).Value
       
        ' Set Up Table
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ' Location of each Ticker in Table
        Dim TableRow As Integer
        TableRow = 2
        
        ' Find Last Row of Column "A"
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
            ' Loop through all Stocks
            For i = 2 To LastRow - 1
               
                ' Check for last row of Ticker
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
                 ' Set Ticker Name
                 Ticker = ws.Cells(i, 1).Value
                
                 ' Add to TotalVol
                 TotalVal = TotalVal + ws.Cells(i, 7).Value
        
                 ' Set CloseRate Value
                 CloseRate = ws.Cells(i, 6).Value
        
                 ' Determine Yearly Change
                 YearlyChange = (CloseRate - OpenRate)
                
                 ' Determine Percent Change
                 PercentChange = YearlyChange / OpenRate
                
                 ' Print Ticker, YearlyChange, PercentChange and Total Volume in Table
                 ws.Range("I" & TableRow).Value = Ticker
                 ws.Range("J" & TableRow).Value = YearlyChange
                 ws.Range("K" & TableRow).Value = PercentChange
                 ws.Range("L" & TableRow).Value = TotalVol
                    
                    ' Set Formatting for PercentChange
                    If ws.Range("J" & TableRow).Value >= 0 Then
                    ws.Range("J" & TableRow).Interior.ColorIndex = 4
                    Else
                    ws.Range("J" & TableRow).Interior.ColorIndex = 3
                    End If
                    
                 'Define Next Row in Table
                 TableRow = TableRow + 1
                 
                 ' Reset TotalVol
                 TotalVol = 0
                 
                 ' Set Next OpenRate
                 OpenRate = ws.Cells(i + 1, 3).Value
                
                ' If same Ticker
                Else
                TotalVol = TotalVol + ws.Cells(i, 7).Value
                
                End If
            
            Next i
            
            ' Make Percent Change Column into % Format
            ws.Range("K:K").Style = "Percent"
            ws.Range("K:K").NumberFormat = "0.00%"
            
            ' Create Table of Greatest % Increase, Greatest % Decrease, and Greatest Total volume.
            ' Set Up Table
            ws.Range("O1").Value = "Ticker"
            ws.Range("P1").Value = "Value"
            ws.Range("N2").Value = "Greatest % Increase"
            ws.Range("N3").Value = "Greatest % Decrease"
            ws.Range("N4").Value = "Greatest Total Volume"
            
            ' Find Last Row of Column "I"
            Last = ws.Cells(Rows.Count, 9).End(xlUp).Row
            
        ' Find Greatest % Increase
        ' Loop through all Table Rows
        For i = 2 To Last
            
            If ws.Cells(i, 11).Value > ws.Range("P2").Value Then
            ' Overwrite Value with Highest Value
            ws.Range("P2").Value = ws.Cells(i, 11).Value
            ' Determine Ticker Name
            ws.Range("O2").Value = ws.Cells(i, 9).Value
            
            End If
            
        Next i
            
        ' Find Greatest % Decrease
        ' Loop through all Table Rows
        For i = 2 To Last
            
            If ws.Cells(i, 11).Value < ws.Range("P3").Value Then
            ' Overwrite Value with Highest Value
            ws.Range("P3").Value = ws.Cells(i, 11).Value
            ' Determine Ticker Name
            ws.Range("O3").Value = ws.Cells(i, 9).Value
            
            End If
            
        Next i
            
        ' Find Greatest Total Volume
        ' Loop through all Table Rows
        For i = 2 To Last
            
            If ws.Cells(i, 12).Value > ws.Range("P4").Value Then
            ' Overwrite Value with Highest Value
            ws.Range("P4").Value = ws.Cells(i, 12).Value
            ' Determine Ticker Name
            ws.Range("O4").Value = ws.Cells(i, 9).Value
            
            End If
            
        Next i
            
            ' Format Table
            ws.Range("P2:P3").Style = "Percent"
            ws.Range("P2:P3").NumberFormat = "0.00%"
            
            ' Make Both Table's Columns Autofit
            ws.Columns("I:P").AutoFit
        Next ws
        
End Sub
