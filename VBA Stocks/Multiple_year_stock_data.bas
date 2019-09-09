Attribute VB_Name = "Module1"
Sub stockMultiple():
    
    Dim sheet As Worksheet
    
    For Each sheet In ThisWorkbook.Worksheets
    
    Dim ticker As String
    Dim yearChange As String
    Dim percentChange As String
    Dim totalStock As String
    Dim Value As String
    
    ticker = "Ticker"
    yearChange = "Yearly Change"
    percentChange = "Percent Change"
    totalStock = "Total Stock Volume"
    Value = "Value"
    
    
    sheet.Range("I1, P1").Value = ticker
    sheet.Cells(1, 10).Value = yearChange
    sheet.Cells(1, 11).Value = percentChange
    sheet.Cells(1, 12).Value = totalStock
    sheet.Range("Q1").Value = Value
    
    'Dim header As Worksheet
    
        'Set header = ThisWorkbook.Sheets("A")
        'Sheets.FillAcrossSheets header.Range("1:1")
        

    
    Dim lastRow As Long
    lastRow = sheet.Cells(Rows.count, 1).End(xlUp).Row
    
    Dim tickerName As String
    
    Dim ticker_total As Long
    ticker_total_volume = 0
        
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    Dim yearly_change As Double
    yearly_change = 0
    
    Dim open_price As Double
    open_price = 0
    
    Dim close_price As Double
    close_price = 0
    
    Dim price_difference As Double
    price_difference = 0
        
    
    For i = 2 To lastRow
    
        If sheet.Cells(i + 1, 1).Value <> sheet.Cells(i, 1).Value Then
        
            tickerName = sheet.Cells(i, 1).Value
            
            ticker_total_volume = ticker_total_volume + sheet.Cells(i, 7).Value
            
            sheet.Range("I" & Summary_Table_Row).Value = tickerName
            
            sheet.Range("L" & Summary_Table_Row).Value = ticker_total_volume
            
            Summary_Table_Row = Summary_Table_Row + 1
            
            ticker_total_volume = 0
            
        Else
        
            ticker_total_volume = ticker_total_volume + sheet.Cells(i, 7).Value
            
        End If
    
    Next i
       
    Summary_Table_Row = 2
    
    Dim count As Double
    count = -1
       
    For i = 2 To lastRow
    
        If sheet.Cells(i + 1, 1).Value <> sheet.Cells(i, 1).Value Then
        
            close_price = sheet.Cells(i, 6).Value
            'MsgBox ("close " & close_price)
            
            price_difference = close_price - open_price
            
            sheet.Range("J" & Summary_Table_Row).Value = price_difference
            
                If open_price = 0 Then
                
                    sheet.Range("K" & Summary_Table_Row).Value = 1
                    
                    sheet.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                    
                    Summary_Table_Row = Summary_Table_Row + 1
                    
                    count = -1
                    
                Else
            
                    sheet.Range("K" & Summary_Table_Row).Value = (price_difference / open_price)
            
                    sheet.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
            
                    Summary_Table_Row = Summary_Table_Row + 1
                    
                    'open_price = Cells(i + 1, 3).Value
                    
                    count = -1
                
                End If
        
        Else
            
            count = count + 1
            open_price = sheet.Cells(i - count, 3).Value
            'MsgBox ("open " & open_price)
            
        End If
    
    Next i
    
    
    Dim table_row As Double
    table_row = sheet.Cells(Rows.count, 9).End(xlUp).Row
    
    For i = 2 To table_row
            
        If sheet.Cells(i, 10).Value > 0 Then
            
            sheet.Cells(i, 10).Interior.ColorIndex = 4
                
        Else
            
            sheet.Cells(i, 10).Interior.ColorIndex = 3
                
        End If
        
    Next i
    
    
    Dim Great_Vol As Double
    Great_Vol = sheet.Range("L2").Value
    
    Dim Great_Vol_Tick As String
    
    For i = 3 To table_row
    
        If sheet.Cells(i, 12).Value > Great_Vol Then
        
            Great_Vol = sheet.Cells(i, 12).Value
            Great_Vol_Tick = sheet.Cells(i, 9).Value
            
        Else
            
        End If
        
    Next i
    
    sheet.Cells(4, 17).Value = Great_Vol
    sheet.Cells(4, 16).Value = Great_Vol_Tick
    sheet.Cells(4, 15).Value = "Greatest Total Volume"
    
    Dim Great_Per As Double
    Great_Per = sheet.Cells(2, 11).Value
    
    Dim Great_Per_Tick As String
    
    For i = 3 To table_row
    
        If (sheet.Cells(i, 11).Value > Great_Per) Then
        
            Great_Per = sheet.Cells(i, 11).Value
            sheet.Range("Q2").NumberFormat = "0.00%"
            Great_Per_Tick = sheet.Cells(i, 9).Value
            
        Else
        
        End If
        
    Next i
    
    sheet.Cells(2, 17).Value = Great_Per
    sheet.Cells(2, 16).Value = Great_Per_Tick
    sheet.Cells(2, 15).Value = "Greatest % Increase"
    
    
    Dim Less_Per As Double
    Less_Per = sheet.Cells(2, 11).Value
    
    Dim Less_Per_Tick As String
    
    For i = 3 To table_row
    
        If (sheet.Cells(i, 11).Value < Less_Per) Then
        
            Less_Per = sheet.Cells(i, 11).Value
            sheet.Range("Q3").NumberFormat = "0.00%"
            Less_Per_Tick = sheet.Cells(i, 9).Value
            
        Else
        
        End If
        
    Next i
    
    sheet.Cells(3, 17).Value = Less_Per
    sheet.Cells(3, 16).Value = Less_Per_Tick
    sheet.Cells(3, 15).Value = "Greatest % Decrease"
    
    
    sheet.Cells.EntireColumn.AutoFit
        
    Next sheet

End Sub

