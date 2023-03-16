Attribute VB_Name = "Module1"
Sub TickerInfo()
    ' Set up to iterate through worksheets
    For Each ws In Worksheets
        
        ' Worksheet specific variables
        Dim WorksheetName As String
        Dim TotalRecords As LongLong
        TotalRecords = ws.Cells(Rows.Count, 1).End(xlUp).Row ' Collect the total number of rows with data in it and save the count for this worksheet.
        WorksheetName = ws.Name
        ' MsgBox ("New Worksheet Started NAME: " + WorksheetName) ' If user wants to know when the loop moves to the next worksheet uncomment the previous msgbox
        
        'Data and placement variables
        Dim Ticker As String
        Dim OpenPrice As Double
        Dim ClosePrice As Double
        Dim Volume As LongLong
        Dim PlacefromStart As Integer
        PlacefromStart = 2
        
        ' Arrays to hold both highest and lowest percent change as well as highest total volume.
        Dim TopTickers(0 To 2) As String
        Dim TopAmounts(0 To 2) As Double
        
        ' Initialize and re-initialize top amounts every worksheet.
        ' This makes sure that we are only comparing worksheets to their own data and not the data from prior worksheets.
        TopAmounts(0) = 0
        TopAmounts(1) = 0
        TopAmounts(2) = 0
        
        Dim Header(0 To 9) As String
        Header(0) = "Ticker"
        Header(1) = "Yearly Change"
        Header(2) = "Percent Change"
        Header(3) = "Total Stock Volume"
        ' First sequence of headers end here
        Header(4) = "Greatest % Increase"
        Header(5) = "Greatest % Decrease"
        Header(6) = "Greatest Total Volume"
        ' Second sequence of headers end here
        Header(7) = ""
        Header(8) = "Ticker"
        Header(9) = "Value"
    
        ' -------------------------------------------------------------
        
        ' Print new header
        For i = 9 To 12
            ws.Cells(1, i).Value = Header(i - 9)
        Next i
        
        
        ' Print greatest % and ticker/value headers
        For j = 15 To 17
            ws.Cells(1, j).Value = Header(j - 8) 'Creates the headers " ,Ticker, Value"
            If j = 17 Then
                For i = 2 To 4
                    ws.Cells(i, 15).Value = Header(i + 2) ' Creates the row headers "Greatest % Increase, Greatest % Decrease, Greatest Total Volume"
                Next i
            End If
        Next j
    
    
        ' -------------------------------------------------------------
    
        For i = 1 To TotalRecords
            ' If the Ticker has changed on the table then put highest and lowest in the correct spots change the Ticker.
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value And ws.Cells(i, 1).Value <> ws.Cells(1, 1).Value Then
                
                Volume = Volume + ws.Cells(i, 7).Value 'Get the very last volume for the year
                ClosePrice = ws.Cells(i, 6).Value 'Take the close price from the current row (The last close of the year for the ticker)
        
                ' -------------------------------------------------------------
                'Place the values collected and the Ticker
                ws.Cells(PlacefromStart, 9).Value = Ticker 'Print Current Ticker
                ws.Cells(PlacefromStart, 10).Value = ClosePrice - OpenPrice 'Print change in value from the years 1st open price to last close price
                ws.Cells(PlacefromStart, 11).Value = (ClosePrice - OpenPrice) / (OpenPrice) 'Print Percentage change of the opening price of the year to the closing price of the year
                ws.Cells(PlacefromStart, 12).Value = Volume 'Print the total volume collected
                ' -------------------------------------------------------------
                ' Change the cell types so they accurately depict the values within them.
                ws.Cells(PlacefromStart, 10).NumberFormat = "$0.00" 'Switch these to the percent type
                ws.Cells(PlacefromStart, 11).NumberFormat = "0.00%" 'Switch these to the percent type
                ws.Cells(PlacefromStart, 12).NumberFormat = "0"     'Switch these to the integer type
                
                ' -------------------------------------------------------------
                ' If statements to check if the current highest and lowest percent change and highest total volume are smaller/larger than the new ones
                ' Tried to do this in a elseif statement but there may be instances where we may want to change both two amounds at the same time
                ' The elseif statement would have to have a third else that states if both criteria were met at the same time making the code difficult to parse
                ' Instead of one long ifelse statement it is easier to understand if I split it into three seperate if statements
                ' Replace the old ticker and the old score with the newly placed ones if they meet the criteria
                If ws.Cells(PlacefromStart, 11).Value > TopAmounts(0) Then
                    TopTickers(0) = Ticker
                    TopAmounts(0) = ws.Cells(PlacefromStart, 11).Value
                End If
                
                If ws.Cells(PlacefromStart, 11).Value < TopAmounts(1) Then
                    TopTickers(1) = Ticker
                    TopAmounts(1) = ws.Cells(PlacefromStart, 11).Value
                End If
                
                If ws.Cells(PlacefromStart, 12).Value > TopAmounts(2) Then
                    TopTickers(2) = Ticker
                    TopAmounts(2) = ws.Cells(PlacefromStart, 12).Value
                End If
                ' -------------------------------------------------------------
                ' If Else statement to change the colors of the yearly change column.
                If (ws.Cells(PlacefromStart, 10).Value <= 0) Then
                    ws.Cells(PlacefromStart, 10).Interior.Color = RGB(255, 0, 0)
                Else
                    ws.Cells(PlacefromStart, 10).Interior.Color = RGB(0, 255, 0)
                End If
                ' -------------------------------------------------------------
                
                Ticker = ws.Cells(i + 1, 1).Value 'Replace old ticker with the next one
                OpenPrice = ws.Cells(i + 1, 3).Value 'Set open price to the opening price of the new ticker current year
                Volume = 0 'Collect the starting volume for the new ticker
                
                ' -------------------------------------------------------------
                
                'Move where we place the Ticker info down one row so it is ready for the next ticker
                PlacefromStart = PlacefromStart + 1
            
            ElseIf ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value And Ticker = "" Then
                
                Ticker = ws.Cells(i + 1, 1).Value 'Replace old ticker with the next one
                OpenPrice = ws.Cells(i + 1, 3).Value 'Set open price to the opening price of the new ticker current year
                
            Else
                Volume = Volume + ws.Cells(i, 7).Value ' There are no other changes just collect the volume of that cell and add it to the total
            End If
        
        Next i
        
        ' Place the final highest and lowest percent change and highest total volume
        For i = 1 To 3
            ws.Cells(i + 1, 16).Value = TopTickers(i - 1)
            ws.Cells(i + 1, 17).Value = TopAmounts(i - 1)
        Next i
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        
    Next ' Move to the next worksheet

End Sub

