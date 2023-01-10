Sub Stock_Analysis()

For Each Sheet In Sheets
Sheet.Select

' Set an initial variables for the ticker, yearly_change, percent_change, total_vol
Dim ticker As String

Dim yearly_change As Double
    yearly_change = 0
    
Dim percent_change As Double
    percent_change = 0
    
Dim total_vol As Double
    total_vol = 0

Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2

' Set variables for open and close price used for yearly_change
Dim open_price As Double
Dim close_price As Double

' Loop all worksheets (NEED TO FIX)


    'Last row
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    Range("L1") = "Ticker"
    Range("M1") = "Yearly Change"
    Range("N1") = "Percent Change"
    Range("O1") = "Total Stock Volume"

    'Loop for summary table
    For i = 2 To lastrow

        'If statements to determine open and close price for the year
        If Right(Cells(i, 2), 4) = "0102" Then

            open_price = Cells(i, 3)

        End If

        If Right(Cells(i, 2), 4) = "1231" Then

            close_price = Cells(i, 6)

        End If

        'Check if you are still within same ticker
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            'Set ticker column
            ticker = Cells(i, 1).Value
            
            'Set yearly_change
            yearly_change = close_price - open_price
            
            'Set percent_change
            percent_change = yearly_change / open_price
            
            ' Add to the total_vol
            total_vol = total_vol + Cells(i, 7).Value
            
            ' Print the ticker in the Summary Table
            Range("L" & Summary_Table_Row).Value = ticker
            
            ' Print the yearly_change in the Summary Table
            Range("M" & Summary_Table_Row).Value = yearly_change
            Range("M" & Summary_Table_Row).NumberFormat = "#,##0.0"
            
            'Color Formatting
            If yearly_change < 0 Then

               Range("M" & Summary_Table_Row).Interior.ColorIndex = 3

                Else

                    Range("M" & Summary_Table_Row).Interior.ColorIndex = 4

            End If
            
            ' Print the yearly_change in the Summary Table
            Range("N" & Summary_Table_Row).Value = percent_change
            Range("N" & Summary_Table_Row).NumberFormat = "0.00%"

            'Print the total_vol in the Summary Table
            Range("O" & Summary_Table_Row).Value = total_vol
            
            ' Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
            
            ' Reset the total_vol
            total_vol = 0
            yearly_change = 0
        
        ' If next cell is same ticker
         Else

            ' Add to the total_vol
            total_vol = total_vol + Cells(i, 7).Value
        
        End If
        
    Next i

Next

End Sub