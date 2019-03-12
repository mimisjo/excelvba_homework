Attribute VB_Name = "Module1"
Sub WallStreet_Moderate()

    ' Set a variable for holdng the ticker
    Dim Ticker As String
    
    ' Set an initial variable for the volume totals
    Dim Volume_Total As Double
    Volume_Total = 0
    
    'Set variables for beginning and ending stock prices
    Dim Beg_Price As Double
    Dim End_Price As Double
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    
    ' Keep track of the ticker location in the summary table
    Dim Summary_Table_Row As Double
    Summary_Table_Row = 2
    
    ' Set location of summary table on sheet A
    ' Set wsa = Worksheets("A")
    
    ' Create summary table labels
    Cells(1, 9) = "Ticker"
    Cells(1, 10) = "Yearly Change"
    Cells(1, 11) = "Percent Change"
    Cells(1, 12) = "Total Stock Volume"

    ' Find the last row of each worksheet
    Dim Last_Row As Double
    Last_Row = Cells(Rows.Count, 1).End(xlUp).Row
                     
            'Loop through all stock ticker volumes
            For i = 2 To Last_Row
            
                ' Set beginning price
                If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
                    Beg_Price = Cells(i, 3).Value
                    
                    ' Check: Print beginning price in summary table
                    ' Range("M" & Summary_Table_Row).Value = Beg_Price
                
                ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                    'Set the Ticker
                    Ticker = Cells(i, 1).Value
                    
                    'Set the end price
                    End_Price = Cells(i, 6).Value
                    
                    ' Compute yearly change and percent change
                    Yearly_Change = End_Price - Beg_Price
                    Percent_Change = Yearly_Change / Beg_Price
    
                    ' Check: Print end price in summary table
                    ' Range("N" & Summary_Table_Row).Value = End_Price
                    
                    ' Add to the volume total
                    Volume_Total = Volume_Total + Cells(i, 7)

                    ' Print the stock ticker in the summary table
                    Range("I" & Summary_Table_Row).Value = Ticker
                    
                    ' Print the yearly change in the summary table
                    Range("J" & Summary_Table_Row).Value = Yearly_Change
                    
                    ' Print the percent change in the summary table
                    Range("K" & Summary_Table_Row).Value = Percent_Change
                    
                    ' Print the volume total in the summary table
                    Range("L" & Summary_Table_Row).Value = Volume_Total

                    ' Add one to the summary table row
                    Summary_Table_Row = Summary_Table_Row + 1
      
                    ' Reset the volume total
                    Volume_Total = 0
                    
                ' If the cell immediately following a row is the same ticker...
                Else
                
                    ' Add to the volume total
                    Volume_Total = Volume_Total + Cells(i, 7).Value
                
                End If
                
            Next i
    
    ' Format summary table
        ' Set last row for summary table
        Dim Last_Row_Table As Double
        Last_Row_Table = Cells(Rows.Count, 9).End(xlUp).Row
        
        ' Autofit to display data
        Columns("I:L").AutoFit
        
        ' Color code yearly change column
        For i = 2 To Last_Row_Table
            If Cells(i, 10).Value > 0 Then
            
                ' Color the cells green
                Cells(i, 10).Interior.ColorIndex = 4
                
            Else
                ' Color the cells red
                Cells(i, 10).Interior.ColorIndex = 3
                
            End If
            
        Next i
        
        ' Format percent and volume total columns
        For i = 2 To Last_Row_Table
            ' Format percent column
            Cells(i, 11).NumberFormat = "0.00%"
            
            ' Format volume total column
            Cells(i, 12).NumberFormat = "0,000"
            
        Next i
         
End Sub

