Sub Multiple_year_stock_Data()

Dim ws As Worksheet
'Activating each worksheet for looping
For Each ws In Worksheets

ws.Activate

'Declaring all variables

Dim i As Long
Dim ticker As String
Dim OpenValue As Double
Dim CloseValue As Double
Dim YearlyChange As Double
Dim PerChange As Double


Dim total_stock As Double
total_stock = 0

'Declaring variable for summary table

Dim summary_table_row As Long
summary_table_row = 2

Dim p As Long
p = 2

Dim lrow As Long
lrow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row

'Creating header values

Cells(1, 10).Value = "Ticker"
Cells(1, 13).Value = "Total Stock Volume"
Cells(1, 11).Value = "Yearly Change"
Cells(1, 12).Value = "Percent Change"
Cells(2, 16).Value = "Greatest % Increase"
Cells(3, 16).Value = "Greatest % Decrease"
Cells(4, 16).Value = "Greatest total Volume"
Cells(1, 17).Value = "Ticker"
Cells(1, 18) = "Value"

'Performing for loop to find ticker and related data.

For i = 2 To lrow

If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

'To find yearly change values

    ticker = Cells(i, 1).Value
    total_stock = total_stock + Cells(i, 7).Value
            OpenValue = Cells(p, 3).Value
            CloseValue = Cells(i, 6).Value
            YearlyChange = CloseValue - OpenValue

       
             'Excluding division by 0 error

             If OpenValue <> 0 Then
           '

                PerChange = YearlyChange / OpenValue

            'Round off function

                Range("L:L").NumberFormat = "0.00%"
            
            Else
            
            End If
            
           
            'Naming summary table column

                Range("J" & summary_table_row).Value = ticker
                Range("M" & summary_table_row).Value = total_stock
                Range("K" & summary_table_row).Value = YearlyChange
                Range("L" & summary_table_row).Value = PerChange

                'Adding color to cells
                
                    If Range("K" & summary_table_row).Value >= 0 Then
                        Range("K" & summary_table_row).Interior.ColorIndex = 4
                    Else
                        Range("K" & summary_table_row).Interior.ColorIndex = 3
                    End If
                
                
                summary_table_row = summary_table_row + 1
                

    total_stock = 0
    p = i + 1
   

Else

    total_stock = total_stock + Cells(i, 7).Value

End If

Next i

'Declaring variables and finding min and max values

Dim Gin As Double
Dim Gde As Double
Dim Gtov As Double

Gin = Application.WorksheetFunction.Max(Range("L:L"))
    Cells(2, 18).Value = Gin
    Cells(2, 17).Value = Cells(Application.WorksheetFunction.Match(Range("R2"), Range("l:l"), 0), 10)
    Range("R2").NumberFormat = "0.00%"

Gde = Application.WorksheetFunction.Min(Range("L:L"))
    Cells(3, 18).Value = Gde
    Cells(3, 17).Value = Cells(Application.WorksheetFunction.Match(Range("R3"), Range("l:l"), 0), 10)
    Range("R3").NumberFormat = "0.00%"
    
Gtov = Application.WorksheetFunction.Max(Range("M:M"))
    Cells(4, 18).Value = Gtov
    Cells(4, 17).Value = Cells(Application.WorksheetFunction.Match(Range("R4"), Range("M:M"), 0), 10)
    
Next ws

End Sub
