Sub Stonks_1()
'----------------------------------------------------------------------------------
'Acknowlegments and referencing
'I have not used, nor borrowed any code. Nor have I collaborated with any classmates. 
'I have referred for guidance in developing my own script to class materials provided
'I would like to acknowledge and thank fellow student Tracey Ha and Instructor Ryan Collingwood for their
'responses provided to a specific technical question, regarding the use of the Long data type.
'-----------------------------------------------------------------------------------
'Declare Variables
Dim Ticker As String
Dim Volume As Double
Dim trade_date As Date
Dim Open_Price As Double
Dim High_Price As Double
Dim Low_Price As Double
Dim Close_Price As Double
Dim last_row As Long
Dim row_number As Long
Dim lastday As Integer
Dim i As Long




Volume = 0
row_number = 2
last_row = Cells(Rows.Count, 1).End(xlUp).Row
day_counter = 0

'---------------------------------------------------------

'Set out results table

Range("M1:P1").Font.Bold = True

Cells(1, 13).Value = "Ticker Symbol"
Cells(1, 14).Value = "Annual $ Change"
Cells(1, 15).Value = "Annual % Change"
Cells(1, 16).Value = "Total Volume Traded"

Range("M1:P1").Columns.AutoFit

'----------------------------------------------------------------------


'Begin Loop to identify stock tickers

For i = 2 To CLng(last_row)
If Cells(i, 1).Value = Cells(i + 1, 1).Value Then

Ticker = Cells(i, 1).Value
Volume = Volume + Cells(i, 7).Value
Cells(row_number, 13).Value = Ticker

'Find opening stock price on first date for each ticker

        If day_counter = 0 Then
            Open_Price = Cells(i, 3).Value
        End If

'update day counter
day_counter = day_counter + 1


Else

Volume = Volume + Cells(i, 7).Value


Cells(row_number, 16).Value = Volume

'Find close price at end of day
Close_Price = Cells(i, 6).Value

'Calculate % change, with error handling for anomolous stock results
If Open_Price = 0 Then
Cells(row_number, 15).Value = "N/A"
Else: Cells(row_number, 15).Value = (Close_Price - Open_Price) / Open_Price
End If

Cells(row_number, 14).Value = Close_Price - Open_Price


'---------------------------------------------------------------------------------

row_number = row_number + 1

'Reset Volume to 0
Volume = 0
day_counter = 0

End If
 Next i
'------------------------------------------------------------------------------------------
 
'Format results table columns
 Range("O:O").NumberFormat = "0.00%"
 Range("N:N").NumberFormat = "$#,##0.00"
 Range("P:P").NumberFormat = "#,##0"
    
'Do a row count of results table, for use in conditional formatting looping, where number of rows not fixed

Dim last_row_results As Long
last_row_results = Cells(Rows.Count, 13).End(xlUp).Row

  For i = 2 To last_row_results
    If Cells(i, 14).Value > 0 Then
      Cells(i, 14).Interior.ColorIndex = 4
    ElseIf Cells(i, 14).Value < 0 Then
      Cells(i, 14).Interior.ColorIndex = 3
    End If
    Next i
'-------------------------------------------------------------------------------------------

End Sub
