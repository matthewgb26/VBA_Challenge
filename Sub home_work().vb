Sub home_work()

 Dim Ticker_name As String

 Dim summary_table As Integer

 Dim yearly_chnage As Integer

 Dim opening_price As Integer

 Dim closing_price As Double

 Dim percent_change As Double

 Dim total_volume As Double 

 Dim greatest_total_volume As Double

 Dim greatest_total_volume_ticker As String


  For Each ws In Worksheets

   total_volume = 0

  
   greatest_total_volume = 0
   greatest_total_volume_ticker = ""
  
   greatest_percent_decrease = 999999999
   greatest_percent_decrease_ticker = ""
  
   greatest_percent_increase = 0
   greatest_percent_increase_ticker = ""
  
   lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
   opening_price = ws.Cells(2, 3).Value
  
   summary_table = 2

   ws.Cells(1, 11).Value = "Ticker"
   ws.Cells(1, 12).Value = "Yearly Change"
   ws.Cells(1, 13).Value = "Percent Change"
   ws.Cells(1, 14).Value = "Total Stock Volume"
   ws.Cells(1, 18).Value = "Ticker"
   ws.Cells(1, 19).Value = "Value"
   ws.Cells(2, 17).Value = "greatest % increase"
   ws.Cells(3, 17).Value = "greatest % decrease"
   ws.Cells(4, 17).Value = "greatest total Volume"


   For i = 2 To lastrow

     total_volume = total_volume + ws.Cells(i, 7).Value


     If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then

       closing_price = ws.Cells(i, 6).Value

    
       yearly_change = closing_price - opening_price
    
       If opening_price > 0 Then
       percent_change = (closing_price - opening_price) / opening_price
       Else
       percent_change = 0
       End If
    
    

       Ticker_name = ws.Cells(i, 1).Value
    
       opening_price = ws.Cells(i + 1, 3).Value

       ws.Cells(summary_table, 11).Value = Ticker_name
       ws.Cells(summary_table, 12).Value = yearly_change
       ws.Cells(summary_table, 14).Value = total_volume
       ws.Cells(summary_table, 13).Value = percent_change
    
       If total_volume > greatest_total_volume Then
       greatest_total_volume = total_volume
       greatest_total_volume_ticker = ws.Cells(i, 1).Value
       End If
    
    
    
       If percent_change < greatest_percent_decrease Then
       greatest_percent_decrease = percent_change
       greatest_percent_decrease_ticker = ws.Cells(i, 1).Value
       End If
    
    
       If percent_change > greatest_percent_increase Then
       greatest_percent_increase = percent_change
       greatest_percent_increase_ticker = ws.Cells(i, 1).Value
       End If



       Totale_volume = 0

       summary_table = summary_table + 1



      End If
    
    Next i
   ws.Cells(4, 19).Value = greatest_total_volume
   ws.Cells(4, 18).Value = greatest_total_volume_ticker
   ws.Cells(3, 19).Value = greatest_percent_decrease
   ws.Cells(3, 18).Value = greatest_total_volume_ticker
   ws.Cells(2, 19).Value = greatest_percent_increase
   ws.Cells(2, 18).Value = greatest_percent_increase_ticker

  Next ws


End Sub