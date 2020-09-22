Sub Stock_Calculator()
    For Each ws In Worksheets
        ws.Activate
        Call Set_Title
    Next ws

End Sub
Sub Set_Title()

    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
Call calc_vol
    
End Sub
Sub calc_vol()

Dim current_ticker As String
Dim next_ticker As String
Dim total_rows As Long
Dim total_cols As Integer
Dim total_volume As Long
Dim summary_row As Integer

total_rows = Cells(Rows.Count, "A").End(xlUp).Row

summary_row = 2


For current_row = 2 To total_rows
    
    current_ticker = Cells(current_row, 1).Value
    
    next_ticker = Cells(current_row + 1, 1).Value
    
    Total = Total + Cells(current_row, "G").Value
    
    If current_ticker <> next_ticker Then
        
        Cells(summary_row, "I").Value = current_ticker
        Cells(summary_row, "L").Value = Total
        
        summary_row = summary_row + 1
        
        Total = 0
        
    End If
    
Next current_row
    
Call calc_year_change
    
End Sub
Sub calc_year_change()

Dim current_ticker As String
Dim next_ticker As String
Dim cur_close_ticker As String
Dim nxt_close_ticker As String
Dim total_rows As Long
Dim total_cols As Integer
Dim total_volume As Long
Dim summary_row As Integer
Dim yearly_change As Double
Dim Percent_change As Double
Dim current_open As Double
Dim open_vol As Double
Dim close_vol As Double
Dim ticker_date As String
Dim exit_loop As Boolean

total_rows = Cells(Rows.Count, "A").End(xlUp).Row

summary_row = 2

For current_row = 2 To total_rows
    
    current_ticker = Cells(current_row, 1).Value
    next_ticker = Cells(current_row + 1, 1).Value
    x = current_row
    
        For cur_vol_row = x To total_rows
           current_open = Cells(cur_vol_row, "C").Value
        
        If current_open = 0 Then
            exit_loop = True
            Exit For

        End If
            
        If open_vol And close_vol > 0 Then
            exit_loop = True
            Exit For
    
        End If
    
        If current_open > 0 Then
            open_vol = Cells(cur_vol_row, "C").Value
            
        End If
         
                For close_vol_row = x + 1 To total_rows
                
                    cur_close_ticker = Cells(close_vol_row, 1).Value
                    nxt_close_ticker = Cells(close_vol_row + 1, 1).Value
                    
                    If cur_close_ticker <> nxt_close_ticker Then
                    close_vol = Cells(close_vol_row, "F").Value
                    
                End If
                
                    If close_vol > 0 Then
                        exit_loop = True
                        Exit For
                    End If
                
                Next close_vol_row
                
        Next cur_vol_row
        
If open_vol = 0 Then

    Percent_change = 0
    open_vol = 1
    close_vol = 1

End If

If current_ticker <> next_ticker Then
        
    yearly_change = close_vol - open_vol
    Percent_change = yearly_change / open_vol
   
        'Cells(summary_row, "M").Value = open_vol
        'Cells(summary_row, "N").Value = close_vol
        Cells(summary_row, "J").Value = yearly_change
        Cells(summary_row, "K").Value = Percent_change
        
        summary_row = summary_row + 1
        cur_vol_row = cur_vol_row + 1
        
        open_vol = 0
        close_vol = 0
        
   
End If
       
If exitLoop = True Then
        Exit For
        End If
       
Next current_row

Call cond_format

End Sub

Sub cond_format()

total_rows = Cells(Rows.Count, "K").End(xlUp).Row

For i = 2 To total_rows

  If Cells(i, "J").Value < 0 Then

      Cells(i, "J").Interior.ColorIndex = 3
      
  ElseIf Cells(i, "J").Value > 0 Then

      Cells(i, "J").Interior.ColorIndex = 43
      
  ElseIf Cells(i, "J").Value = 0 Then

      Cells(i, "J").Interior.ColorIndex = 15

  End If

Next i

Range("K:K").NumberFormat = "0.00%"
Range("I:O").Columns.AutoFit

End Sub

