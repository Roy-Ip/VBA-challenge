Sub stock()

    Dim ws As Worksheet
    Dim Ticker As String
    Dim Total_Stock_Volume As Double
    Dim Yearly_Change As Double
    Dim Opening_Price As Double
    Dim Closing_Price As Double
    Dim Percent_Change As Double
    Dim Summary_Row As Integer
    Dim Greatest_Increase As Double
    Dim Greatest_Decrease As Double
    Dim Greatest_Volume As Double
    Dim Ticker_Greatest_Increase As String
    Dim Ticker_Greatest_Decrease As String
    Dim Ticker_Greatest_Volume As String
    
    For Each ws In Worksheets
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        Dim LastRow As Long
        LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        Total_Stock_Volume = 0
        Yearly_Change = 0
        Opening_Price = 0
        Closing_Price = 0
        Summary_Row = 2
        Greatest_Increase = 0
        Greatest_Decrease = 0
        Greatest_Volume = 0
        Ticker_Greatest_Increase = ""
        Ticker_Greatest_Decrease = ""
        Ticker_Greatest_Volume = ""
        
        For i = 2 To LastRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Ticker = ws.Cells(i, 1).Value
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
                
                
                Closing_Price = ws.Cells(i, 6).Value
                If Opening_Price <> 0 Then
                    Yearly_Change = Closing_Price - Opening_Price
                    ws.Range("J" & Summary_Row).Value = Yearly_Change
                    
                        If Yearly_Change >= 0 Then
                        ws.Range("J" & Summary_Row).Interior.ColorIndex = 4
                        Else
                        ws.Range("J" & Summary_Row).Interior.ColorIndex = 3
                        End If

                    If Opening_Price <> 0 Then
                        Percent_Change = Yearly_Change / Opening_Price
                        ws.Range("K" & Summary_Row).Value = Percent_Change
                        ws.Range("K" & Summary_Row).NumberFormat = "0.00%"
                        
                        If Percent_Change >= 0 Then
                        ws.Range("K" & Summary_Row).Interior.ColorIndex = 4
                        Else
                        ws.Range("K" & Summary_Row).Interior.ColorIndex = 3
                        End If
                        
                            If Percent_Change > Greatest_Increase Then
                                Greatest_Increase = Percent_Change
                                Ticker_Greatest_Increase = Ticker
                            ElseIf Percent_Change < Greatest_Decrease Then
                                Greatest_Decrease = Percent_Change
                                Ticker_Greatest_Decrease = Ticker
                            End If
                            
                            If Total_Stock_Volume > Greatest_Volume Then
                            Greatest_Volume = Total_Stock_Volume
                            Ticker_Greatest_Volume = Ticker
                            End If
                        
                    Else
                        ws.Range("K" & Summary_Row).Value = 0
                    End If
                    
                End If
                
                ws.Range("I" & Summary_Row).Value = Ticker
                ws.Range("L" & Summary_Row).Value = Total_Stock_Volume
                
                Summary_Row = Summary_Row + 1
                Total_Stock_Volume = 0
                Yearly_Change = 0
                Opening_Price = 0
                Closing_Price = 0
                
            Else
            
                If Opening_Price = 0 Then
                    Opening_Price = ws.Cells(i, 3).Value
                End If
                
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
                
            End If
            
        Next i
        
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        ws.Range("P2").Value = Ticker_Greatest_Increase
        ws.Range("P3").Value = Ticker_Greatest_Decrease
        ws.Range("P4").Value = Ticker_Greatest_Volume
        
        ws.Range("Q2").Value = Greatest_Increase
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").Value = Greatest_Decrease
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("Q4").Value = Greatest_Volume
        
        ws.Columns("I:Q").AutoFit
        
    Next ws
    
End Sub





