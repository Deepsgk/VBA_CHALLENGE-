Attribute VB_Name = "Module1"
Sub Ticker_over_the_years()

    Dim Ws As Worksheet
'Using boolean operator with "if" command to set up the logic that can run across years/alphabets
    Dim Summary_Table_Header As Boolean
    Summary_Table_Header = True
    Dim active_spreadsheet As Boolean
    active_spreadsheet = True
    
     For Each Ws In Worksheets
        'Variables for question 1
        Dim Ticker_Name As String
        Dim Total_Ticker_Volume As Double
        Total_Ticker_Volume = 0
        Dim Open_Price As Double
        Open_Price = 0
        Dim Close_Price As Double
        Close_Price = 0
        Dim Price_change As Double
        Price_change = 0
        Dim Percent_change As Double
        Percent_change = 0
        
       'Variables for question 2
        Dim Max_value_ticker_name As String
        Dim Min_value_ticker_name As String
        Dim Increase_percentage As Double
        Increase_percentage = 0
        Dim Decrease_percentage As Double
        Decrease_percentage = 0
        Dim Maximum_volume_ticker As String
        Dim Maximum_volume As Double
        Maximum_volume = 0
        Dim Summary_Table_Row As Long
        Summary_Table_Row = 2
        
        Dim finalrow As Long
        Dim i As Long
        finalrow = Ws.Cells(Rows.Count, 1).End(xlUp).Row

        If Summary_Table_Header Then
            ' Set Titles for the question 1
            Ws.Range("I1").Value = "Ticker"
            Ws.Range("J1").Value = "Yearly Change"
            Ws.Range("K1").Value = "Percent Change"
            Ws.Range("L1").Value = "Total Stock Volume"
            ' Set Titles for the quetsion 2
            Ws.Range("O2").Value = "Greatest % Increase"
            Ws.Range("O3").Value = "Greatest % Decrease"
            Ws.Range("O4").Value = "Greatest Total Volume"
            Ws.Range("P1").Value = "Ticker"
            Ws.Range("Q1").Value = "Value"
        Else
            Summary_Table_Header = False
        End If
        
        Open_Price = Ws.Cells(2, 3).Value
        
        For i = 2 To finalrow
        
            If Ws.Cells(i + 1, 1).Value <> Ws.Cells(i, 1).Value Then
            
                Ticker_Name = Ws.Cells(i, 1).Value
                Close_Price = Ws.Cells(i, 6).Value
                Price_change = Close_Price - Open_Price
                
                If Open_Price <> 0 Then
                    Percent_change = (Price_change / Open_Price) * 100
                Else
                    On Error Resume Next
                End If
                Total_Ticker_Volume = Total_Ticker_Volume + Ws.Cells(i, 7).Value
              
            Ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
            Ws.Range("J" & Summary_Table_Row).Value = Price_change
                'Positive change
                If (Price_change > 0) Then
                    Ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                 'Negative change
                ElseIf (Price_change <= 0) Then
                    Ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                End If
                Ws.Range("K" & Summary_Table_Row).Value = (CStr(Percent_change) & "%")
                Ws.Range("L" & Summary_Table_Row).Value = Total_Ticker_Volume
                
                Summary_Table_Row = Summary_Table_Row + 1
                
                Price_change = 0
                Close_Price = 0
                Open_Price = Ws.Cells(i + 1, 3).Value

                If (Percent_change > Increase_percentage) Then
                    Increase_percentage = Percent_change
                    Max_value_ticker_name = Ticker_Name
                ElseIf (Percent_change < Decrease_percentage) Then
                    Decrease_percentage = Percent_change
                     Min_value_ticker_name = Ticker_Name
                End If
                       
                If (Total_Ticker_Volume > Maximum_volume) Then
                    Maximum_volume = Total_Ticker_Volume
                    Maximum_volume_ticker = Ticker_Name
                End If
                
                Percent_change = 0
                Total_Ticker_Volume = 0
            Else
                
                Total_Ticker_Volume = Total_Ticker_Volume + Ws.Cells(i, 7).Value
            End If
        
        Next i

            If active_spreadsheet Then
            
                Ws.Range("Q2").Value = (CStr(Increase_percentage) & "%")
                Ws.Range("Q3").Value = (CStr(Decrease_percentage) & "%")
                Ws.Range("P2").Value = Max_value_ticker_name
                Ws.Range("P3").Value = Min_value_ticker_name
                Ws.Range("Q4").Value = Maximum_volume
                Ws.Range("P4").Value = Maximum_volume_ticker
            Else
            active_spreadsheet = False
            End If
     Next Ws
     End Sub


