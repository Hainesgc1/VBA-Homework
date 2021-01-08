Sub ticker()

        ' Set ws as a lead worksheet
        Dim ws As Worksheet
        Dim Need_Summary_Table_Header
        Dim Main_Space
        
        Need_Summary_Table_Header = True
        Main_Space = True
        
        ' Loop through all of the worksheets in the active workbook.
        For Each ws In Worksheets
        
            ' Set initial variable for holding the ticker name
            Dim Ticker_Name As String
            Ticker_Name = " "
            Dim Total_Ticker_Volume As Double
            Total_Ticker_Volume = 0
            
            ' Set all variables
            Dim Open_Price As Double
            Open_Price = 0
            Dim Close_Price As Double
            Close_Price = 0
            Dim Change_Price As Double
            Change_Price = 0
            Dim Change_Percent As Double
            Change_Percent = 0
            Dim High_Ticker_Name As String
            High_Ticker_Name = " "
            Dim Low_Ticker_Name As String
            Low_Ticker_Name = " "
            Dim High_Percent As Double
            High_Percent = 0
            Dim Low_Percent As Double
            Low_Percent = 0
            Dim Max_Volume_Ticker As String
            Max_Volume_Ticker = " "
            Dim Max_Volume As Double
            Max_Volume = 0
            
            Dim Summary_Table_Row As Long
            Summary_Table_Row = 2
            
            ' Set initial row count for the active worksheet
            Dim Lastrow As Long
            Dim i As Long
            
            Lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row

            ' Set headers for all worksheets
            If Need_Summary_Table_Header Then
                ' Set Titles for the Summary Table for active worksheet
                ws.Range("I1").Value = "Ticker"
                ws.Range("J1").Value = "Yearly Change"
                ws.Range("K1").Value = "Percent Change"
                ws.Range("L1").Value = "Total Stock Volume"
                ws.Range("O2").Value = "Greatest % Increase"
                ws.Range("O3").Value = "Greatest % Decrease"
                ws.Range("O4").Value = "Greatest Total Volume"
                ws.Range("P1").Value = "Ticker"
                ws.Range("Q1").Value = "Value"
            Else
                Need_Summary_Table_Header = True
            End If
            
            ' Set value of Open Price for Ticker of ws
            Open_Price = ws.Cells(2, 3).Value
            
            ' Loop from the beginning of the active worksheet (Row2) until last row
            For i = 2 To Lastrow

                ' Monitor the information in the ticker
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                    ' Set ticker name
                    Ticker_Name = ws.Cells(i, 1).Value
                    ' Calculate Change_Price and Change_Percent
                    Close_Price = ws.Cells(i, 6).Value
                    Change_Price = Close_Price - Open_Price
                    If Open_Price <> 0 Then
                        Change_Percent = (Change_Price / Open_Price) * 100
                    End If
                    
                    ' Add to the Ticker Total Volume
                    Total_Ticker_Volume = Total_Ticker_Volume + ws.Cells(i, 7).Value
                  
                    ' Print the Ticker Name in the Summary Table, Column I
                    ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
                    ws.Range("J" & Summary_Table_Row).Value = Change_Price
                    ' Color code with green or red depending on the Yearly Change
                    If (Change_Price > 0) Then
                        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                    ElseIf (Change_Price <= 0) Then
                        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                    End If
                    
                     ' Add Percent Change and Ticker Volume Total to Column I and Column J
                    ws.Range("K" & Summary_Table_Row).Value = (CStr(Change_Percent) & "%")
                    ws.Range("L" & Summary_Table_Row).Value = Total_Ticker_Volume
                    Summary_Table_Row = Summary_Table_Row + 1
                    Change_Price = 0
                    Close_Price = 0
                    
                    ' Add next ticker's Open_Price
                    Open_Price = ws.Cells(i + 1, 3).Value

                    If (Change_Percent > High_Percent) Then
                        High_Percent = Change_Percent
                        High_Ticker_Name = Ticker_Name
                    ElseIf (Change_Percent < Low_Percent) Then
                        Low_Percent = Change_Percent
                        Low_Ticker_Name = Ticker_Name
                    End If
                           
                    If (Total_Ticker_Volume > Max_Volume) Then
                        Max_Volume = Total_Ticker_Volume
                        Max_Volume_Ticker = Ticker_Name
                    End If

                    Change_Percent = 0
                    Total_Ticker_Volume = 0

                Else
                    Total_Ticker_Volume = Total_Ticker_Volume + ws.Cells(i, 7).Value
                End If
                'Msgbox (ws.Rows(i).Cells(2,1))
          
            Next i

                If Main_Space Then
                
                    ws.Range("Q2").Value = (CStr(High_Percent) & "%")
                    ws.Range("Q3").Value = (CStr(Low_Percent) & "%")
                    ws.Range("P2").Value = High_Ticker_Name
                    ws.Range("P3").Value = Low_Ticker_Name
                    ws.Range("Q4").Value = Max_Volume
                    ws.Range("P4").Value = Max_Volume_Ticker
                    
                Else
                    Main_Space = True
                End If
            
         Next ws
End Sub
