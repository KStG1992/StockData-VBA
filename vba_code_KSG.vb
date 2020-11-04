'This subroutine will loop through all the stocks on any active worksheet and output the Ticker Symbol, Yearly Change, Percent Change, and Volume'

Sub wall_street()

'=============================================================================='
    'Declaring Variables'
    Dim ws As Worksheet
    
'=============================================================================='
    'Looping Through the Worksheets'
    For Each ws In Worksheets
    
'=============================================================================='
        'Declaring Variables'
        Dim col As Long
        Dim yearly_change As Double
        Dim percent_change As Double
        Dim length_ticker As Double
        Dim i As Double
        Dim start_ticker As Double
        Dim end_ticker As Double
        Dim table As String
        Dim vol_total As Double
'=============================================================================='
        'Setting Counter to Zero'
        length_ticker = 0
        'Keeping Track of Location for Summary Table'
        table = 2
        'Setting vol_total to Zero'
        vol_total = 0
'=============================================================================='
        'Creating Headers'
        ws.Range("J1").Value = "Ticker"
        ws.Range("K1").Value = "Yearly Change"
        ws.Range("L1").Value = "Percent Change"
        ws.Range("M1").Value = "Total Volume"
'=============================================================================='
    
        'Finding The Length of Column of Entire Dataset'
        col = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
        'Looping Through Stocks'
        For i = 2 To col
            
            'Finding The Length of Each Ticker'
            If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
            
                length_ticker = length_ticker + 1
                
                vol_total = vol_total + ws.Cells(i, 7).Value
                
            Else
                
                'Finding the Value of the Opening and Closing Stock Value'
                start_ticker = ws.Cells(i - length_ticker, 3).Value
                end_ticker = ws.Cells(i, 6).Value
                
                'Outputting Ticker Symbol in Column J'
                ws.Range("J" & table) = ws.Cells(i, 1).Value
                
                'Calculating Yearly Change'
                yearly_change = end_ticker - start_ticker
                ws.Range("K" & table).Value = yearly_change
                 
                'Determining the Color of the Yearly Change'
                If yearly_change > 0 Then
                    ws.Range("K" & table).Interior.ColorIndex = 4
                    
                Else
                    
                    ws.Range("K" & table).Interior.ColorIndex = 3
                
                End If
                    
                'Calculating Percent Change'
                percent_change = Round((yearly_change / start_ticker) * 100, 2)
                ws.Range("L" & table).Value = Str(percent_change) + "%"
                
                'Calculating the Total Volume of the Stock'
                vol_total = vol_total + ws.Cells(i, 7).Value
                ws.Range("M" & table).Value = vol_total
                
                'Adjusting the Location of the Summary Table'
                table = table + 1
                
                'Resetting Volume Total'
                vol_total = 0
                
            End If
        
        Next i
        
    Next

'=============================================================================='
End Sub

