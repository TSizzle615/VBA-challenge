VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Module2_Challenge()

'define worksheet variable
Dim ws As Worksheet

'loop through worksheets
For Each ws In ActiveWorkbook.Worksheets
    With ws


        'define variables
        Dim i As Long
        Dim endRow As Long
        Dim Ticker_Count As Integer
        Dim Cumulative_Vol As Long
        Dim Current_Ticker As String
        Dim Qtrly_Open As Double
        Dim Max_Value As Long
        
        'name columns
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Columns("I:L").AutoFit
        
        'format spreadsheet
            ws.Columns("K").NumberFormat = "0.00%"
            ws.Range("Q2:Q3").NumberFormat = "0.00%"
            
        'count number of rows in worksheet (will need to reset when doing multiple sheets)
        endRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'define current ticker and Qrtly_Open
        Current_Ticker = ws.Cells(2, 1).Value
        ws.Cells(2, 9).Value = Current_Ticker
        
        'set values
        Ticker_Count = 1
        Qtrly_Open = ws.Cells(2, 3).Value
        Cumulative_Vol = 0
         
         'loop through data
        For i = 2 To endRow
        
            'identify change in ticker and adjust values
            If (ws.Cells(i, 1).Value) <> Current_Ticker Then
                Ticker_Count = Ticker_Count + 1
                ws.Cells(Ticker_Count, 9) = Current_Ticker
                ws.Cells(Ticker_Count, 10) = ws.Cells(i - 1, 6).Value - Qtrly_Open
                
                    'conditionally format qrtrly change
                    If (ws.Cells(Ticker_Count, 10).Value > 0) Then
                        ws.Cells(Ticker_Count, 10).Interior.ColorIndex = 4
                        ws.Cells(Ticker_Count, 10).Font.ColorIndex = 1
                    ElseIf (ws.Cells(Ticker_Count, 10).Value < 0) Then
                        ws.Cells(Ticker_Count, 10).Interior.ColorIndex = 3
                        ws.Cells(Ticker_Count, 10).Font.ColorIndex = 1
                    Else
                        ws.Cells(Ticker_Count, 10).Interior.ColorIndex = 2
                        ws.Cells(Ticker_Count, 10).Font.ColorIndex = 1
                        
                    End If
                    
                
                'calculate percent change and total volume
                ws.Cells(Ticker_Count, 11) = ws.Cells(Ticker_Count, 10) / Qtrly_Open
                ws.Cells(Ticker_Count, 12).Value = Application.SumIf(ws.Range("A:A"), Current_Ticker, ws.Range("G:G"))
                Current_Ticker = ws.Cells(i, 1).Value
                Qtrly_Open = ws.Cells(i, 3).Value
          
                
        
            End If
            
        Next i
        
        'create and populate summation table
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Range("O2:O4").Columns.AutoFit
        ws.Range("P1:P4").Columns.AutoFit
        ws.Range("Q1:Q2").Columns.AutoFit
        
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        ws.Range("Q2") = Application.Max(ws.Range("K:K"))
        ws.Range("Q3") = Application.Min(ws.Range("K:K"))
        ws.Range("Q4") = Application.Max(ws.Range("L:L"))
        
        ws.Columns("P:R").AutoFit
         ws.Range("P2") = Application.Index(ws.Range("I:I"), Application.Match(ws.Range("Q2"), ws.Range("K:K"), 0))
         ws.Range("P3") = Application.Index(ws.Range("I:I"), Application.Match(ws.Range("Q3"), ws.Range("K:K"), 0))
         ws.Range("P4") = Application.Index(ws.Range("I:I"), Application.Match(ws.Range("Q4"), ws.Range("L:L"), 0))
         
  End With
 
 Next ws

End Sub
