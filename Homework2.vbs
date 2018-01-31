Sub Stock_Volume()
    Dim WS As Worksheet
    Set WS = Worksheets("2014")
    Dim Ticker As String
    Ticker = 0
    Dim TotalStockVolume As Integer
    TotalStockVolume = 2
    Dim lastRow As Integer
    lastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
        Print (lastRow)
    Ticker = WS.Cells(2, 1).Value
    For I = 2 To lastRow
    TotalStockVolume = TotalStockVolume + 1
    WS.Range ("A1:B2")
    
End Sub
Sub Test()
    Dim WS As Worksheet
    Set WS = Worksheets("2014")
    Dim RowI As Integer
    RowI = 2
    
    Debug.Print (WS.Cells(2, 1))
    
    'WS.Range ("G2")
    Debug.Print (WS.Range("G2"))
    Dim Total As Double
    Total = 0
    
    For RowI = 2 To 5
    
        Debug.Print (WS.Cells(RowI, 1))
        Debug.Print (WS.Cells(RowI, 7))
        Total = Total + CDbl(WS.Cells(RowI, 7).Value)
        Debug.Print (Total)
        
        
        
    Next RowI
    
    
    
End Sub

Sub WorksheetLoop()

         Dim WS_Count As Integer
         Dim I As Integer
         
         WS_Count = ActiveWorkbook.Worksheets.Count
         
         For I = 1 To WS_Count

            MsgBox ActiveWorkbook.Worksheets(I).Name

         Next I

      End Sub
