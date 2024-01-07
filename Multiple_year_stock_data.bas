Attribute VB_Name = "Module1"
Sub LoopDeLoops()
    
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
    
    Dim LastRow As Long
    LastRow = ActiveSheet.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    
    Dim TickerPrint As Integer
    TickerPrint = 2
    
    For ticker = 2 To LastRow
        If ws.Cells(ticker, 1).Value <> ws.Cells(ticker - 1, 1) Then ws.Cells(TickerPrint, 9).Value = ws.Cells(ticker, 1)
        If ws.Cells(TickerPrint, 9).Value <> "" Then TickerPrint = TickerPrint + 1
    Next ticker
    
    TickerPrint = 2
    Dim YearStart As Double
    YearStart = 0
    Dim YearEnd As Double
    YearEnd = 0
   
    For yearly = 2 To LastRow
        If ws.Cells(yearly, 1).Value = ws.Cells(TickerPrint, 9).Value And ws.Cells(yearly - 1, 1).Value <> ws.Cells(TickerPrint, 9) Then YearStart = ws.Cells(yearly, 3).Value
        If ws.Cells(yearly, 1).Value = ws.Cells(TickerPrint, 9).Value And ws.Cells(yearly + 1, 1).Value <> ws.Cells(TickerPrint, 9) Then YearEnd = ws.Cells(yearly, 6).Value
        ws.Cells(TickerPrint, 10).Value = (YearEnd - YearStart)
        If ws.Cells(TickerPrint, 10).Value = 0 Then ws.Cells(TickerPrint, 11).Value = 0 Else: ws.Cells(TickerPrint, 11).Value = (ws.Cells(TickerPrint, 10).Value / ws.Cells(yearly, 3).Value)
        If ws.Cells(yearly + 1, 1).Value <> ws.Cells(TickerPrint, 9).Value Then YearStart = 0
        If ws.Cells(yearly + 1, 1).Value <> ws.Cells(TickerPrint, 9).Value Then YearEnd = 0
        If ws.Cells(yearly + 1, 1).Value <> ws.Cells(TickerPrint, 9).Value Then TickerPrint = TickerPrint + 1
    Next yearly

    TickerPrint = 2
    Dim Vols As LongLong
  
    For volume = 2 To LastRow
        If ws.Cells(volume, 1).Value = ws.Cells(TickerPrint, 9).Value Then Vols = Vols + ws.Cells(volume, 7).Value
        If ws.Cells(volume + 1, 1).Value <> ws.Cells(TickerPrint, 9).Value Then ws.Cells(TickerPrint, 12).Value = Vols
        If ws.Cells(volume + 1, 1).Value <> ws.Cells(TickerPrint, 9).Value Then Vols = 0
        If ws.Cells(volume + 1, 1).Value <> ws.Cells(TickerPrint, 9).Value Then TickerPrint = TickerPrint + 1
    Next volume
    
    Dim GreatPlus As Double
    GreatPlus = 0
    Dim GreatMinus As Double
    GreatMinus = 0
    Dim GreatVol As LongLong
    GreatVol = 0
    Dim PlusTicker As String
    Dim MinusTicker As String
    Dim VolTicker As String
    
    For greatest = 2 To LastRow
        If ws.Cells(greatest, 11).Value > GreatPlus Then GreatPlus = ws.Cells(greatest, 11).Value
        ws.Cells(2, 17).Value = GreatPlus
        If ws.Cells(2, 17).Value = ws.Cells(greatest, 11).Value Then PlusTicker = ws.Cells(greatest, 9).Value
        ws.Cells(2, 16).Value = PlusTicker
        
        If ws.Cells(greatest, 11).Value < GreatMinus Then GreatMinus = ws.Cells(greatest, 11).Value
        ws.Cells(3, 17).Value = GreatMinus
        If ws.Cells(3, 17).Value = ws.Cells(greatest, 11).Value Then MinusTicker = ws.Cells(greatest, 9).Value
        ws.Cells(3, 16).Value = MinusTicker
        
        If ws.Cells(greatest, 12).Value > GreatVol Then GreatVol = ws.Cells(greatest, 12).Value
        ws.Cells(4, 17).Value = GreatVol
        If ws.Cells(4, 17).Value = ws.Cells(greatest, 12).Value Then VolTicker = ws.Cells(greatest, 9).Value
        ws.Cells(4, 16).Value = VolTicker
    Next greatest
    

   Next ws

End Sub



