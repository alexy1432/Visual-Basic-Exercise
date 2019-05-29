VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub EasyAssignment()
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row
        Cells(1, 8).Value = "Ticker"
        Cells(1, 9).Value = "Total Stock Volume"

        Dim Ticker_Name As String
        Dim Volume As Double
        Volume = 0
        Dim Row As Double
        Row = 2
        Dim Column As Integer
        Column = 1
        Dim i As Long

        
        For i = 2 To LastRow
            If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then
                Ticker_Name = Cells(i, Column).Value
                Cells(Row, Column + 8).Value = Ticker_Name
                Volume = Volume + Cells(i, Column + 6).Value
                Cells(Row, Column + 9).Value = Volume
                Row = Row + 1
                Volume = 0
            End If
        Next i
        
        

        
End Sub
