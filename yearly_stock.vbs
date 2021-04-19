VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub yearly_stock():

    'declare variables and counter
    Dim lRow As Double  ' to be not over flow
    Dim i As Double ' to be not over flow
    Dim count_result As Double
    Dim ticker As String
    Dim stock_volume As Double
    Dim open_price As Double
    Dim close_price As Double
    
    
    Set wb = ActiveWorkbook
    Set ws = wb.Sheets
    
    'outer loop to fo every sheet of the entire workwook
    For Each ws In ThisWorkbook.Worksheets
        ' declare to be sure every sheets ending row could be varies
        lRow = ws.Cells(Rows.count, 1).End(xlUp).Row
        ' initialize from 2 since it starts row 2
        i = 2
        ticker = ws.Cells(i, 1)
        open_price = ws.Cells(i, 3)
        
        ' loop through row
        count_result = 2
        Do While i <= lRow ' variable row ending for each sheet make sure picks up all row
            If ws.Cells(i, 1) <> ticker Then
            ' assign cariables for all new cells
                ws.Cells(count_result, 9) = ticker
                ws.Cells(count_result, 12) = stock_volume
                close_price = ws.Cells(i - 1, 6)
                ws.Cells(count_result, 10) = close_price - open_price
                ' format cpolors
                If ws.Cells(count_result, 10) < 0 Then
                    ws.Cells(count_result, 10).Interior.Color = vbRed
                Else
                    ws.Cells(count_result, 10).Interior.Color = vbGreen
                End If
                ' format percentage
                If close_price = 0 Then
                    If open_price = 0 Then
                        ws.Cells(count_result, 11) = FormatPercent(0, 2)
                    Else
                        ws.Cells(count_result, 11) = FormatPercent(-1, 2)
                    End If
                Else
                    ws.Cells(count_result, 11) = FormatPercent((close_price - open_price) / close_price, 2)
                End If
                ' copy actual values to cells
                ticker = ws.Cells(i, 1)
                stock_volume = ws.Cells(i, 7)
                open_price = ws.Cells(i, 3)
                count_result = count_result + 1
            Else
            ' accumulates stock value
                stock_volume = stock_volume + ws.Cells(i, 7)
            End If
            i = i + 1
        Loop
        'calculating part
        close_price = ws.Cells(i - 1, 6)
        ws.Cells(count_result, 10) = close_price - open_price
        'format color
        If ws.Cells(count_result, 10) < 0 Then
            ws.Cells(count_result, 10).Interior.Color = vbRed
        Else
            ws.Cells(count_result, 10).Interior.Color = vbGreen
        End If
        'format percentage
        If close_price = 0 Then
            If open_price = 0 Then
                ws.Cells(count_result, 11) = FormatPercent(0, 2)
            Else
                ws.Cells(count_result, 11) = FormatPercent(-1, 2)
            End If
        Else
            ws.Cells(count_result, 11) = FormatPercent((close_price - open_price) / close_price, 2)
        End If
        
        ws.Cells(count_result, 9) = ws.Cells(i - 1, 1)
        ws.Cells(count_result, 12) = stock_volum
        
    Next ws
End Sub


