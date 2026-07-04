Sub MaxConsecutivePurchase()


Dim i As Long, j As Long
Dim max_consecutive As Long, consecutive As Long
Dim max_purchase As Long
Dim purchases() As Variant

Dim last_col As Long
Dim lastColumn As Long

Worksheets("會員消費次數").Select
lastColumn = Cells(1, Columns.Count).End(xlToLeft).Column

If Cells(1, lastColumn).Value = "當天最大購買次數" Then
   Range(Cells(1, lastColumn - 1), Cells(ActiveSheet.Rows.Count, lastColumn)).EntireColumn.Delete
    'MsgBox "已刪除最後兩欄。"
Else
    'MsgBox "最後一欄表頭不是「當天最大購買次數」，無法刪除最後兩欄。"
End If

last_col = Cells(1, Columns.Count).End(xlToLeft).Column


For i = 2 To Cells(Rows.Count, "B").End(xlUp).Row
    ' Resize purchases array based on number of columns with dates
    ReDim purchases(1 To last_col - 3)
    ' Fill purchases array with values from cells
    For j = 3 To last_col - 1
        If IsNumeric(Cells(i, j).Value) Then
            purchases(j - 2) = Cells(i, j).Value
        Else
            purchases(j - 2) = 0
        End If
    Next j
    ' Find max consecutive purchases and max purchase in a single month
    consecutive = 0
    max_consecutive = 0
    max_purchase = 0
    For j = LBound(purchases) To UBound(purchases)
        If purchases(j) > 0 Then
            consecutive = consecutive + 1
            If consecutive > max_consecutive Then
                max_consecutive = consecutive
            End If
            If purchases(j) > max_purchase Then
                max_purchase = purchases(j)
            End If
        Else
            consecutive = 0
        End If
    Next j
    ' Output results to last column
    Cells(i, last_col + 1).Value = max_consecutive
    Cells(i, last_col + 2).Value = max_purchase
Next i

   ' Add headers to new columns
    Cells(1, last_col + 1).Value = "最大連續購買天數"
    Cells(1, last_col + 2).Value = "當天最大購買次數"
    
 ' 檢查統計是否存在
    Dim sheet6 As Worksheet
    On Error Resume Next
    Set sheet6 = Worksheets("統計")
    On Error GoTo 0
    If sheet6 Is Nothing Then
        Set sheet6 = Worksheets.Add(after:=Worksheets(Worksheets.Count))
        sheet6.Name = "統計"
    Else
        ' 如果統計已存在，則先清除所有內容
        sheet6.Cells.ClearContents
    End If
    
    ' 將會員消費次數複製到統計
    Worksheets("會員消費次數").Cells.Copy
    sheet6.Cells.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, _
                               Operation:=xlNone, _
                               SkipBlanks:=False, _
                               Transpose:=False
    
    ' 刪除"C"欄到倒數第四欄
    Dim lastCol6 As Long
    lastCol6 = sheet6.Cells(1, Columns.Count).End(xlToLeft).Column
    If lastCol6 >= 4 Then
        sheet6.Range(sheet6.Cells(1, 3), sheet6.Cells(sheet6.Rows.Count, lastCol6 - 3)).EntireColumn.Delete
    End If

    Worksheets("Start").select
    MsgBox "會員消費次數執行完成"


End Sub


Sub addQuantity()

    Worksheets("會員消費主機數量").select

    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim lastRow1 As Long, lastRow2 As Long
    Dim i As Long, j As Long
    Dim vipcode As String, quantity As Long
    Dim vipCol1 As Long, vipCol2 As Long
    
    '指定要操作的工作表
    Set ws1 = ThisWorkbook.Worksheets("統計")
    Set ws2 = ThisWorkbook.Worksheets("會員消費主機數量")
    
    '找到每個工作表的最後一行
    lastRow1 = ws1.Cells(ws1.Rows.Count, "B").End(xlUp).Row
    lastRow2 = ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row
    
    '找到VIP代碼和數量所在的欄位
    vipCol1 = ws1.Range("B:B").Column
    vipCol2 = ws2.Range("A:A").Column + 1 '數量所在的欄位是A欄的下一欄
    
    '對於每個VIP代碼，在會員消費次數中查找相應的行
     For i = 2 To lastRow1 '假設第一行是標題，所以從第二行開始查找
        vipcode = ws1.Cells(i, 2).Value '取得VIP代碼
        
        Set sumRange = ws2.Range("B:B")
        Set criteriaRange = ws2.Range("A:A")
        
        '使用 SUMIF 函數進行加總
        ws1.Cells(i, ws1.Columns.Count).End(xlToLeft).Offset(0, 1).Value = WorksheetFunction.SumIf(criteriaRange, vipcode, sumRange)
    Next i
        ws1.Cells(1, ws1.Cells(1, ws1.Columns.Count).End(xlToLeft).Column + 1).Value = "購買次數"
    Worksheets("Start").select
    MsgBox "會員消費主機數量"
End Sub

Sub xadd888()

    Worksheets("會員消費點折金").select
    Application.StatusBar = "execute program..."
    Dim ws5 As Worksheet, ws6 As Worksheet
    Dim lastrow5 As Long, lastrow6 As Long
    Dim i As Long, j As Long
    Dim vipcode As String
    Dim StartTime As Double
    Dim Minutes As Integer
    Dim Seconds As Integer

    StartTime = Timer ' 取得目前時間

    Set ws5 = ThisWorkbook.Worksheets("會員消費點折金")
    Set ws6 = ThisWorkbook.Worksheets("統計")

    lastrow5 = ws5.Cells(ws5.Rows.Count, "A").End(xlUp).Row
    lastrow6 = ws6.Cells(ws6.Rows.Count, "A").End(xlUp).Row
    lastColumn6 = ws6.Cells(1, Columns.Count).End(xlToLeft).Column

    Set VlookupRange = ws5.Range("A:C")
    Set CountRange = ws5.Range("B:B")
    Set criteriaRange = ws5.Range("A:A")
    Set AmountRange = ws5.Range("C:C")


    For i = 2 To lastrow6
        vipcode = ws6.Cells(i, 2).Value
        ws6.Cells(i, lastColumn6 + 1).Value = "=ifna(VLOOKUP(B" & i & ",會員消費點折金!A:C,2,FALSE),0)"
        ws6.Cells(i, lastColumn6 + 2).Value = "=ifna(VLOOKUP(B" & i & ",會員消費點折金!A:C,3,FALSE),0)"

        'ws6.Cells(i, ws6.Columns.Count).End(xlToLeft).Offset(0, 1).Value = IIf(IsError(Application.VLookup(vipcode, VlookupRange, 2, False)), 0, Application.VLookup(vipcode, VlookupRange, 2, False))
        'ws6.Cells(i, ws6.Columns.Count).End(xlToLeft).Offset(0, 1).Value = IIf(IsError(Application.VLookup(vipcode, VlookupRange, 3, False)), 0, Application.VLookup(vipcode, VlookupRange, 3, False))
        Application.StatusBar = "已完成 " & Format(i / lastrow6, "0%") & "。"
    Next i
    ws6.Cells(1, ws6.Columns.Count).End(xlToLeft).Offset(0, 1).Value = "888使用次數"
    ws6.Cells(1, ws6.Columns.Count).End(xlToLeft).Offset(0, 1).Value = "888使用點數"
  
    Minutes = Int((Timer - StartTime) / 60)
    Seconds = Round((Timer - StartTime) Mod 60, 2)

    Application.StatusBar = False
    Worksheets("Start").select
    MsgBox "會員消費點折金計算執行完成 花費：" & Minutes & " 分 " & Seconds & " 秒"
End Sub


Sub CopyData()
    
    Worksheets("統計").select
    Dim last_col As Integer
    Dim i As Integer, j As Integer
    Dim ws1 As Worksheet, wsx As Worksheet
    
    Dim TotalSalesDay As Integer, ContinueSalesDay As Integer  '會員消費次數：總天數、連續消費天數'
    Dim TopSalesCountbyDay As Integer '會員消費次數：同一天最大消費次數'
    Dim VipBuyItemTotal As Integer  '會員消費主機：總數量'
    Dim Vip888Count As Integer, VipUsed888Point As Long     '會員消費點折金：點折金次數、點折金使用數量'

    Set wsfrst = ThisWorkbook.Worksheets("Start")
    Set ws1 = ThisWorkbook.Worksheets("統計")
    Set wsx = ThisWorkbook.Worksheets("比對結果")
    
    TotalSalesDay = wsfrst.Range("E5").Value
    ContinueSalesDay = wsfrst.Range("E6").Value
    TopSalesCountbyDay = wsfrst.Range("E7").Value
    VipBuyItemTotal = wsfrst.Range("E10").Value
    Vip888Count = wsfrst.Range("E14").Value
    VipUsed888Point = wsfrst.Range("E15").Value

    last_col = ws1.Cells(1, ws1.Columns.Count).End(xlToLeft).Column
    
    Dim row_counter As Integer
    row_counter = 2

    ws1.Range("A1:H1").Copy
    wsx.Range("A1:H1").PasteSpecial Paste:=xlPasteAll
    
    For i = 2 To ws1.Cells(ws1.Rows.Count, "B").End(xlUp).Row
        If ws1.Cells(i, 3).Value > TotalSalesDay Or ws1.Cells(i, 4).Value > ContinueSalesDay Or ws1.Cells(i, 5).Value > TopSalesCountbyDay Or ws1.Cells(i, 6).Value > VipBuyItemTotal Or Abs(ws1.Cells(i, 7).Value) > Vip888Count Or Abs(ws1.Cells(i, 7).Value) > VipUsed888Point Then
            ws1.Cells(i, 1).Copy
            wsx.Cells(row_counter, 1).PasteSpecial Paste:=xlPasteAll
            wsx.Cells(row_counter, 2).Value = ws1.Cells(i, 2).Value
            wsx.Cells(row_counter, 3).Value = ws1.Cells(i, 3).Value
            wsx.Cells(row_counter, 4).Value = ws1.Cells(i, 4).Value
            wsx.Cells(row_counter, 5).Value = ws1.Cells(i, 5).Value
            wsx.Cells(row_counter, 6).Value = ws1.Cells(i, 6).Value
            wsx.Cells(row_counter, 7).Value = ws1.Cells(i, 7).Value
            wsx.Cells(row_counter, 8).Value = ws1.Cells(i, 8).Value
            row_counter = row_counter + 1
            
            'If ws1.Cells(i + 1, last_col - 1).Value > 3 Or ws1.Cells(i + 1, last_col).Value > 3 Then
                ' Continue copying data to next row in wsx
            'Else
                ' Move to next row in ws1
            ''    i = i + 1
            'End If
        End If
    Next i
    Worksheets("比對結果").select
    msgbox "比對完成"
    
End Sub

Sub DeleteWorksheet()
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("會員消費次數") ' 替換為你的工作表名稱
    Set ws1 = thisworkbook.Worksheets("統計")
    set ws2 = ThisWorkbook.Worksheets("會員消費主機數量")
    set ws3 = ThisWorkbook.Worksheets("會員消費點折金")
    set ws4 = ThisWorkbook.Worksheets("比對結果")
    
    Application.DisplayAlerts = False ' 禁用刪除警告
    ws.Cells.ClearContents ' 清空所有儲存格的內容
    ws.Cells.ClearFormats ' 清空所有儲存格的格式
    ws1.Cells.ClearContents ' 清空所有儲存格的內容
    ws1.Cells.ClearFormats ' 清空所有儲存格的格式
    ws2.Cells.ClearContents ' 清空所有儲存格的內容
    ws2.Cells.ClearFormats ' 清空所有儲存格的格式
    ws3.Cells.ClearContents ' 清空所有儲存格的內容
    ws3.Cells.ClearFormats ' 清空所有儲存格的格式
    ws4.Cells.ClearContents ' 清空所有儲存格的內容
    ws4.Cells.ClearFormats ' 清空所有儲存格的格式
    Application.DisplayAlerts = True ' 啟用刪除警告
    
    msgbox "全部工作表資料已清除"
End Sub
