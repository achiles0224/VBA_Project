Sub ValidateInvoiceAndAllowance()
    Dim wsEPB As Worksheet
    Dim wsInvoice As Worksheet
    Dim wsAllowance As Worksheet
    Dim lastRowEPB As Long
    Dim lastRowInvoice As Long
    Dim lastRowAllowance As Long
    Dim i As Long
    Dim taxCode As String
    Dim voidFlag As String
    Dim invoiceStatus As String
    Dim found As Range
    Dim progress As Double
    Dim startTime As Double
    Dim elapsedTime As Double
    Dim remainingTime As Double
    Dim estimatedEndTime As Date
    Dim progressMessage As String
    Dim exitLoop As Boolean
    Dim cell As Range
    Dim colorRange As Range

    ' 設定工作表
    Set wsEPB = ThisWorkbook.Sheets("EPB200FD")
    Set wsInvoice = ThisWorkbook.Sheets("鯨躍-匯出發票")
    Set wsAllowance = ThisWorkbook.Sheets("鯨躍-匯出折讓單")
    
    ' 獲取工作表的最後一行
    lastRowEPB = wsEPB.Cells(wsEPB.Rows.Count, "A").End(xlUp).Row
    lastRowInvoice = wsInvoice.Cells(wsInvoice.Rows.Count, "A").End(xlUp).Row
    lastRowAllowance = wsAllowance.Cells(wsAllowance.Rows.Count, "F").End(xlUp).Row
    
    ' 清除EPB200FD A2之後的格式
    wsEPB.Range("A2:G" & lastRowEPB).ClearFormats
    
    ' 記錄開始時間
    startTime = Timer
    
    ' 開始處理
    Application.StatusBar = "Processing..."
    Application.ScreenUpdating = False ' 禁止屏幕更新，提高速度
    Application.Calculation = xlCalculationManual ' 禁止自動計算
    
    ' 從EPB200FD的A欄開始，進行檢查
    For i = 2 To lastRowEPB ' 假設從第2行開始有資料
        ' 初始化exitLoop
        exitLoop = False
        
        ' 計算進度並顯示
        progress = (i - 1) / (lastRowEPB - 1)
        elapsedTime = Timer - startTime
        
        ' 計算剩餘時間
        If i > 2 Then
            remainingTime = elapsedTime / (i - 1) * (lastRowEPB - i)
        Else
            remainingTime = 0 ' 當處理不到兩行時，設為0
        End If
        
        ' 預估結束時間
        estimatedEndTime = Now + remainingTime / 86400 ' Convert seconds to days
        
        ' 顯示進度和預估結束時間
        progressMessage =  "Processing: " & Format(progress, "0%") & " Estimated Completion:" & Format(estimatedEndTime, "hh:mm:ss AM/PM")
        Application.StatusBar = progressMessage
        
        ' 取得稅碼 (F欄) 和作廢標籤 (G欄)
        taxCode = wsEPB.Cells(i, 6).Value ' F欄
        voidFlag = wsEPB.Cells(i, 7).Value ' G欄
        
        ' 稅碼是33，檢查折讓單資料
        If taxCode = "33" Then
            ' 使用 Find 來搜尋折讓單中的發票號碼
            Set found = wsAllowance.Range("F2:F" & lastRowAllowance).Find(wsEPB.Cells(i, 1).Value, LookIn:=xlValues)
            If found Is Nothing Then
                ' 如果找不到，將EPB200FD A欄的顏色設為黃色
                wsEPB.Cells(i, 1).Interior.Color = vbYellow
            End If
        
        ' 稅碼是35，檢查作廢標籤
        ElseIf taxCode = "35" Then
            If voidFlag = "Yes" Then
                ' 作廢標籤為Yes，檢查發票狀態是否為"作廢"
                Set found = wsInvoice.Range("A2:A" & lastRowInvoice).Find(wsEPB.Cells(i, 1).Value, LookIn:=xlValues)
                If Not found Is Nothing Then
                    invoiceStatus = wsInvoice.Cells(found.Row, 16).Value ' P欄
                    If invoiceStatus <> "作廢" Then
                        ' 如果狀態不是"作廢"，標記紅色
                        wsEPB.Cells(i, 1).Interior.Color = vbRed
                    End If
                Else
                    ' 如果發票號碼在發票工作表找不到，標記紅色
                    wsEPB.Cells(i, 1).Interior.Color = vbRed
                End If
            ElseIf voidFlag = "No" Then
                ' 作廢標籤為No，檢查發票狀態是否為"開立"
                Set found = wsInvoice.Range("A2:A" & lastRowInvoice).Find(wsEPB.Cells(i, 1).Value, LookIn:=xlValues)
                If Not found Is Nothing Then
                    invoiceStatus = wsInvoice.Cells(found.Row, 16).Value ' P欄
                    If invoiceStatus <> "開立" Then
                        ' 如果狀態不是"開立"，標記紅色
                        wsEPB.Cells(i, 1).Interior.Color = vbRed
                    End If
                Else
                    ' 如果發票號碼在發票工作表找不到，標記紅色
                    wsEPB.Cells(i, 1).Interior.Color = vbRed
                End If
            End If
        End If
    Next i

    ' 進行顏色排序，將黃色和紅色的資料排在前面
    With wsEPB
        ' 設定範圍（從A2開始直到最後一行）
        Set colorRange = .Range("A2:G" & lastRowEPB)
        
        ' 先排序顏色為紅色的資料
        colorRange.Sort Key1:=.Range("A2"), Order1:=xlDescending, Type:=xlSortNormal, _
            OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, Header:=xlNo

        ' 然後排序顏色為黃色的資料
        colorRange.Sort Key1:=.Range("A2"), Order1:=xlAscending, Type:=xlSortNormal, _
            OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, Header:=xlNo
    End With

    ' 重設狀態列
    Application.StatusBar = False
    Application.ScreenUpdating = True ' 重新啟用屏幕更新
    Application.Calculation = xlCalculationAutomatic ' 恢復自動計算

    MsgBox "Invoice validation is complete.", vbInformation
End Sub

Sub ClearDataAndFormats()
    Dim wsEPB As Worksheet
    Dim wsInvoice As Worksheet
    Dim wsAllowance As Worksheet
    Dim lastRowEPB As Long
    Dim lastRowInvoice As Long
    Dim lastRowAllowance As Long
    Dim lastColEPB As Long
    Dim lastColInvoice As Long
    Dim lastColAllowance As Long
    
    ' 設定工作表
    Set wsEPB = ThisWorkbook.Sheets("EPB200FD")
    Set wsInvoice = ThisWorkbook.Sheets("鯨躍-匯出發票")
    Set wsAllowance = ThisWorkbook.Sheets("鯨躍-匯出折讓單")
    
    ' 獲取工作表的最後一行和最後一列
    lastRowEPB = wsEPB.Cells(wsEPB.Rows.Count, "A").End(xlUp).Row
    lastRowInvoice = wsInvoice.Cells(wsInvoice.Rows.Count, "A").End(xlUp).Row
    lastRowAllowance = wsAllowance.Cells(wsAllowance.Rows.Count, "F").End(xlUp).Row
    
    lastColEPB = wsEPB.Cells(1, wsEPB.Columns.Count).End(xlToLeft).Column
    lastColInvoice = wsInvoice.Cells(1, wsInvoice.Columns.Count).End(xlToLeft).Column
    lastColAllowance = wsAllowance.Cells(1, wsAllowance.Columns.Count).End(xlToLeft).Column
    
    ' 清除 EPB200FD 工作表的資料和格式 (不清除表頭)
    wsEPB.Range(wsEPB.Cells(2, 1), wsEPB.Cells(lastRowEPB, lastColEPB)).ClearContents
    wsEPB.Range(wsEPB.Cells(2, 1), wsEPB.Cells(lastRowEPB, lastColEPB)).ClearFormats
    
    ' 清除 鯨躍-匯出發票 工作表的資料和格式 (不清除表頭)
    wsInvoice.Range(wsInvoice.Cells(2, 1), wsInvoice.Cells(lastRowInvoice, lastColInvoice)).ClearContents
    wsInvoice.Range(wsInvoice.Cells(2, 1), wsInvoice.Cells(lastRowInvoice, lastColInvoice)).ClearFormats
    
    ' 清除 鯨躍-匯出折讓單 工作表的資料和格式 (不清除表頭)
    wsAllowance.Range(wsAllowance.Cells(2, 1), wsAllowance.Cells(lastRowAllowance, lastColAllowance)).ClearContents
    wsAllowance.Range(wsAllowance.Cells(2, 1), wsAllowance.Cells(lastRowAllowance, lastColAllowance)).ClearFormats

    MsgBox "Data and formats have been cleared (except headers).", vbInformation
End Sub
