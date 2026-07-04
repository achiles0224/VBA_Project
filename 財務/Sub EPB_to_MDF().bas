Sub EPB_to_MDF()

    Application.ScreenUpdating = False

    Dim ws_raw_MDF As Worksheet, ws As Worksheet
    Dim MDF_rowsCounts As Long, MDF_lastRow As Long
    Dim MDF_list_rowCounts As Long, mdf_list_lastRow As Long
    Dim mdf_list As Range
    Dim col As Long
    
    '--- 來源工作表
    Set ws_raw_MDF = ThisWorkbook.Sheets("800FA")
    MDF_rowsCounts = ws_raw_MDF.Range("A1").CurrentRegion.Rows.Count
    MDF_lastRow = ws_raw_MDF.Cells(ws_raw_MDF.Rows.Count, 1).End(xlUp).Row
    
    '--- MDF 對照表
    Set ws = ThisWorkbook.Sheets("Readme")
    Set mdf_list = ws.Range("A20").CurrentRegion
    mdf_list_lastRow = mdf_list.Rows.Count
    
    '========== 檢查是否已有 I、J、S 欄位 ==========
    ' 找第一列的標題，若有相同名稱先刪除
    For col = ws_raw_MDF.Cells(1, ws_raw_MDF.Columns.Count).End(xlToLeft).Column To 1 Step -1
        Select Case ws_raw_MDF.Cells(1, col).Value
            Case "MDF%", "類別", "MDF"
                ws_raw_MDF.Columns(col).Delete
        End Select
    Next col
    
    '========== 插入新欄位 ==========
    ws_raw_MDF.Columns("I:J").Insert Shift:=xlToRight
    ws_raw_MDF.Columns("S:S").Insert Shift:=xlToRight
    
    ws_raw_MDF.Range("I1").Value = "MDF%"
    ws_raw_MDF.Range("J1").Value = "類別"
    ws_raw_MDF.Range("S1").Value = "MDF"

    With ws_raw_MDF.Range("I2:I" & MDF_rowsCounts)
        .Formula = "=MDF_Lookup_Text(K2, M2, " & _
               "Readme!$A$20:$A$" & mdf_list_lastRow + 20 & ", " & _
               "Readme!$C$20:$C$" & mdf_list_lastRow + 20 & ", " & _
               "Readme!$D$20:$D$" & mdf_list_lastRow + 20 & ")"
        .NumberFormat = "0.00%"
    End With

    With ws_raw_MDF.Range("J2:J" & MDF_rowsCounts)
        .Formula = "=MDF_Lookup_Category(K2, M2, N2, " & _
               "Readme!$A$20:$A$" & mdf_list_lastRow + 20 & ", " & _
               "Readme!$C$20:$C$" & mdf_list_lastRow + 20 & ", " & _
               "Readme!$E$20:$E$" & mdf_list_lastRow + 20 & ")"
    End With

    With ws_raw_MDF.Range("S2:S" & MDF_rowsCounts)
        .Formula = "=ROUND(Q2*R2*I2,0)"
    End With

    msgbox "已完成巨集，稍等公式更新..."
    
End Sub


' 文字版：cat4/cat6 與對照表一律先正規化為字串後再比對
Public Function MDF_Lookup_Text( _
    ByVal cat4 As Variant, ByVal cat6 As Variant, _
    ByVal rngCat4 As Range, ByVal rngCat6 As Range, ByVal rngMDF As Range, _
    Optional ByVal caseSensitive As Boolean = False) As Variant

    Dim v4, v6, vM
    Dim i As Long, n As Long
    Dim sCat4 As String, sCat6 As String
    Dim cmpMode As VbCompareMethod

    ' 一次讀入（最快）
    v4 = rngCat4.Value2
    v6 = rngCat6.Value2
    vM = rngMDF.Value2
    n = rngCat4.Rows.Count

    ' 比對模式：預設不分大小寫
    cmpMode = IIf(caseSensitive, vbBinaryCompare, vbTextCompare)

    ' 將輸入鍵正規化為文字
    sCat4 = NormText(cat4)
    sCat6 = NormText(cat6)

    ' 1) 先找 (類別4, 類別6) 全配
    For i = 1 To n
        If StrComp(NormText(v4(i, 1)), sCat4, cmpMode) = 0 _
        And StrComp(NormText(v6(i, 1)), sCat6, cmpMode) = 0 Then
            MDF_Lookup_Text = vM(i, 1)
            Exit Function
        End If
    Next i

    ' 2) 再找 (類別4 配, 類別6 空白)
    For i = 1 To n
        If StrComp(NormText(v4(i, 1)), sCat4, cmpMode) = 0 _
        And Len(NormText(v6(i, 1))) = 0 Then
            MDF_Lookup_Text = vM(i, 1)
            Exit Function
        End If
    Next i

    MDF_Lookup_Text = ""   ' 找不到 → 回傳空白
End Function

' 將任何輸入轉成「去頭尾空白且正規化後」的字串
Private Function NormText(ByVal v As Variant) As String
    Dim s As String
    s = CStr(v & "")
    ' 換掉常見的不可見空白：NBSP 與全形空白
    s = Replace$(s, Chr(160), " ")
    s = Replace$(s, ChrW(&H3000), " ")
    NormText = Trim$(s)
End Function


Public Function MDF_Lookup_Category( _
    ByVal cat4 As Variant, ByVal cat6 As Variant, _
    ByVal cat6_name As String, _
    ByVal rngCat4 As Range, ByVal rngCat6 As Range, _
    ByVal rngCategory As Range, _
    Optional ByVal caseSensitive As Boolean = False) As Variant
    
    Dim v4, v6, vCat
    Dim i As Long, n As Long
    Dim sCat4 As String, sCat6 As String
    Dim cmpMode As VbCompareMethod
    
    ' 把對照表資料一次讀入陣列
    v4 = rngCat4.Value2
    v6 = rngCat6.Value2
    vCat = rngCategory.Value2
    n = rngCat4.Rows.Count
    
    ' 比對模式：預設不分大小寫
    cmpMode = IIf(caseSensitive, vbBinaryCompare, vbTextCompare)
    
    ' 把輸入鍵轉成字串
    sCat4 = NormText(cat4)
    sCat6 = NormText(cat6)
    
    ' 1) 先找 (cat4, cat6) 完全匹配
    For i = 1 To n
        If StrComp(NormText(v4(i, 1)), sCat4, cmpMode) = 0 _
        And StrComp(NormText(v6(i, 1)), sCat6, cmpMode) = 0 Then
            If NormText(vCat(i, 1)) = "類別6" Then
                MDF_Lookup_Category = cat6_name      ' 回傳原本工作表的 cat6 值
            Else
                MDF_Lookup_Category = vCat(i, 1) ' 回傳對照表的類別名稱
            End If
            Exit Function
        End If
    Next
    
    ' 2) 再找 (cat4 匹配, cat6 空白)
    For i = 1 To n
        If StrComp(NormText(v4(i, 1)), sCat4, cmpMode) = 0 _
        And Len(NormText(v6(i, 1))) = 0 Then
            If NormText(vCat(i, 1)) = "類別6" Then
                MDF_Lookup_Category = cat6_name
            Else
                MDF_Lookup_Category = vCat(i, 1)
            End If
            Exit Function
        End If
    Next
    
    ' 找不到 → 空白
    MDF_Lookup_Category = ""
End Function
