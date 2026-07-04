Sub EDI_SLS()
    Worksheets("TempData").Range("A:CC").Clear
    Worksheets("SLS_CSV").Range("A:A").Clear
    Worksheets("APPLE_EDI_SLS").Select

    Dim ExcelRow, SLSRows As Integer
    'Dim sriQTY As Long
    Dim d, saleDate As Date
    saleDate = Worksheets("APPLE_EDI_SLS").Range("I1").Value
    d = Worksheets("APPLE_EDI_SLS").Range("L1").Value
    If Range("L1") = Range("M1") Then
        a = 0
    Else
        Worksheets("ShopID").Range("I1").Value = Format(1, "0000")
        Worksheets("APPLE_EDI_SLS").Range("M1").Value = Range("L1").Value
    End If

    'dim SerialNumber as
    SLSRows = Worksheets("APPLE_EDI_SLS").Range("A1").CurrentRegion.Rows.Count
    '檢查序號是否為空'
    sriQTY = 0
    For Row = 2 To SLSRows Step 1
        SriID = Cells(Row, 3).Value
        sriQTY = Cells(Row, 4).Value
        If SriID = "" Then
            MsgBox ("Row. " & Row & " Serial number is Empty, Please Check Data.")
            Worksheets("APPLE_EDI_SLS").Cells(Row, 3).Select
            Exit Sub
        End If
        If sriQTY > 1 Then
            MsgBox ("Row. " & Row & " QTY > 1 or < -1, Please Check QTY")
            Cells(Row, 4).Select
            Exit Sub
        End If
    Next Row

    '序號加總'
    With Application
        ' Turn off screen updating to increase performance
        .ScreenUpdating = True
        Dim LastColumn As Integer
        LastColumn = Cells.Find(what:="*", after:=Range("C1"), SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
        With Range("C1:C" & Cells(Rows.Count, 1).End(xlUp).Row)
            ' Use AdvanceFilter to filter unique values
            .AdvancedFilter Action:=xlFilterCopy, copytorange:=Worksheets("TempData").Range("C1"), Unique:=True
            '.SpecialCells(xlCellTypeVisible).Offset(0, LastColumn - 1).Value = 1
            On Error Resume Next
            ActiveSheet.ShowAllData
            'Delete the blank rows
            'Columns(LastColumn).SpecialCells(xlCellTypeBlanks).EntireRow.Delete
            Err.Clear
        End With
        'Columns(LastColumn).Clear
        .ScreenUpdating = True
    End With

    Worksheets("TempData").Select
    Worksheets("TempData").Range("A1").Value = "SHopID"
    Worksheets("TempData").Range("B1").Value = "Moedl"
    Worksheets("TempData").Range("C1").Value = "SerialNumber"
    Worksheets("TempData").Range("D1").Value = "QTY"

    Set srcRange = Worksheets("APPLE_EDI_SLS").Range("C1").CurrentRegion.Columns(3)
    sriCount = Worksheets("TempData").Range("C1").CurrentRegion.Rows.Count

    '帶入門市型號，消退在型號備註'
    For i = 2 To sriCount Step 1
        Set fndRange = srcRange.Find(what:=Cells(i, 3))
        If Not fndRange Is Nothing Then
            fstAddress = fndRange.Address
            'i = 1
            Worksheets("TempData").Cells(i, 1).Value = fndRange.Offset(, -2).Value
            Worksheets("TempData").Cells(i, 2).Value = fndRange.Offset(, -1).Value
            TempQTY = 0
            Do
                
                TempQTY = TempQTY + fndRange.Offset(, 1).Value
                Set fndRange = srcRange.FindNext(after:=fndRange)
                'i = i + 1
            Loop Until fndRange.Address = fstAddress
            Worksheets("TempData").Cells(i, 4).Value = TempQTY
            
            If TempQTY < 0 Then
                Cells(i, 2).Value = Cells(i, 2).Value & "_Return"
            End If
        End If
    Next i

    '刪除數量為0'
    For i = sriCount To 1 Step -1
 
        If Cells(i, 4) = 0 Then
     
            Rows(i).Select
     
            Selection.Delete Shift:=xlUp
     
        End If

    Next i

    '插入一欄 將門市與型號合併'
    Range("C:C").Insert xlShiftToRight

    sriCount = Worksheets("TempData").ragne("A1").CurrentRegion.Rows.Count
    Range("C1").Value = "SM"
    For i = 2 To sriCount Step 1
        Cells(i, 3).Value = Cells(i, 1).Value & Cells(i, 2).Value
    Next i

    '篩選ＳＭ重複資料'
    With Application
            ' Turn off screen updating to increase performance
            .ScreenUpdating = True
            'Dim LastColumn As Integer
            LastColumn = Cells.Find(what:="*", after:=Range("C1"), SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
            With Range("C1:C" & Cells(Rows.Count, 1).End(xlUp).Row)
                ' Use AdvanceFilter to filter unique values
                .AdvancedFilter Action:=xlFilterCopy, copytorange:=Worksheets("TempData").Range("L1"), Unique:=True
                '.SpecialCells(xlCellTypeVisible).Offset(0, LastColumn - 1).Value = 1
                On Error Resume Next
                ActiveSheet.ShowAllData
                'Delete the blank rows
                'Columns(LastColumn).SpecialCells(xlCellTypeBlanks).EntireRow.Delete
                Err.Clear
            End With
            'Columns(LastColumn).Clear
            .ScreenUpdating = True
        End With
   
    Worksheets("TempData").Range("J1").Value = "SHopID"
    Worksheets("TempData").Range("K1").Value = "Moedl"
    Worksheets("TempData").Range("M1").Value = "QTY"

    Set srcRange = Worksheets("TempData").Range("A1").CurrentRegion.Columns(3)
    SMCount = Worksheets("TempData").Range("L1").CurrentRegion.Rows.Count

    '列出同型號序號'
    For i = 2 To SMCount Step 1
        Set fndRange = srcRange.Find(what:=Cells(i, 12), LOOKAT:=xlWhole)
        If Not fndRange Is Nothing Then
            fstAddress = fndRange.Address
            'i = 1
            
            TempQTY = 0
            j = 14
            Do
                Worksheets("TempData").Cells(i, 10).Value = fndRange.Offset(, -2).Value
                Worksheets("TempData").Cells(i, 11).Value = fndRange.Offset(, -1).Value
                TempQTY = TempQTY + fndRange.Offset(, 2).Value
                Cells(i, j).Value = fndRange.Offset(, 1).Value
                j = j + 1
                Set fndRange = srcRange.FindNext(after:=fndRange)
                'i = i + 1
            Loop Until fndRange.Address = fstAddress
            Worksheets("TempData").Cells(i, 13).Value = TempQTY
            If TempQTY < 0 Then
                Cells(i, 11).Value = Left(Cells(i, 11), Len(Cells(i, 11)) - 7)
            End If
            
        End If
    Next i

        '計算門市數與門市型號數'

    With Application
            ' Turn off screen updating to increase performance
            .ScreenUpdating = True
            'Dim LastColumn As Integer
            LastColumn = Cells.Find(what:="*", after:=Range("J1"), SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
            With Range("J1:J" & Cells(Rows.Count, 1).End(xlUp).Row)
                ' Use AdvanceFilter to filter unique values
                .AdvancedFilter Action:=xlFilterCopy, copytorange:=Worksheets("TempData").Range("G1"), Unique:=True
                '.SpecialCells(xlCellTypeVisible).Offset(0, LastColumn - 1).Value = 1
                On Error Resume Next
                ActiveSheet.ShowAllData
                'Delete the blank rows
                'Columns(LastColumn).SpecialCells(xlCellTypeBlanks).EntireRow.Delete
                Err.Clear
            End With
            'Columns(LastColumn).Clear
            .ScreenUpdating = True
        End With

        ShopCount = Worksheets("TempData").Range("G1").CurrentRegion.Rows.Count
        Range("H1").Value = ShopCount - 1

        Set srcRange = Worksheets("TempData").Range("J1").CurrentRegion.Columns(1)
        c = 0
        For i = 2 To ShopCount Step 1
            Set fndRange = srcRange.Find(what:=Cells(i, 7))
            CountModel = 0
            If Not fndRange Is Nothing Then
                fstAddress = fndRange.Address
                Do
                    CountModel = CountModel + 1
                    Set fndRange = srcRange.FindNext(after:=fndRange)
                Loop Until fndRange.Address = fstAddress
            End If
            Cells(i, 8).Value = CountModel
        Next i


        sriCount = Range("J1").CurrentRegion.Rows.Count
        Range("N:N").Insert xlShiftToRight


        Set srcRange = Worksheets("APPLE_EDI_SLS").Range("A1").CurrentRegion.Columns(2)
        For i = 2 To sriCount Step 1
            Set fndRange = srcRange.Find(what:=Cells(i, 11))
            If Not fndRange Is Nothing Then
                Cells(i, 14).Value = fndRange.Offset(, 3).Value
            End If
        Next i

'vvvvvvvvvvvvvvvvvvvvvvvvv'
        '製作CSV'

    INVyyyy = Format(d, "yyyy")
    INVyy = Format(d, "yy")
    INVmm = Format(d, "mm")
    INVdd = Format(d, "dd")
    reportData_yyyy = Format(saleDate, "yyyy")
    reportData_mm = Format(saleDate, "mm")
    reportdata_dd = Format(saleDate, "dd")
    'msgbox(d)
    ms = Int(Rnd() * 9000 + 1000)
    SerialNumber = Format(Worksheets("ShopID").Range("I1").Value, "0000")
    
    HostID = Worksheets("ShopID").Range("E1").Value
    HQID = Worksheets("ShopID").Range("F1").Value
    'UNB+UNOA:3+總公司代碼:ZZ+060704780800T:01+年月日yymmdd:隨機四位數+流水號++INVRPT''
    'UNB+UNOA:3+X015715036XXX:ZZ+060704780800T:01+080408:0811+00000000000660++INVRPT''
    ExcelRow = 1
    Worksheets("SLS_CSV").Cells(ExcelRow, 1).Value = "UNB+UNOA:3+" & HostID & ":ZZ+060704780800T:01+" & INVyyyy & INVmm & INVdd & ":" & ms & "+" & SerialNumber & "++SLSRPT'"
    UNH_UNZ = 0
    
    Prod_run = 1
    lastProdCount = 0
    For i = 2 To ShopCount Step 1
        'UNH_UNZ = UNH_UNZ + 1
        totalSriCount = 0
        BGM = Format(Worksheets("ShopID").Range("I2").Value, "0000")
        UNT = 0
        NAD_SHOPID = Worksheets("SHOPID").Cells(i, 2).Value
        ExcelRow = ExcelRow + 1
        Worksheets("SLS_CSV").Cells(ExcelRow, 1).Value = "UNH+" & BGM & "+SLSRPT:D:99A:UN'"
        'UNT = UNT + 1
        ExcelRow = ExcelRow + 1
        Worksheets("SLS_CSV").Cells(ExcelRow, 1).Value = "BGM+190+" & BGM & "+9'"
        ExcelRow = ExcelRow + 1
        '文件日期當天'
        Worksheets("SLS_CSV").Cells(ExcelRow, 1).Value = "DTM+137:" & INVyyyy & INVmm & INVdd & ":102'"
        ExcelRow = ExcelRow + 1
        '報送當天'
        Worksheets("SLS_CSV").Cells(ExcelRow, 1).Value = "DTM+90:" & INVyyyy & INVmm & INVdd & ":102'"
        ExcelRow = ExcelRow + 1
        '報哪一天'
        Worksheets("SLS_CSV").Cells(ExcelRow, 1).Value = "DTM+91:" & reportData_yyyy & reportData_mm & reportdata_dd & ":102'"
        ExcelRow = ExcelRow + 1
        Worksheets("SLS_CSV").Cells(ExcelRow, 1).Value = "NAD+BY+" & HQID & "::91'"
        ExcelRow = ExcelRow + 1
        Worksheets("SLS_CSV").Cells(ExcelRow, 1).Value = "LOC+162+" & NAD_SHOPID & "::90'"
        ProdCount = Worksheets("TempData").Cells(i, 8).Value

        For j = 1 To ProdCount Step 1

            LIN_VP = Worksheets("TEmpData").Cells(j + 1 + lastProdCount, 11).Value
            QTY = Worksheets("TEmpData").Cells(j + 1 + lastProdCount, 13).Value
            ExcelRow = ExcelRow + 1
            Worksheets("SLS_CSV").Cells(ExcelRow, 1).Value = "LIN+" & j & "++" & LIN_VP & ":VP'"
            ExcelRow = ExcelRow + 1
            '銷售日期'
            Worksheets("SLS_CSV").Cells(ExcelRow, 1).Value = "DTM+3:" & 銷售日期 & ":102'"
            ExcelRow = ExcelRow + 1
            Worksheets("SLS_CSV").Cells(ExcelRow, 1).Value = "RFF+IV:INVOICE_" & Format(j, "0000") & "'"
            ExcelRow = ExcelRow + 1
            Worksheets("SLS_CSV").Cells(ExcelRow, 1).Value = "PRI+CAL:" & Cells(j + 1 + lastProdCount, 14).Value & "'"
            For k = 1 To Abs(QTY) Step 1
                ExcelRow = ExcelRow + 1
                Worksheets("SLS_CSV").Cells(ExcelRow, 1).Value = "GIN+BN+" & Right(Cells(j + 1 + lastProdCount, 14 + k).Value, 12) & "'"
            Next k
            ExcelRow = ExcelRow + 1
            If QTY > 0 Then
                Worksheets("SLS_CSV").Cells(ExcelRow, 1).Value = "QTY+153:" & QTY & "'"
            Else
                Worksheets("SLS_CSV").Cells(ExcelRow, 1).Value = "QTY+61:" & Abs(QTY) & "'"
            End If
            totalSriCount = totalSriCount + Abs(QTY)
        Next j
        Worksheets("ShopID").Range("I2").Value = Worksheets("ShopID").Range("I2").Value + 1
        ExcelRow = ExcelRow + 1
        UNT = ProdCount * 5 + 8 + totalSriCount
        Worksheets("SLS_CSV").Cells(ExcelRow, 1).Value = "UNT+" & UNT & "+" & BGM & "'"
        lastProdCount = lastProdCount + ProdCount
        If BGM = 9999 Then
            Worksheets("ShopID").Range("I2").Value = 1
        End If
    Next i
    ExcelRow = ExcelRow + 1
    Worksheets("SLS_CSV").Cells(ExcelRow, 1).Value = "UNZ+" & ShopCount - 1 & "+" & SerialNumber & "'"

    MsgBox ("Done")
End Sub











