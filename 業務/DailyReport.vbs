Sub 每日營業()
    
    TrowsCounts = Worksheets("每日營業").Range("A11").CurrentRegion.Rows.Count
    TlineCounts = Worksheets("每日營業").Range("A11").CurrentRegion.Columns.Count

    Dim TotalAmount, CPUAmount, A3PPAmount, CPU_TQTY, iPad_TQTY, iPadPro_TQTY, iPhone_Amount, Watch_TQTY, SAC_Watch As Long
    Dim iPod_TQTY, ACPP_CPU, SAC_CPU, SAC_iOS, P3PP_Amount, Price_Good, ePen, GPM As Long
    '變數定義
    TotalAmount = 0 ' 總營業額
    CPUAmount = 0 'Apple主機營業額
    A3PPAmount = 0 'Apple配件營業額
    CPU_TQTY = 0 'CPU台數
    iPad_TQTY = 0 'iPad台數
    iPadPro_TQTY = 0 'iPadPro 台數
    iPhone_TQTY = 0 'iPhone台數
    iPhone_Amount = 0 'iPhone營業額
    Watch_TQTY = 0 'watch台數
    SAC_Watch = 0 'SACare for watch
    iPod_TQTY = 0 'iPod 台數
    ACPP_CPU = 0 'ACPP_MAC
    ACPP_IOS = 0 'ACPP for iPhone
    ACPP_W = 0 'ACPP for watch
    ACPP_iPod = 0 'ACPP for iPod
    ACPP_Acc = 0 'ACPP for handphones
    SAC_CPU = 0 'SA Care for CPU
    SAC_iOS = 0 'SA Care for iOS
    P3PP_Amount = 0 '3PP配件營業額
    Price_Good = 0 '時架商品
    Cell_count = 0 '門號數量
    Envelope_Amount = 0 '包膜金額
    cc_ACPP = 0 '代收款
    ePen = 0
    GPM = 0

    TotalAmount = Worksheets("每日營業").Cells(TrowsCounts + 10, TlineCounts).Value
    GPM = Worksheets("每日營業").Cells(2, 8).Value
    'SAC_Watch = Worksheets("每日營業").Cells(3, 9).Value
    ePen = Worksheets("每日營業").Cells(3, 12).Value
    
    Dim tempC2, tempC3, tempC4, tempC6 As String
    Dim Qty, Amt As Double
    
    For i = 11 To TrowsCounts + 10 Step 1
        tempC2 = Worksheets("每日營業").Cells(i, 1).Value
        tempC3 = Worksheets("每日營業").Cells(i, 3).Value
        tempC4 = Worksheets("每日營業").Cells(i, 5).Value
        tempC6 = Worksheets("每日營業").Cells(i, 9).Value
        Qty = Worksheets("每日營業").Cells(i, 11).Value
        Amt = Worksheets("每日營業").Cells(i, 12).Value

        Select Case tempC2
            Case Is = "2003"
                Price_Good = Price_Good + Amt
            Case Is = "2002"
                P3PP_Amount = P3PP_Amount + Amt
                Select Case tempC6
                    Case Is = "6508"
                        SAC_iOS = SAC_iOS + Qty
                    Case Is = "6509"
                        SAC_CPU = SAC_CPU + Qty
                    Case Is = "6522"
                        SAC_Watch = SAC_Watch + Qty
                    Case Is = "6104"
                        Envelope_Amount = Envelope_Amount + Amt
                        'P3PP_Amount = P3PP_Amount - Amt

                End Select
            Case Is = "2016"
                Select Case tempC3
                    Case Is = "3003"
                       P3PP_Amount = P3PP_Amount + Amt
                End Select
            Case Is = "2019"
                Cell_count = Cell_count + Qty
            Case Is = "2028"
                cc_ACPP = cc_ACPP + Amt
                Select Case tempC6
                    Case Is = "6524"
                        ACPP_CPU = ACPP_CPU + Qty
                    Case Is = "6523"
                        ACPP_IOS = ACPP_IOS + Qty
                    Case Is = "6528"
                        ACPP_W = ACPP_W + Qty
                    Case Is = "6527"
                        ACPP_iPod = ACPP_iPod + Qty
                    Case Is = "6529"
                        ACPP_Acc = ACPP_Acc + Qty
                End Select
        End Select
                   
        Select Case tempC3
            Case Is = "3001"
                   CPUAmount = CPUAmount + Amt
            Case Is = "3002"
                   A3PPAmount = A3PPAmount + Amt
            'Case Is = "3003"
                'P3PP_Amount = P3PP_Amount + Amt
        End Select
        
        Select Case tempC4
            Case Is = "4001"
                CPU_TQTY = CPU_TQTY + Qty
            Case Is = "4002"
                CPU_TQTY = CPU_TQTY + Qty
            Case Is = "4003"
                iPod_TQTY = iPod_TQTY + Qty
            Case Is = "4004"
                iPhone_TQTY = iPhone_TQTY + Qty
                 iPhone_Amount = iPhone_Amount + Amt
            Case Is = "4049"
                iPhone_TQTY = iPhone_TQTY + Qty
                 iPhone_Amount = iPhone_Amount + Amt
            Case Is = "4005"
                iPad_TQTY = iPad_TQTY + Qty
            Case Is = "4006"
                iPad_TQTY = iPad_TQTY + Qty
            Case Is = "4038"
                Watch_TQTY = Watch_TQTY + Qty
            Case Is = "4041"
                iPadPro_TQTY = iPadPro_TQTY + Qty
        End Select

        'Select Case tempC6
            'Case Is = "6524"
                'ACPP_CPU = ACPP_CPU + Qty
            'Case Is = "6523"
                'ACPP_IOS = ACPP_IOS + Qty
            'Case Is = "6528"
                'ACPP_W = ACPP_W + Qty
            'Case Is = "6527"
                'ACPP_iPod = ACPP_iPod + Qty
            'Case Is = "6529"
                'ACPP_Acc = ACPP_Acc + Qty
            'Case Is = "6508"
                'SAC_iOS = SAC_iOS + Qty
            'Case Is = "6509"
                'SAC_CPU = SAC_CPU + Qty
        'End Select
    Next i
    'SAC_iOS = SAC_iOS + SAC_Watch
    Worksheets("日報").Cells(7, 3).Value = TotalAmount
    Worksheets("日報").Cells(8, 3).Value = GPM
    Worksheets("日報").Cells(12, 3).Value = CPUAmount
    Worksheets("日報").Cells(13, 3).Value = A3PPAmount
    Worksheets("日報").Cells(14, 3).Value = CPU_TQTY
    Worksheets("日報").Cells(15, 3).Value = iPad_TQTY
    Worksheets("日報").Cells(16, 3).Value = iPadPro_TQTY
    Worksheets("日報").Cells(17, 3).Value = iPhone_TQTY
    Worksheets("日報").Cells(18, 3).Value = iPhone_Amount
    Worksheets("日報").Cells(19, 3).Value = Watch_TQTY
    Worksheets("日報").Cells(20, 3).Value = SAC_Watch
    Worksheets("日報").Cells(22, 3).Value = iPod_TQTY
    Worksheets("日報").Cells(23, 3).Value = ACPP_CPU
    Worksheets("日報").Cells(25, 3).Value = SAC_CPU
    Worksheets("日報").Cells(28, 3).Value = SAC_iOS
    Worksheets("日報").Cells(30, 3).Value = P3PP_Amount
    Worksheets("日報").Cells(34, 3).Value = Price_Good
    Worksheets("日報").Cells(35, 3).Value = Cell_count
    Worksheets("日報").Cells(36, 3).Value = ePen
    Worksheets("日報").Cells(37, 3).Value = Envelope_Amount
    
    MsgBox ("每日營業計算完成")

End Sub
Sub 每月累積()
    
    TrowsCounts = Worksheets("每月累積").Range("A11").CurrentRegion.Rows.Count
    TlineCounts = Worksheets("每月累積").Range("A11").CurrentRegion.Columns.Count

    Dim TotalAmount, CPUAmount, A3PPAmount, CPU_TQTY, iPad_TQTY, iPadPro_TQTY, iPhone_Amount, Watch_TQTY, SAC_Watch, cc_ACPP As Long
    Dim iPod_TQTY, ACPP_CPU, SAC_CPU, SAC_iOS, P3PP_Amount, Price_Good, ePen, GPM, Old_SACW As Long
    '變數定義
    TotalAmount = 0 ' 總營業額
    CPUAmount = 0 'Apple主機營業額
    A3PPAmount = 0 'Apple配件營業額
    CPU_TQTY = 0 'CPU台數
    iPad_TQTY = 0 'iPad台數
    iPadPro_TQTY = 0 'iPadPro 台數
    iPhone_TQTY = 0 'iPhone台數
    iPhone_Amount = 0 'iPhone營業額
    Watch_TQTY = 0 'watch台數
    SAC_Watch = 0 'SACare for watch
    iPod_TQTY = 0 'iPod 台數
    ACPP_CPU = 0 'ACPP_MAC
    ACPP_IOS = 0 'ACPP for iPhone
    ACPP_W = 0 'ACPP for watch
    ACPP_iPod = 0 'ACPP for iPod
    ACPP_Acc = 0 'ACPP for handphones
    SAC_CPU = 0 'SA Care for CPU
    SAC_iOS = 0 'SA Care for iOS
    Old_SACW = 0 'old SA Care for watch
    P3PP_Amount = 0 '3PP配件營業額
    Price_Good = 0 '時架商品
    Cell_count = 0 '門號數量
    Envelope_Amount = 0 '包膜金額
    cc_ACPP = 0 '代收款

    ePen = 0
    GPM = 0

    
    TotalAmount = Worksheets("每月累積").Cells(TrowsCounts + 10, TlineCounts).Value
    GPM = Worksheets("每月累積").Cells(2, 8).Value
    Old_SACW = Worksheets("每月累積").Cells(3, 9).Value
    ePen = Worksheets("每月累積").Cells(3, 12).Value
    
    Dim tempC2, tempC3, tempC4, tempC6 As String
    Dim Qty, Amt As Double
    
    For i = 11 To TrowsCounts + 10 Step 1
        tempC2 = Worksheets("每月累積").Cells(i, 1).Value
        tempC3 = Worksheets("每月累積").Cells(i, 3).Value
        tempC4 = Worksheets("每月累積").Cells(i, 5).Value
        tempC6 = Worksheets("每月累積").Cells(i, 9).Value
        Qty = Worksheets("每月累積").Cells(i, 11).Value
        Amt = Worksheets("每月累積").Cells(i, 12).Value

        Select Case tempC2
            Case Is = "2003"
                Price_Good = Price_Good + Amt
            Case Is = "2002"
                P3PP_Amount = P3PP_Amount + Amt
                Select Case tempC6
                    Case Is = "6508"
                        SAC_iOS = SAC_iOS + Qty
                    Case Is = "6509"
                        SAC_CPU = SAC_CPU + Qty
                    Case Is = "6522"
                        SAC_Watch = SAC_Watch + Qty
                    Case Is = "6104"
                        Envelope_Amount = Envelope_Amount + Amt
                        'P3PP_Amount = P3PP_Amount - Amt
                End Select
            Case Is = "2016"
                Select Case tempC3
                    Case Is = "3003"
                       P3PP_Amount = P3PP_Amount + Amt
                End Select
             Case Is = "2019"
                Cell_count = Cell_count + Qty
             Case Is = "2028"
                 cc_ACPP = cc_ACPP + Amt
                Select Case tempC6
                    Case Is = "6524"
                        ACPP_CPU = ACPP_CPU + Qty
                    Case Is = "6523"
                        ACPP_IOS = ACPP_IOS + Qty
                    Case Is = "6528"
                        ACPP_W = ACPP_W + Qty
                    Case Is = "6527"
                        ACPP_iPod = ACPP_iPod + Qty
                    Case Is = "6529"
                        ACPP_Acc = ACPP_Acc + Qty
                End Select
        End Select
                   
        Select Case tempC3
            Case Is = "3001"
                   CPUAmount = CPUAmount + Amt
            Case Is = "3002"
                   A3PPAmount = A3PPAmount + Amt
            'Case Is = "3003"
                'P3PP_Amount = P3PP_Amount + Amt
        End Select
        
        Select Case tempC4
            Case Is = "4001"
                CPU_TQTY = CPU_TQTY + Qty
            Case Is = "4002"
                CPU_TQTY = CPU_TQTY + Qty
            Case Is = "4003"
                iPod_TQTY = iPod_TQTY + Qty
            Case Is = "4004"
                iPhone_TQTY = iPhone_TQTY + Qty
                 iPhone_Amount = iPhone_Amount + Amt
            Case Is = "4049"
                iPhone_TQTY = iPhone_TQTY + Qty
                 iPhone_Amount = iPhone_Amount + Amt
            Case Is = "4005"
                iPad_TQTY = iPad_TQTY + Qty
            Case Is = "4006"
                iPad_TQTY = iPad_TQTY + Qty
            Case Is = "4038"
                Watch_TQTY = Watch_TQTY + Qty
            Case Is = "4041"
                iPadPro_TQTY = iPadPro_TQTY + Qty
        End Select

       'Select Case tempC6
            'Case Is = "6524"
                'ACPP_CPU = ACPP_CPU + Qty
            'Case Is = "6523"
                'ACPP_IOS = ACPP_IOS + Qty
           ' Case Is = "6528"
                'ACPP_W = ACPP_W + Qty
            'Case Is = "6527"
                'ACPP_iPod = ACPP_iPod + Qty
            'Case Is = "6529"
                'ACPP_Acc = ACPP_Acc + Qty
            'Case Is = "6508"
                'SAC_iOS = SAC_iOS + Qty
            'Case Is = "6509"
                'SAC_CPU = SAC_CPU + Qty
        'End Select
    Next i
    'SAC_iOS = SAC_iOS - Old_SACW
    'SAC_Watch = SAC_Watch + Old_SACW
    Worksheets("日報").Cells(7, 4).Value = TotalAmount
    Worksheets("日報").Cells(8, 4).Value = GPM
    Worksheets("日報").Cells(12, 4).Value = CPUAmount
    Worksheets("日報").Cells(13, 4).Value = A3PPAmount
    Worksheets("日報").Cells(14, 4).Value = CPU_TQTY
    Worksheets("日報").Cells(15, 4).Value = iPad_TQTY
    Worksheets("日報").Cells(16, 4).Value = iPadPro_TQTY
    Worksheets("日報").Cells(17, 4).Value = iPhone_TQTY
    Worksheets("日報").Cells(18, 4).Value = iPhone_Amount
    Worksheets("日報").Cells(19, 4).Value = Watch_TQTY
    Worksheets("日報").Cells(20, 4).Value = SAC_Watch
    Worksheets("日報").Cells(22, 4).Value = iPod_TQTY
    Worksheets("日報").Cells(23, 4).Value = ACPP_CPU
    Worksheets("日報").Cells(25, 4).Value = SAC_CPU
    Worksheets("日報").Cells(28, 4).Value = SAC_iOS
    Worksheets("日報").Cells(30, 4).Value = P3PP_Amount
    Worksheets("日報").Cells(34, 4).Value = Price_Good
    Worksheets("日報").Cells(35, 4).Value = Cell_count
    Worksheets("日報").Cells(36, 4).Value = ePen
    Worksheets("日報").Cells(37, 4).Value = Envelope_Amount

    MsgBox ("每月累積計算完成")

End Sub
Sub 同期比較()
    
    TrowsCounts = Worksheets("同期比較").Range("A11").CurrentRegion.Rows.Count
    TlineCounts = Worksheets("同期比較").Range("A11").CurrentRegion.Columns.Count

    Dim TotalAmount, CPUAmount, A3PPAmount, CPU_TQTY, iPad_TQTY, iPadPro_TQTY, iPhone_Amount, Watch_TQTY, SAC_Watch As Long
    Dim iPod_TQTY, ACPP_CPU, SAC_CPU, SAC_iOS, P3PP_Amount, Price_Good, ePen, Old_SACW, GPM As Long
    '變數定義
    TotalAmount = 0 ' 總營業額
    CPUAmount = 0 'Apple主機營業額
    A3PPAmount = 0 'Apple配件營業額
    CPU_TQTY = 0 'CPU台數
    iPad_TQTY = 0 'iPad台數
    iPadPro_TQTY = 0 'iPadPro 台數
    iPhone_TQTY = 0 'iPhone台數
    iPhone_Amount = 0 'iPhone營業額
    Watch_TQTY = 0 'watch台數
    SAC_Watch = 0 'SACare for watch
    Old_SACW = 0 '舊SACare for watch
    iPod_TQTY = 0 'iPod 台數
    ACPP_CPU = 0 'ACPP_MAC
    ACPP_IOS = 0 'ACPP for iPhone
    ACPP_W = 0 'ACPP for watch
    ACPP_iPod = 0 'ACPP for iPod
    ACPP_Acc = 0 'ACPP for handphones
    SAC_CPU = 0 'SA Care for CPU
    SAC_iOS = 0 'SA Care for iOS
    P3PP_Amount = 0 '3PP配件營業額
    Price_Good = 0 '時架商品
    Cell_count = 0 '門號數量
    Envelope_Amount = 0 '包膜金額
    cc_ACPP = 0 '代收款
    ePen = 0
    GPM = 0
 
    TotalAmount = Worksheets("同期比較").Cells(TrowsCounts + 10, TlineCounts).Value
    GPM = Worksheets("同期比較").Cells(2, 8).Value
    Old_SACW = Worksheets("同期比較").Cells(3, 9).Value
    ePen = Worksheets("同期比較").Cells(3, 12).Value
    
    Dim tempC2, tempC3, tempC4, tempC6 As String
    Dim Qty, Amt As Double
    
    For i = 11 To TrowsCounts + 10 Step 1
        tempC2 = Worksheets("同期比較").Cells(i, 1).Value
        tempC3 = Worksheets("同期比較").Cells(i, 3).Value
        tempC4 = Worksheets("同期比較").Cells(i, 5).Value
        tempC6 = Worksheets("同期比較").Cells(i, 9).Value
        Qty = Worksheets("同期比較").Cells(i, 11).Value
        Amt = Worksheets("同期比較").Cells(i, 12).Value

        Select Case tempC2
            Case Is = "2003"
                Price_Good = Price_Good + Amt
            Case Is = "2002"
                P3PP_Amount = P3PP_Amount + Amt
                Select Case tempC6
                    Case Is = "6508"
                        SAC_iOS = SAC_iOS + Qty
                    Case Is = "6509"
                        SAC_CPU = SAC_CPU + Qty
                    Case Is = "6522"
                        SAC_Watch = SAC_Watch + Qty
                    Case Is = "6104"
                        Envelope_Amount = Envelope_Amount + Amt
                        'P3PP_Amount = P3PP_Amount - Amt
                End Select
            Case Is = "2016"
                Select Case tempC3
                    Case Is = "3003"
                       P3PP_Amount = P3PP_Amount + Amt
                End Select
             Case Is = "2019"
                Cell_count = Cell_count + Qty
             Case Is = "2028"
                cc_ACPP = cc_ACPP + Amt
                Select Case tempC6
                    Case Is = "6524"
                        ACPP_CPU = ACPP_CPU + Qty
                    Case Is = "6523"
                        ACPP_IOS = ACPP_IOS + Qty
                    Case Is = "6528"
                        ACPP_W = ACPP_W + Qty
                    Case Is = "6527"
                        ACPP_iPod = ACPP_iPod + Qty
                    Case Is = "6529"
                        ACPP_Acc = ACPP_Acc + Qty
                End Select
        End Select
                   
        Select Case tempC3
            Case Is = "3001"
                   CPUAmount = CPUAmount + Amt
            Case Is = "3002"
                   A3PPAmount = A3PPAmount + Amt
            'Case Is = "3003"
                'P3PP_Amount = P3PP_Amount + Amt
        End Select
        
        Select Case tempC4
            Case Is = "4001"
                CPU_TQTY = CPU_TQTY + Qty
            Case Is = "4002"
                CPU_TQTY = CPU_TQTY + Qty
            Case Is = "4003"
                iPod_TQTY = iPod_TQTY + Qty
            Case Is = "4004"
                iPhone_TQTY = iPhone_TQTY + Qty
                 iPhone_Amount = iPhone_Amount + Amt
            Case Is = "4049"
                iPhone_TQTY = iPhone_TQTY + Qty
                 iPhone_Amount = iPhone_Amount + Amt
            Case Is = "4005"
                iPad_TQTY = iPad_TQTY + Qty
            Case Is = "4006"
                iPad_TQTY = iPad_TQTY + Qty
            Case Is = "4038"
                Watch_TQTY = Watch_TQTY + Qty
            Case Is = "4041"
                iPadPro_TQTY = iPadPro_TQTY + Qty
        End Select

        'Select Case tempC6
            'Case Is = "6524"
                'ACPP_CPU = ACPP_CPU + Qty
            'Case Is = "6523"
               ' ACPP_IOS = ACPP_IOS + Qty
            'Case Is = "6528"
                'ACPP_W = ACPP_W + Qty
            'Case Is = "6527"
                'ACPP_iPod = ACPP_iPod + Qty
            'Case Is = "6529"
                'ACPP_Acc = ACPP_Acc + Qty
            'Case Is = "6508"
                'SAC_iOS = SAC_iOS + Qty
            'Case Is = "6509"
                'SAC_CPU = SAC_CPU + Qty
        'End Select
    Next i
   'SAC_iOS = SAC_iOS - Old_SACW
   ' SAC_Watch = SAC_Watch + Old_SACW
    Worksheets("日報").Cells(7, 7).Value = TotalAmount
    Worksheets("日報").Cells(8, 7).Value = GPM
    Worksheets("日報").Cells(12, 7).Value = CPUAmount
    Worksheets("日報").Cells(13, 7).Value = A3PPAmount
    Worksheets("日報").Cells(14, 7).Value = CPU_TQTY
    Worksheets("日報").Cells(15, 7).Value = iPad_TQTY
    Worksheets("日報").Cells(16, 7).Value = iPadPro_TQTY
    Worksheets("日報").Cells(17, 7).Value = iPhone_TQTY
    Worksheets("日報").Cells(18, 7).Value = iPhone_Amount
    Worksheets("日報").Cells(19, 7).Value = Watch_TQTY
    Worksheets("日報").Cells(20, 7).Value = SAC_Watch
    Worksheets("日報").Cells(22, 7).Value = iPod_TQTY
    Worksheets("日報").Cells(23, 7).Value = ACPP_CPU
    Worksheets("日報").Cells(25, 7).Value = SAC_CPU
    Worksheets("日報").Cells(28, 7).Value = SAC_iOS
    Worksheets("日報").Cells(30, 7).Value = P3PP_Amount
    Worksheets("日報").Cells(34, 7).Value = Price_Good
    Worksheets("日報").Cells(35, 7).Value = Cell_count
    Worksheets("日報").Cells(36, 7).Value = ePen
    Worksheets("日報").Cells(37, 7).Value = Envelope_Amount

    
    MsgBox ("同期比較計算完成")

End Sub

Sub cleardaily()
    Worksheets("每日營業").Range(Cells(11, 1), Cells(999, 999)).Clear
End Sub
Sub clearmonthly()
    Worksheets("每月累積").Range(Cells(11, 1), Cells(999, 999)).Clear
End Sub
Sub clearsamemonth()
    Worksheets("同期比較").Range(Cells(11, 1), Cells(999, 999)).Clear
End Sub

Sub ClearSACare()

    Worksheets("SAcare").Range(Cells(10, 1), Cells(999, 999)).Clear

End Sub
Sub sacaremoney() '每日營業
    
    Worksheets("SAcare").Range(Cells(10, 6), Cells(999, 999)).Clear

    SArowsCounts = Worksheets("SAcare").Range("A10").CurrentRegion.Rows.Count
    SAlineCounts = Worksheets("SAcare").Range("A10").CurrentRegion.Columns.Count
    
    Dim SACtotal, SACQua, SACPrice As Long
    

    SACQua = 0
    SACPrice = 0
    SACtotal = 0

    For i = 10 To SArowsCounts + 8 Step 1
        SACQua = Cells(i, 5).Value
        Cells(i, 6).Value = "=iferror(VLOOKUP(A" & i & ",SACare價目表!A:C,3,false),0)"
        SACPrice = Cells(i, 6).Value
        Cells(i, 7).Value = Cells(i, 5) * Cells(i, 6)

        SACtotal = Cells(i, 7) + SACtotal

    Next i

    Worksheets("日報").Cells(7, 3).Value = Worksheets("日報").Cells(7, 3).Value + SACtotal
    'Worksheets("日報").Cells(30, 3).Value = Worksheets("日報").Cells(30, 3).Value + SACtotal
    Worksheets("日報").Cells(32, 3).Value =  SACtotal
    Worksheets("日報").Cells(8, 3).Value = Worksheets("日報").Cells(8, 3).Value + SACtotal * 0.45

    MsgBox ("完成")
End Sub

Sub sacaremoney2() '每月累積

    
    Worksheets("SAcare").Range(Cells(10, 6), Cells(999, 999)).Clear

    SArowsCounts = Worksheets("SAcare").Range("A10").CurrentRegion.Rows.Count
    SAlineCounts = Worksheets("SAcare").Range("A10").CurrentRegion.Columns.Count
    
    Dim SACtotal, SACQua, SACPrice As Long
    

    SACQua = 0
    SACPrice = 0
    SACtotal = 0

    For i = 10 To SArowsCounts + 8 Step 1
        SACQua = Cells(i, 5).Value
        Cells(i, 6).Value = "=iferror(VLOOKUP(A" & i & ",SACare價目表!A:C,3,false),0)"
        SACPrice = Cells(i, 6).Value
        Cells(i, 7).Value = Cells(i, 5) * Cells(i, 6)

        SACtotal = Cells(i, 7) + SACtotal

    Next i

    Worksheets("日報").Cells(7, 4).Value = Worksheets("日報").Cells(7, 4).Value + SACtotal
    'Worksheets("日報").Cells(30, 4).Value = Worksheets("日報").Cells(30, 4).Value + SACtotal
    Worksheets("日報").Cells(32, 4).Value =  SACtotal
    Worksheets("日報").Cells(8, 4).Value = Worksheets("日報").Cells(8, 4).Value + SACtotal * 0.45
    MsgBox ("完成")
End Sub

Sub sacaremoney3() '同期比較

    
    Worksheets("SAcare").Range(Cells(10, 6), Cells(999, 999)).Clear

    SArowsCounts = Worksheets("SAcare").Range("A10").CurrentRegion.Rows.Count
    SAlineCounts = Worksheets("SAcare").Range("A10").CurrentRegion.Columns.Count
    
    Dim SACtotal, SACQua, SACPrice As Long
    

    SACQua = 0
    SACPrice = 0
    SACtotal = 0

    For i = 10 To SArowsCounts + 8 Step 1
        SACQua = Cells(i, 5).Value
        Cells(i, 6).Value = "=iferror(VLOOKUP(A" & i & ",SACare價目表!A:C,3,false),0)"
        SACPrice = Cells(i, 6).Value
        Cells(i, 7).Value = Cells(i, 5) * Cells(i, 6)

        SACtotal = Cells(i, 7) + SACtotal

    Next i

    
    Worksheets("日報").Cells(7, 7).Value = Worksheets("日報").Cells(7, 7).Value + SACtotal
    'Worksheets("日報").Cells(30, 7).Value = Worksheets("日報").Cells(30, 7).Value + SACtotal
    Worksheets("日報").Cells(32, 7).Value =  SACtotal
    Worksheets("日報").Cells(8, 7).Value = Worksheets("日報").Cells(8, 7).Value + SACtotal * 0.45
    MsgBox ("完成")
End Sub


Sub clearTemp()
Worksheets("每日營業").Select
Range("A11:ZZ65536").EntireRow.Delete
ActiveSheet.UsedRange
Worksheets("同期比較").Select
Range("A11:ZZ65536").EntireRow.Delete
ActiveSheet.UsedRange
Worksheets("每月累積").Select
Range("A11:ZZ65536").EntireRow.Delete
ActiveSheet.UsedRange
Worksheets("日報").Select
MsgBox ("完成")
End Sub


