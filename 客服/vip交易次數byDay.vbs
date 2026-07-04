Sub VipTransactionsByDay()

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.EnableEvents = False

    mTime = Time()
    oldTime = Time()

    worksheets("Data").range("E:E").clear
    'Worksheets("Single").UsedRange.ClearContents
    'Worksheets("Multiple").UsedRange.ClearContents

    Datarows = Worksheets("Data").Range("A1").CurrentRegion.Rows.Count
    DataLines = Worksheets("Data").Range("A1").CurrentRegion.Columns.Count
    
    Worksheets("Data").Select

    'Range("A:R").Sort key1:=Range("G1"), order1:=xlAscending, Header:=xlYes

    Dim TransactionsData ()
    dim nodupVipID ()

    Dim i As long
    'Dim Singletable, MultipTable As Variant
    reDim TransactionsData(Datarows,5) 
    
    for i = 2 to Datarows step 1
        'TransactionsData(i,0) = worksheets("Data").cells(i,3).value
        TransactionsData(i,1) = worksheets("Data").cells(i,1).value
        TransactionsData(i,2) = worksheets("Data").cells(i,2).value
        TransactionsData(i,3) = worksheets("Data").cells(i,3).value
        TransactionsData(i,4) = worksheets("Data").cells(i,1).value & worksheets("Data").cells(i,2).value
        If TransactionsData(i,3) = "銷售" or TransactionsData(i,3) = "尾款" Then
            TransactionsData(i,0) = 1
        elseif TransactionsData(i,3) = "訂金" or TransactionsData(i,3) = "退訂" then
            TransactionsData(i,0) = 0
        Else
            TransactionsData(i,0) = -1
        End If
        'TransactionsData(i,0) = worksheets("Data").cells(i,5).value
    next i
    'msgbox TransactionsData(13,0)

    for j = 2 to Datarows step 1
        if TransactionsData(j,3) = "尾款" then
            for k = 2 to Datarows step 1
                if TransactionsData(j,4) = TransactionsData(k,4) and TransactionsData(k,3) = "銷售" then
                    TransactionsData(j,0) = 0
                end if
            next k
         end if   
         worksheets("Data").cells(j,5).value = TransactionsData(j,0) 
    next j


    Columns("A:A").Select
    Selection.Copy
    Columns("I:I").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveSheet.Range("$I:$I").RemoveDuplicates Columns:=1, Header:= _
        xlYes

    noDupDatarows = Worksheets("Data").Range("I1").CurrentRegion.Rows.Count
    nodupVipID = Worksheets("Data").Range("I:I").Value

    for i = 2 to noDupDatarows step 1
        for j = 2 to Datarows step 1
            if nodupVipID(i,1) = TransactionsData(j,1) then
                Worksheets("Data").cells(i,10).value = Worksheets("Data").cells(i,10).value + TransactionsData(j,0)
            end if
        next j
    next i
    msgbox "Done"
end sub
    exceltable = Worksheets("Data").Range(Cells(1, 1), Cells(Datarows, DataLines)).Value
    

    j = 1
    flagX = 0
    k = 2
    f = 1
    stateReUpgrade = 0
    stateUpgrade = 0
    stateDown = 0
    stateReNew = 0
    stateRollback = 0
    ReDim DynRepIDarray(f)
    
    '檢查資料重複(VIPID)，用Data表單去產生single跟Multiple'
    For i = 2 To Datarows Step 1 'Data迴圈 i'
        vipID = exceltable(i, 1)
        levelState = exceltable(i, 3)
        if len(exceltable(i,5)) = 2 or left(exceltable(i,5),3) = "ST0"  then

            For j = 2 To Datarows Step 1 'Data第二次迴圈 j'
                If vipID = exceltable(j, 1) Then
                    flagX = flagX + 1 '如果vipID重複，標記+1'
                End If
                If flagX = 2 Then '標記如果兩筆代表重複'
                    For l = 1 To f Step 1   '檢查重複的名單內是否已經有資料 l'
                        If vipID = Worksheets("Multiple").Cells(l, 1).Value Then
                            Exit For '如果i迴圈與重複名單符合跳出檢查名單迴圈'
                        ElseIf l = f Then '檢查到最後一筆如果沒有重複,往下新增一筆在Multiple名單下'
                            Worksheets("Multiple").Cells(l + 1, 1).Value = vipID
                            DynRepIDarray(f) = vipID '寫入有重複的VIPID到陣列中'
                            f = f + 1
                            ReDim Preserve DynRepIDarray(f) '重新定義陣列長度'
                        Else
                        End If
                    Next l
                    Exit For
                End If
            Next j
            '如果標記只有一筆，寫入single表格'

            If flagX = 1 Then
                Worksheets("Single").Cells(k, 1).Value = vipID
                Worksheets("Single").Cells(k, 2).Value = levelState
                Worksheets("single").Cells(k, 3).Value = exceltable(i, 4) '備註'
                Worksheets("single").Cells(k, 4).NumberFormatLocal = "@"
                Worksheets("single").Cells(k, 4).Value = Format(right(exceltable(i, 5), 2), "00") '舊等級'

                Worksheets("single").Cells(k, 5).NumberFormatLocal = "@"
                Worksheets("single").Cells(k, 5).Value = Format(right(exceltable(i, 6), 2), "00") '新等級'
                If InStr(1, exceltable(i, 4), "Top") > 0 Then   '判斷狀態如果是回升就將UPGRADE改成REUPGRADE'
                    Worksheets("Single").Cells(k, 2).Value = "REUPGRADE"
                'ElseIf levelState = "UPGRADE" Then
                '   stateUpgrade = stateUpgrade + 1
                'elseif levelState = "DOWNGRADE" then
                '    stateDown =stateDown + 1
                'elseif levelState = "RENEW" then
                '    stateReNew = stateReNew + 1
                'elseif levelState = "ROLLBACK" then
                '    stateRollback = stateRollback + 1
                End If
                k = k + 1 'single筆數'
            End If
        end if
    flagX = 0
    Next i

    '單一名單的標頭'
    Worksheets("Single").Cells(1, 1).Value = "VIPID"
    Worksheets("Single").Cells(1, 2).Value = "狀態"
    Worksheets("Single").Cells(1, 3).Value = "備註"
    Worksheets("Single").Cells(1, 4).Value = "舊等級"
    Worksheets("Single").Cells(1, 5).Value = "新等級"

    '重複名單的標頭'
    Worksheets("Multiple").Cells(1, 1).Value = "VIPID"
    Worksheets("Multiple").Cells(1, 2).Value = "第一次變動原因"
    Worksheets("Multiple").Cells(1, 3).Value = "第一次舊等級"
    Worksheets("Multiple").Cells(1, 4).Value = "第一次新等級"
    Worksheets("Multiple").Cells(1, 5).Value = "第一次變動備註"

    Worksheets("Multiple").Cells(1, 6).Value = "最後變動原因"
    Worksheets("Multiple").Cells(1, 7).Value = "最後舊等級"
    Worksheets("Multiple").Cells(1, 8).Value = "最後新等級"
    Worksheets("Multiple").Cells(1, 9).Value = "最後變動備註"

    Worksheets("Multiple").Cells(1, 10).Value = "判斷狀態"
    Worksheets("Multiple").Cells(1, 11).Value = "最終狀態變動"

    '檢查多筆紀錄的會員並寫入(用Multiple檢查Data並抓出第一筆跟最後一筆'
    For i = 1 To UBound(DynRepIDarray) Step 1
        flagX = 1
        For j = 1 To Datarows Step 1
            '檢查如果與重複相同第一筆標記'
            If DynRepIDarray(i) = exceltable(j, 1) And flagX = 1 Then
                If InStr(1, exceltable(j, 4), "Top") > 0 Then '第一次變動原因，狀態是回升的變更為REUPGRADE'
                    Worksheets("Multiple").Cells(i + 1, 2).Value = "REUPGRADE"
                Else
                    Worksheets("Multiple").Cells(i + 1, 2).Value = exceltable(j, 3)
                End If
                '表頭，並將等級格式變更為文字'
                Worksheets("Multiple").Cells(i + 1, 3).NumberFormatLocal = "@"
                Worksheets("Multiple").Cells(i + 1, 3).Value = Format(right(exceltable(j, 5), 2), "00") '第一次的舊等級'
                Worksheets("Multiple").Cells(i + 1, 4).NumberFormatLocal = "@"
                Worksheets("Multiple").Cells(i + 1, 4).Value = Format(right(exceltable(j, 6), 2), "00") '第一次的新等級'

                Worksheets("Multiple").Cells(i + 1, 5).Value = exceltable(j, 4) '第一次變動備註'
                flagX = 2 '標記為已有第一筆'

                '檢查重複第二筆以上複寫欄位上去,只留最後一筆'
            ElseIf DynRepIDarray(i) = exceltable(j, 1) And flagX = 2 Then
                If InStr(1, exceltable(j, 4), "Top") > 0 Then '第一次之後的變動原因，狀態是回升的變更為REUPGRADE'
                    Worksheets("Multiple").Cells(i + 1, 6).Value = "REUPGRADE"
                Else
                    Worksheets("Multiple").Cells(i + 1, 6).Value = exceltable(j, 3)
                End If
                Worksheets("Multiple").Cells(i + 1, 7).NumberFormatLocal = "@"
                Worksheets("Multiple").Cells(i + 1, 7).Value = Format(right(exceltable(j, 5), 2), "00") '最後的舊等級'
                Worksheets("Multiple").Cells(i + 1, 8).NumberFormatLocal = "@"
                Worksheets("Multiple").Cells(i + 1, 8).Value = Format(right(exceltable(j, 6), 2), "00") '最後的新等級'
                Worksheets("Multiple").Cells(i + 1, 9).Value = exceltable(j, 4) '最後變動備註'
            End If
        Next j
    Next i

    Multiplerows = Worksheets("Multiple").Range("A1").CurrentRegion.Rows.Count
    MultipleLines = Worksheets("Multiple").Range("A1").CurrentRegion.Columns.Count

    Worksheets("Multiple").Select
    MultipTable = Worksheets("Multiple").Range(Cells(1, 1), Cells(Multiplerows, MultipleLines)).Value

    For i = 2 To Multiplerows Step 1 '確認最後等級與變動前等級差異，判斷等級變動'
        If MultipTable(i, 3) < MultipTable(i, 8) Then
            Worksheets("Multiple").Cells(i, 10).Value = "UP"    '變動前等級小於最後等級標記UP'
                If MultipTable(i, 2) = "UPGRADE" Or MultipTable(i, 6) = "UPGRADE" Then  '判斷最初跟最後變動狀態'
                    Worksheets("Multiple").Cells(i, 11).Value = "UPGRADE"
                ElseIf MultipTable(i, 2) = "REUPGRADE" And MultipTable(i, 6) = "REUPGRADE" Then
                    Worksheets("Multiple").Cells(i, 11).Value = "REUPGRADE"
                Else
                End If
        ElseIf MultipTable(i, 3) > MultipTable(i, 8) Then
            Worksheets("Multiple").Cells(i, 10).Value = "DOWN" '變動前等級大於最後等級標記DOWN'
            If MultipTable(i, 2) = "DOWNGRADE" Or MultipTable(i, 6) = "DOWNGRADE" Then
                Worksheets("Multiple").Cells(i, 11).Value = "DOWNGRADE"
            End If
        Else
            Worksheets("Multiple").Cells(i, 10).Value = "REMAIN"    '變動前等級等於最後等級標記REMAIN'
            If MultipTable(i, 2) = "DOWNGRADE" And MultipTable(i, 6) = "REUPGRADE" Then
                Worksheets("Multiple").Cells(i, 11).Value = "REUPGRADE"
            ElseIf MultipTable(i, 2) = "RENEW" And MultipTable(i, 6) = "RENEW" Then
                Worksheets("Multiple").Cells(i, 11).Value = "RENEW"
            ElseIf (MultipTable(i, 2) = "REUPGRADE" And MultipTable(i, 6) = "ROLLBACK") Or MultipTable(i, 2) = "ROLLBACK" Then
                Worksheets("Multiple").Cells(i, 11).Value = "UNCHANGE"
            ElseIf MultipTable(i, 2) = "UPGRADE" And MultipTable(i, 6) = "ROLLBACK" Then
                Worksheets("Multiple").Cells(i, 11).Value = "UNCHANGE"
            End If
        End If
    Next i

    'single表單統計'
    SingleCountLevel

    'Multiple表單統計'
    MultipleCountLevel

    '清除動態陣列'
    Erase DynRepIDarray

    '將結果資料填入START頁面'
    Worksheets("START").select
    worksheets("START").cells(9,8).value = UPto03 '升等到鑽石'
    worksheets("START").cells(10,8).value = UPto02 '升等到白金'
    worksheets("START").cells(11,8).value = REto03 '續等到鑽石'
    worksheets("START").cells(14,8).value = REto02 '續等到白金'
    worksheets("START").cells(17,8).value = DOWNfrom03 '原鑽石降等'
    worksheets("START").cells(18,8).value = DOWNfrom02 '原白金降等'

    'MsgBox " ReUpgrade= " & stateReUpgrade & ", Upgrade= " & stateUpgrade & ", REWNEW= " & stateReNew & ", DOWNGRADE= " & stateDown & ", ROLLBACK= " & stateRollback
    

    

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.EnableEvents = True


    mTime = Time() - oldTime
    mSec = Second(mTime)

    'MsgBox "開始時間:" & oldTime & "  結束時間：" & Time()
    'MsgBox ("處理完成" & "花費時間" & mSec & "秒")
    msgbox "處理完成"
    if UncheckMember > 0 then
        msgbox UncheckMember & " 筆，無法確認的會員。再請確認升等狀態"
    end if
End Sub

Sub SingleCountLevel()
    
    '未重複會員統計'

    'mTime = Time
    Worksheets("Single").Select
     
    'Dim stateCount(15)
    Dim Singletable()

    UPGRADE01 = 0
    REUPGRADE01 = 0
    DOWNGRADE01 = 0
    RENEW01 = 0
    ROLLBACK01 = 0

    UPGRADE02 = 0
    REUPGRADE02 = 0
    DOWNGRADE02 = 0
    RENEW02 = 0
    ROLLBACK02 = 0

    UPGRADE03 = 0
    REUPGRADE03 = 0
    DOWNGRADE03 = 0
    RENEW03 = 0
    ROLLBACK03 = 0

    SingleRows = Worksheets("Single").Range("A1").CurrentRegion.Rows.Count
    SingleLines = Worksheets("Single").Range("A1").CurrentRegion.Columns.Count
    Singletable = Worksheets("Single").Range(Cells(1, 1), Cells(SingleRows, SingleLines)).Value


    Dim UP00to01, UP00to02, UP00to03, UP01to02, UP01to03, UP02to03 As Long
    Dim DOWN01to00, DOWN02to00, DOWN02to01, DOWN03to00, DOWN03to02 As Long
    Dim RENEWto01,RENEWto02, RENEWto03, REUPGRADE01to02, REUPGRADE01to03, REUPGRADE02to03 As Long

    UP00to01 = 0
    UP00to02 = 0
    UP00to03 = 0
    UP01to02 = 0
    UP01to03 = 0
    UP02to03 = 0
    DOWN01to00 = 0
    DOWN02to00 = 0
    DOWN02to01 = 0
    DOWN03to00 = 0
    DOWN03to02 = 0
    RENEWto01 = 0
    RENEWto02 = 0
    RENEWto03 = 0
    REUPGRADE01to02 = 0
    REUPGRADE01to03 = 0
    REUPGRADE02to03 = 0

    '先依照等級變動原因判斷，再以新舊等級判斷'
    For i = 2 To SingleRows Step 1
        If Singletable(i, 2) = "UPGRADE" Then
            If Singletable(i, 4) = "00" And Singletable(i, 5) = "01" Then
                UP00to01 = UP00to01 + 1
            ElseIf Singletable(i, 4) = "00" And Singletable(i, 5) = "02" Then
                UP00to02 = UP00to02 + 1
            ElseIf Singletable(i, 4) = "00" And Singletable(i, 5) = "03" Then
                UP00to03 = UP00to03 + 1
            ElseIf Singletable(i, 4) = "01" And Singletable(i, 5) = "02" Then
                UP01to02 = UP01to02 + 1
            ElseIf Singletable(i, 4) = "01" And Singletable(i, 5) = "03" Then
                UP01to03 = UP01to03 + 1
            ElseIf Singletable(i, 4) = "02" And Singletable(i, 5) = "03" Then
                UP02to03 = UP02to03 + 1
            End If
        ElseIf Singletable(i, 2) = "DOWNGRADE" Then
            If Singletable(i, 4) = "01" And Singletable(i, 5) = "00" Then
                DOWN01to00 = DOWN01to00 + 1
            ElseIf Singletable(i, 4) = "02" And Singletable(i, 5) = "00" Then
                DOWN02to00 = DOWN02to00 + 1
            ElseIf Singletable(i, 4) = "02" And Singletable(i, 5) = "01" Then
                DOWN02to01 = DOWN02to01 + 1
            ElseIf Singletable(i, 4) = "03" And Singletable(i, 5) = "00" Then
                DOWN03to00 = DOWN03to00 + 1
            ElseIf Singletable(i, 4) = "03" And Singletable(i, 5) = "01" Then
                DOWN03to01 = DOWN03to01 + 1
            ElseIf Singletable(i, 4) = "03" And Singletable(i, 5) = "02" Then
                DOWN03to02 = DOWN03to02 + 1
            End If
        ElseIf Singletable(i, 2) = "RENEW" Then
            If Singletable(i, 4) = "02" Then
                RENEWto02 = RENEWto02 + 1
            ElseIf Singletable(i, 4) = "03" Then
                RENEWto03 = RENEWto03 + 1
            elseif Singletable(i,4) = "01" then
                RENEWto01 = RENEWto01 + 1
            End If
        ElseIf Singletable(i, 2) = "REUPGRADE" Then
            If Singletable(i, 4) = "01" And Singletable(i, 5) = "02" Then
                REUPGRADE01to02 = REUPGRADE01to02 + 1
            ElseIf Singletable(i, 4) = "01" And Singletable(i, 5) = "03" Then
                REUPGRADE01to03 = REUPGRADE01to03 + 1
            ElseIf Singletable(i, 4) = "02" And Singletable(i, 5) = "03" Then
                REUPGRADE02to03 = REUPGRADE02to03 + 1
            End If
        End If
    Next i

    UPto03 = UPto03 + UP00to03 + UP01to03 + UP02to03
    UPto02 = UPto02 + UP00to02 + UP01to02 
    REto03 = REto03 + RENEWto03 + REUPGRADE01to03 + REUPGRADE02to03
    REto02 = REto02 + RENEWto02 + REUPGRADE01to02
    DOWNfrom03 = DOWNfrom03 + DOWN03to00 + DOWN03to01 + DOWN03to02
    DOWNfrom02 = DOWNfrom02 + DOWN02to00 + DOWN02to01 

    x = 0
    y = 0
    For i = 1 To 16 Step 1
        Worksheets("single").Cells(i + 1, 6).NumberFormatLocal = "@"
        Worksheets("single").Cells(i + 1, 6).Value = Format(x, "00")
        If i Mod 4 = 0 Then
            x = x + 1
        End If
        Worksheets("single").Cells(i + 1, 7).NumberFormatLocal = "@"
        Worksheets("single").Cells(i + 1, 7).Value = Format(y, "00")
        y = y + 1
        If y = 4 Then
            y = 0
        End If
    Next i

    Worksheets("single").Cells(1, 6).Value = "OLD LEVEL"
    Worksheets("single").Cells(1, 7).Value = "NEW LEVEL"
    Worksheets("single").Cells(1, 8).Value = "UPGRADE"
    Worksheets("single").Cells(1, 9).Value = "DOWNGRADE"
    Worksheets("single").Cells(1, 10).Value = "RENEW"
    Worksheets("single").Cells(1, 11).Value = "REUPGRADE"


    Worksheets("single").Cells(3, 8).Value = UP00to01
    Worksheets("single").Cells(4, 8).Value = UP00to02
    Worksheets("single").Cells(5, 8).Value = UP00to03

    Worksheets("single").Cells(8, 8).Value = UP01to02
    Worksheets("single").Cells(9, 8).Value = UP01to03

    Worksheets("single").Cells(13, 8).Value = UP02to03

    Worksheets("single").Cells(6, 9).Value = DOWN01to00

    Worksheets("single").Cells(10, 9).Value = DOWN02to00
    Worksheets("single").Cells(11, 9).Value = DOWN02to01

    Worksheets("single").Cells(14, 9).Value = DOWN03to00
    Worksheets("single").Cells(15, 9).Value = DOWN03to01
    Worksheets("single").Cells(16, 9).Value = DOWN03to02

    Worksheets("single").Cells(7, 10).Value = RENEWto01
    Worksheets("single").Cells(12, 10).Value = RENEWto02
    Worksheets("single").Cells(17, 10).Value = RENEWto03

    Worksheets("single").Cells(4, 11).Value = REUPGRADE00to02
    Worksheets("single").Cells(5, 11).Value = REUPGRADE00to03

    Worksheets("single").Cells(8, 11).Value = REUPGRADE01to02
    Worksheets("single").Cells(9, 11).Value = REUPGRADE01to03

    Worksheets("single").Cells(13, 11).Value = REUPGRADE02to03

'統計二'

    For i = 2 To SingleRows Step 1
        If Singletable(i, 5) = "01" Then
            If Singletable(i, 2) = "UPGRADE" Then
                UPGRADE01 = UPGRADE01 + 1
            ElseIf Singletable(i, 2) = "REUPGRADE" Then
                REUPGRADE01 = REUPGRADE01 + 1
            ElseIf Singletable(i, 2) = "DOWNGRADE" Then
                DOWNGRADE01 = DOWNGRADE01 + 1
            ElseIf Singletable(i, 2) = "RENEW" Then
                RENEW01 = RENEW01 + 1
            ElseIf Singletable(i, 2) = "ROLLBACK" Then
                ROLLBACK01 = ROLLBAXCK01 + 1
            End If
        ElseIf Singletable(i, 5) = "02" Then
            If Singletable(i, 2) = "UPGRADE" Then
                UPGRADE02 = UPGRADE02 + 1
            ElseIf Singletable(i, 2) = "REUPGRADE" Then
                REUPGRADE02 = REUPGRADE02 + 1
            ElseIf Singletable(i, 2) = "DOWNGRADE" Then
                DOWNGRADE02 = DOWNGRADE02 + 1
            ElseIf Singletable(i, 2) = "RENEW" Then
                RENEW02 = RENEW02 + 1
            ElseIf Singletable(i, 2) = "ROLLBACK" Then
                ROLLBACK02 = ROLLBAXCK02 + 1
            End If
        ElseIf Singletable(i, 5) = "03" Then
            If Singletable(i, 2) = "UPGRADE" Then
                UPGRADE03 = UPGRADE03 + 1
            ElseIf Singletable(i, 2) = "REUPGRADE" Then
                REUPGRADE03 = REUPGRADE03 + 1
            ElseIf Singletable(i, 2) = "DOWNGRADE" Then
                DOWNGRADE03 = DOWNGRADE03 + 1
            ElseIf Singletable(i, 2) = "RENEW" Then
                RENEW03 = RENEW03 + 1
            ElseIf Singletable(i, 2) = "ROLLBACK" Then
                ROLLBACK03 = ROLLBAXCK03 + 1
            End If

        End If
    Next i

    x = 20
    Worksheets("Single").Cells(1 + x, 6).Value = "UPGRADE01"
    Worksheets("Single").Cells(2 + x, 6).Value = "REUPGRADE01"
    Worksheets("Single").Cells(3 + x, 6).Value = "DOWNGRADE01"
    Worksheets("Single").Cells(4 + x, 6).Value = "RENEW01"
    Worksheets("Single").Cells(5 + x, 6).Value = "ROLLBACK01"
    Worksheets("Single").Cells(1 + x, 7).Value = UPGRADE01
    Worksheets("Single").Cells(2 + x, 7).Value = REUPGRADE01
    Worksheets("Single").Cells(3 + x, 7).Value = DOWNGRADE01
    Worksheets("Single").Cells(4 + x, 7).Value = RENEW01
    Worksheets("Single").Cells(5 + x, 7).Value = ROLLBACK01

    Worksheets("Single").Cells(6 + x, 6).Value = "UPGRADE02"
    Worksheets("Single").Cells(7 + x, 6).Value = "REUPGRADE02"
    Worksheets("Single").Cells(8 + x, 6).Value = "DOWNGRADE02"
    Worksheets("Single").Cells(9 + x, 6).Value = "RENEW02"
    Worksheets("Single").Cells(10 + x, 6).Value = "ROLLBACK02"
    Worksheets("Single").Cells(6 + x, 7).Value = UPGRADE02
    Worksheets("Single").Cells(7 + x, 7).Value = REUPGRADE02
    Worksheets("Single").Cells(8 + x, 7).Value = DOWNGRADE02
    Worksheets("Single").Cells(9 + x, 7).Value = RENEW02
    Worksheets("Single").Cells(10 + x, 7).Value = ROLLBACK02

    Worksheets("Single").Cells(11 + x, 6).Value = "UPGRADE03"
    Worksheets("Single").Cells(12 + x, 6).Value = "REUPGRADE03"
    Worksheets("Single").Cells(13 + x, 6).Value = "DOWNGRADE03"
    Worksheets("Single").Cells(14 + x, 6).Value = "RENEW03"
    Worksheets("Single").Cells(15 + x, 6).Value = "ROLLBACK03"
    Worksheets("Single").Cells(11 + x, 7).Value = UPGRADE03
    Worksheets("Single").Cells(12 + x, 7).Value = REUPGRADE03
    Worksheets("Single").Cells(13 + x, 7).Value = DOWNGRADE03
    Worksheets("Single").Cells(14 + x, 7).Value = RENEW03
    Worksheets("Single").Cells(15 + x, 7).Value = ROLLBACK03
    
    'MsgBox "開始時間:" & mTime & "  結束時間：" & Time

End Sub

Sub MultipleCountLevel()
    
    'mTime = Time
    Worksheets("Multiple").Select
     
    Dim Multipletable()

    MultipleRows = Worksheets("Multiple").Range("A1").CurrentRegion.Rows.Count
    MultipleLines = Worksheets("Multiple").Range("A1").CurrentRegion.Columns.Count
    Multipletable = Worksheets("Multiple").Range(Cells(1, 1), Cells(MultipleRows, MultipleLines)).Value


    Dim UP00to01, UP00to02, UP00to03, UP01to02, UP01to03, UP02to03 As Long
    Dim DOWN01to00, DOWN02to00, DOWN02to01, DOWN03to00, DOWN03to02 As Long
    Dim RENEWto01,RENEWto02, RENEWto03, REUPGRADE01to02, REUPGRADE01to03, REUPGRADE02to03 As Long
    dim UNCHANGE00,UNCHANGE01,UNCHANGE02,UNCHANGE03,REUPGRADE02to02,REUPGRADE03to03 as long

    UP00to01 = 0
    UP00to02 = 0
    UP00to03 = 0
    UP01to02 = 0
    UP01to03 = 0
    UP02to03 = 0
    DOWN01to00 = 0
    DOWN02to00 = 0
    DOWN02to01 = 0
    DOWN03to00 = 0
    DOWN03to02 = 0
    RENEWto01 = 0
    RENEWto02 = 0
    RENEWto03 = 0
    REUPGRADE01to02 = 0
    REUPGRADE01to03 = 0
    REUPGRADE02to03 = 0
    UNCHANGE00 = 0
    UNCHANGE01 = 0
    UNCHANGE02 = 0
    UNCHANGE03 = 0
    REUPGRADE02to02 = 0
    REUPGRADE03to03 = 0

    ''
    For i = 2 To MultipleRows Step 1
        If Multipletable(i, 11) = "UPGRADE" Then
            If Multipletable(i, 3) = "00" And Multipletable(i, 8) = "01" Then
                UP00to01 = UP00to01 + 1
            ElseIf Multipletable(i, 3) = "00" And Multipletable(i, 8) = "02" Then
                UP00to02 = UP00to02 + 1
            ElseIf Multipletable(i, 3) = "00" And Multipletable(i, 8) = "03" Then
                UP00to03 = UP00to03 + 1
            ElseIf Multipletable(i, 3) = "01" And Multipletable(i, 8) = "02" Then
                UP01to02 = UP01to02 + 1
            ElseIf Multipletable(i, 3) = "01" And Multipletable(i, 8) = "03" Then
                UP01to03 = UP01to03 + 1
            ElseIf Multipletable(i, 3) = "02" And Multipletable(i, 8) = "03" Then
                UP02to03 = UP02to03 + 1
            End If
        ElseIf Multipletable(i, 11) = "DOWNGRADE" Then
            If Multipletable(i, 3) = "01" And Multipletable(i, 8) = "00" Then
                DOWN01to00 = DOWN01to00 + 1
            ElseIf Multipletable(i, 3) = "02" And Multipletable(i, 8) = "00" Then
                DOWN02to00 = DOWN02to00 + 1
            ElseIf Multipletable(i, 3) = "02" And Multipletable(i, 8) = "01" Then
                DOWN02to01 = DOWN02to01 + 1
            ElseIf Multipletable(i, 3) = "03" And Multipletable(i, 8) = "00" Then
                DOWN03to00 = DOWN03to00 + 1
            ElseIf Multipletable(i, 3) = "03" And Multipletable(i, 8) = "01" Then
                DOWN03to01 = DOWN03to01 + 1
            ElseIf Multipletable(i, 3) = "03" And Multipletable(i, 8) = "02" Then
                DOWN03to02 = DOWN03to02 + 1
            End If
        ElseIf Multipletable(i, 11) = "RENEW" Then
            If Multipletable(i, 3) = "02" Then
                RENEWto02 = RENEWto02 + 1
            ElseIf Multipletable(i, 3) = "03" Then
                RENEWto03 = RENEWto03 + 1
            elseif Multipletable(i,4) = "01" then
                RENEWto01 = RENEWto01 + 1
            End If
        ElseIf Multipletable(i, 11) = "REUPGRADE" Then
            If Multipletable(i, 3) = "01" And Multipletable(i, 8) = "02" Then
                REUPGRADE01to02 = REUPGRADE01to02 + 1
            ElseIf Multipletable(i, 3) = "01" And Multipletable(i, 8) = "03" Then
                REUPGRADE01to03 = REUPGRADE01to03 + 1
            ElseIf Multipletable(i, 3) = "02" And Multipletable(i, 8) = "03" Then
                REUPGRADE02to03 = REUPGRADE02to03 + 1
            ElseIf Multipletable(i, 3) = "02" And Multipletable(i, 8) = "02" Then
                REUPGRADE02to02 = REUPGRADE02to02 + 1
            ElseIf Multipletable(i, 3) = "03" And Multipletable(i, 8) = "03" Then
                REUPGRADE03to03 = REUPGRADE03to03 + 1
            End If
        ElseIf Multipletable(i, 11) = "UNCHANGE" Then
            If Multipletable(i, 3) = "01" And Multipletable(i, 8) = "01" Then
                UNCHANGE01 = UNCHANGE01 + 1
            ElseIf Multipletable(i, 3) = "00" And Multipletable(i, 8) = "00" Then
                UNCHANGE00 = UNCHANGE00 + 1
            ElseIf Multipletable(i, 3) = "02" And Multipletable(i, 8) = "02" Then
                UNCHANGE02 = UNCHANGE02 + 1
            ElseIf Multipletable(i, 3) = "03" And Multipletable(i, 8) = "03" Then
                UNCHANGE03 = UNCHANGE03 + 1
            End If
        End If
    Next i

    UPto03 = UPto03 + UP00to03 + UP01to03 + UP02to03
    UPto02 = UPto02 + UP00to02 + UP01to02 
    REto03 = REto03 + RENEWto03 + REUPGRADE01to03 + REUPGRADE02to03 + REUPGRADE03to03
    REto02 = REto02 + RENEWto02 + REUPGRADE01to02 + REUPGRADE02to02
    DOWNfrom03 = DOWNfrom03 + DOWN03to00 + DOWN03to01 + DOWN03to02
    DOWNfrom02 = DOWNfrom02 + DOWN02to00 + DOWN02to01 

    x = 0
    y = 0
    For i = 1 To 16 Step 1
        Worksheets("Multiple").Cells(i + 1, 13).NumberFormatLocal = "@"
        Worksheets("Multiple").Cells(i + 1, 13).Value = Format(x, "00")
        If i Mod 4 = 0 Then
            x = x + 1
        End If
        Worksheets("Multiple").Cells(i + 1, 14).NumberFormatLocal = "@"
        Worksheets("Multiple").Cells(i + 1, 14).Value = Format(y, "00")
        y = y + 1
        If y = 4 Then
            y = 0
        End If
    Next i

    Worksheets("Multiple").Cells(1, 13).Value = "OLD LEVEL"
    Worksheets("Multiple").Cells(1, 14).Value = "NEW LEVEL"
    Worksheets("Multiple").Cells(1, 15).Value = "UPGRADE"
    Worksheets("Multiple").Cells(1, 16).Value = "DOWNGRADE"
    Worksheets("Multiple").Cells(1, 17).Value = "RENEW"
    Worksheets("Multiple").Cells(1, 18).Value = "REUPGRADE"
    Worksheets("Multiple").Cells(1, 19).Value = "UNCHANGE"

    Worksheets("Multiple").Cells(3, 15).Value = UP00to01
    Worksheets("Multiple").Cells(4, 15).Value = UP00to02
    Worksheets("Multiple").Cells(5, 15).Value = UP00to03

    Worksheets("Multiple").Cells(8, 15).Value = UP01to02
    Worksheets("Multiple").Cells(9, 15).Value = UP01to03

    Worksheets("Multiple").Cells(13, 15).Value = UP02to03

    Worksheets("Multiple").Cells(6, 16).Value = DOWN01to00

    Worksheets("Multiple").Cells(10, 16).Value = DOWN02to00
    Worksheets("Multiple").Cells(11, 16).Value = DOWN02to01

    Worksheets("Multiple").Cells(14, 16).Value = DOWN03to00
    Worksheets("Multiple").Cells(15, 16).Value = DOWN03to01
    Worksheets("Multiple").Cells(16, 16).Value = DOWN03to02

    Worksheets("Multiple").Cells(7, 17).Value = RENEWto01
    Worksheets("Multiple").Cells(12, 17).Value = RENEWto02
    Worksheets("Multiple").Cells(17, 17).Value = RENEWto03

    Worksheets("Multiple").Cells(4, 18).Value = REUPGRADE00to02
    Worksheets("Multiple").Cells(5, 18).Value = REUPGRADE00to03

    Worksheets("Multiple").Cells(8, 18).Value = REUPGRADE01to02
    Worksheets("Multiple").Cells(9, 18).Value = REUPGRADE01to03

    Worksheets("Multiple").Cells(12, 18).Value = REUPGRADE02to02
    Worksheets("Multiple").Cells(13, 18).Value = REUPGRADE02to03

    Worksheets("Multiple").Cells(17, 18).Value = REUPGRADE03to03

    Worksheets("Multiple").Cells(2, 19).Value = UNCHANGE00
    Worksheets("Multiple").Cells(7, 19).Value = UNCHANGE01
    Worksheets("Multiple").Cells(12, 19).Value = UNCHANGE02
    Worksheets("Multiple").Cells(17, 19).Value = UNCHANGE03
    
    j = 22
    for i = 2 to Multiplerows step 1
        If Multipletable(i,11) = "" Then
            worksheets("START").cells(j,5).value = Multipletable(i,1)
            UncheckMember = UncheckMember + 1
            j = j + 1
        End If
    next i
    If j > 22 Then
        Worksheets("START").Cells(21, 5).Value = "無法判斷的會員"
    End If
end sub

sub clearSheets()

    Worksheets("Data").UsedRange.ClearContents
    Worksheets("Single").UsedRange.ClearContents
    Worksheets("Multiple").UsedRange.ClearContents
end sub