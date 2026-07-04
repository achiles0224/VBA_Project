Sub ASP_Bouns()
    WorkDay = Worksheets("計算結果").Range("I1").Value
    PointLow = Worksheets("計算結果").Range("X9").Value
    PointBouns = Worksheets("計算結果").Range("X10").Value

    '舊資料---'
    MacLowPCS = Worksheets("計算結果").Range("R9").Value
    MacBounsPSC = Worksheets("計算結果").Range("R10").Value
    iOSLowPCS = Worksheets("計算結果").Range("S9").Value
    iOSBounsPSC = Worksheets("計算結果").Range("S10").Value
    '--------'

    MacLowPrice = Worksheets("計算結果").Range("AG18").Value
    MacBounsPrice = Worksheets("計算結果").Range("AG19").Value
    iPhoneLowPrice = Worksheets("計算結果").Range("AE18").Value
    iPhoneBounsPrice = Worksheets("計算結果").Range("AE19").Value
    OtherLowPrice = Worksheets("計算結果").Range("AF18").Value
    OtherBounsPrice = Worksheets("計算結果").Range("AF19").Value

    ACCpoint_rate = Worksheets("計算結果").range("W13").value
    iPhonepoint_rate = Worksheets("計算結果").range("W14").value
    Macpoint_rate = Worksheets("計算結果").range("W15").value

    MsgBox ("最低點數 " & PointLow & ",加碼門檻 " & PointBouns) '& "Mac金額 " & MacLowPrice & ",加碼獎金" & MacBounsPrice & Chr(10) & "iOS最低件數金額  " & iPhoneLowPrice & "iPhone門檻獎金 " & iPhoneBounsPrice & ",ACC門檻獎金 " & OtherLowPrice & ",iPhone加碼獎金 " & iPhoneBounsPrice & ",ACC加碼獎金 " & OtherBounsPrice & Chr(10) & "請按確認繼續")

    AllRowsCounts = Worksheets("完修總表").Range("A1").CurrentRegion.Rows.Count
    AllColCounts = Worksheets("完修總表").Range("A1").CurrentRegion.Columns.Count
    FinalRowsCounts = Worksheets("計算結果").Range("A4").CurrentRegion.Rows.Count
    FinalColCounts = Worksheets("計算結果").Range("A4").CurrentRegion.Columns.Count
    'FinalColCounts1 = Worksheets("計算結果").Range("A4").value
    'msgbox(FinalRowsCounts)
    For i = 4 To FinalRowsCounts Step 1
        PointCount = 0
        OtherlowQty = 0
        iPhonelowQty = 0
        OtherBounsQty = 0
        iPhoneBounsQty = 0
        MaclowQty = 0
        MacBounsQty = 0

        Application.StatusBar = "處理中 " & (i / FinalRowsCounts) & "%"

        'Worksheets("計算結果").Cells(i, 14).Value = "=IF(F" & i & ">" & MacLowPCS & ",IF(F" & i & "<=" & MacBounsPSC & ",F" & i & "*" & MacLowPrice & "," & MacBounsPSC * MacLowPrice & "+(F" & i & "-" & MacBounsPSC & ")*" & MacBounsPrice & "),0)"
        
        iOScount = Worksheets("計算結果").Cells(i, 4).Value
        ACCCount = Worksheets("計算結果").Cells(i, 3).Value
        iPhoneCount = Worksheets("計算結果").Cells(i, 2).Value
        MacCount = Worksheets("計算結果").Cells(i, 6).Value

        empName = Worksheets("計算結果").Cells(i, 1).Value
        TotalPoint = (ACCCount * ACCpoint_rate) + (iPhoneCount * iPhonepoint_rate) + (MacCount * Macpoint_rate)
        Worksheets("計算結果").Cells(i, 10).Value = TotalPoint

        
        If TotalPoint < PointLow Then
            Worksheets("計算結果").Cells(i, 20).Value = 0

        Else
            For j = 2 To AllRowsCounts Step 1
                
                If Worksheets("完修總表").Cells(j, 23).Value = empName Then
                    If Worksheets("完修總表").Cells(j, 13).Value = "iOS" Then
                        If Worksheets("完修總表").Cells(j, 14).Value = "ACC" Then
                         '計算ACC點數換算獎金，一件0.5點'
                        PointCount = PointCount + ACCpoint_rate
                            '判斷數量是否達標
                            if PointCount <= PointBouns Then
                                OtherlowQty = OtherlowQty + 1
                            Else
                                OtherBounsQty = OtherBounsQty + 1
                            end if
                        Else
                        '計算iPhone點數換算獎金，一件1點'
                        PointCount = PointCount + iPhonepoint_rate
                            if PointCount <= PointBouns Then
                                iPhonelowQty = iPhonelowQty + 1
                            Else
                                iPhoneBounsQty = iPhoneBounsQty + 1
                            end if
                        end if
                    Else
                        '計算MAC點數換算獎金，一件2點'
                        PointCount = PointCount + Macpoint_rate
                    'msgbox(AlliOSCount)
                        If PointCount <= PointBouns Then
                            MaclowQty = MaclowQty + 1
                        Else
                            MacBounsQty = MacBounsQty + 1
                        End If
                    End If
                End If
            Next j
        End If
        Worksheets("計算結果").Cells(i, 12).Value = iPhonelowQty
        Worksheets("計算結果").Cells(i, 11).Value = OtherlowQty
        Worksheets("計算結果").Cells(i, 15).Value = iPhoneBounsQty
        Worksheets("計算結果").Cells(i, 14).Value = OtherBounsQty
        Worksheets("計算結果").Cells(i, 13).Value = MaclowQty
        Worksheets("計算結果").Cells(i, 16).Value = MacBounsQty
        Worksheets("計算結果").Cells(i, 17).Value = OtherlowQty * OtherLowPrice + OtherBounsQty * OtherBounsPrice
        Worksheets("計算結果").Cells(i, 18).Value = iPhonelowQty * iPhoneLowPrice + iPhoneBounsQty * iPhoneBounsPrice
        Worksheets("計算結果").Cells(i, 19).Value = MaclowQty * MacLowPrice + MacBounsQty * MacBounsPrice
        Worksheets("計算結果").Cells(i, 20).Value = Worksheets("計算結果").Cells(i, 17).Value + Worksheets("計算結果").Cells(i, 18).Value + Worksheets("計算結果").Cells(i, 19).Value
        Worksheets("計算結果").Cells(i, 8).Value = "=VLOOKUP(A" & i & ",完修總表!W:Z,2,FALSE)"
        Worksheets("計算結果").Cells(i, 9).Value = "=VLOOKUP(A" & i & ",完修總表!W:Z,3,FALSE)"
    Next i
    Application.StatusBar = False

    MsgBox ("完成")
End Sub

Sub ClearData()
    Worksheets("計算結果").Range(Cells(4, 8), Cells(9999, 20)).Clear
End Sub



