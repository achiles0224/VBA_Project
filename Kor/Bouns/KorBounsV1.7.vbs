Sub CHECK_TYPE()
    targe1 = Worksheets("target").Range("H2").Value
    'MsgBox (targe1)
    If (Range("A11").Value = 0) And (Range("C11").Value <> 0) Then
        Worksheets("target").Range("H2:K99").Clear
        StoreBonus
        
        copyAttch
        Worksheets("final").Select
        MsgBox ("Shop incentive is Done")
    ElseIf (Range("A11").Value <> 0 And Range("C11").Value <> 0) Then
        If targe1 = "" Then
            MsgBox ("Please execute the Shop program first")
        Else
            EmpoloyeeBonus
            EmpBounsfor3pp
            MsgBox ("Staff incentive is Done")
        End If
    Else
        MsgBox ("Please Check Data")
    End If
End Sub


Sub StoreBonus()
    Application.ScreenUpdating = False

    Worksheets("target").Range("H2:K99").Clear
    Dim GPRowCount As Long
    Dim StoreMargin As Double
    Dim srcRange, fndRange As Range
    GPRowCount = Worksheets("A_GP").Range("C10").CurrentRegion.Rows.Count
    GPRowCount = GPRowCount - 1
    Worksheets("A_GP").Activate
    Set srcRange = Worksheets("Target").Range("A1").CurrentRegion.Columns(1)

    For i = 11 To GPRowCount + 10 Step 1
        Set fndRange = srcRange.Find(what:=Cells(i, 3).Value)
        
        If Not fndRange Is Nothing Then
           Worksheets("A_GP").Cells(i, 7).Value = fndRange.Offset(, 2).Value
           If Worksheets("A_GP").Cells(i, 6).Value = 0 Then
                MsgBox ("Row no." & i & Cells(i, 3).Value & " target margin is zero, pls check data.")
            Else
                StoreMargin = Cells(i, 6).Value / Cells(i, 7).Value
           End If
            
            Worksheets("A_GP").Cells(i, 8).Value = StoreMargin
            If StoreMargin >= 0.8 Then
                Worksheets("A_GP").Cells(i, 9).Value = Cells(i, 6).Value * 0.01 * StoreMargin
            Else
                Worksheets("A_GP").Cells(i, 9).Value = 0
            End If
        Else
            MsgBox ("Row " & i & " Data is null")
        End If
    Next i

    Range(Cells(11, 1), Cells(i - 1, 9)).Copy
    Worksheets("Final").Select
    Worksheets("Final").Range(Cells(2, 1), Cells(i - 10, 9)).PasteSpecial xlPasteValues
    
    Worksheets("target").Select
    targetRowCount = Worksheets("target").Range("A1").CurrentRegion.Rows.Count
    Set srcRange = Worksheets("Final").Range("A1").CurrentRegion.Columns(3)
       For t = 2 To targetRowCount Step 1
        Set fndRange = srcRange.Find(what:=Cells(t, 1).Value)
        If Not fndRange Is Nothing Then
            Worksheets("target").Cells(t, 8).Value = fndRange.Offset(, 5).Value
            'MSGBOX("!!!!")
        End If
       Next t
    
    'MsgBox ("Done")
End Sub

Sub EmpoloyeeBonus()
    Application.ScreenUpdating = False

    Dim GPRowCount As Long
    Dim EmpMargin As Double
    Dim srcRange, fndRange As Range
    GPRowCount = Worksheets("A_GP").Range("A10").CurrentRegion.Rows.Count
    GPRowCount = GPRowCount - 1
    Worksheets("A_GP").Activate
    Set srcRange = Worksheets("Target").Range("A1").CurrentRegion.Columns(1)

    For i = 11 To GPRowCount + 10 Step 1
        Set fndRange = srcRange.Find(what:=Cells(i, 3).Value)
        
        If Not fndRange Is Nothing Then
           Worksheets("A_GP").Cells(i, 7).Value = fndRange.Offset(, 2).Value
            If Worksheets("A_GP").Cells(i, 6).Value = 0 Then
                MsgBox ("Row" & i & "margin is zero, pls check data.")
                EmpMargin = 0
            Else
                EmpMargin = fndRange.Offset(, 7).Value
            End If
            Worksheets("A_GP").Cells(i, 8).Value = EmpMargin

            If EmpMargin >= 0.8 Then
                Worksheets("A_GP").Cells(i, 9).Value = Cells(i, 6).Value * 0.03 * EmpMargin
            Else
                Worksheets("A_GP").Cells(i, 9).Value = 0
            End If
        Else
            MsgBox ("Row " & i & " Data is null")
        End If
    Next i

    Range(Cells(11, 1), Cells(i - 1, 9)).Copy
    Worksheets("Final").Select
    Worksheets("Final").Range(Cells(2, 1), Cells(i - 10, 9)).PasteSpecial xlPasteValues
    
    'MsgBox ("Done")
End Sub

Sub Bounsfor3pp()
    Application.ScreenUpdating = False

    Dim pppRowCount As Long
    Dim Store3pp, AchRate3pp As Double
    Dim srcRange, fndRange As Range
    AchRate3pp = 1
    pppRowCount = Worksheets("B_3pp").Range("C10").CurrentRegion.Rows.Count
    pppRowCount = pppRowCount - 1
    Worksheets("B_3pp").Activate
    Set srcRange = Worksheets("Target").Range("A1").CurrentRegion.Columns(1)
    
    For i = 11 To pppRowCount + 10 Step 1
        Set fndRange = srcRange.Find(what:=Cells(i, 3).Value)
        If Not fndRange Is Nothing Then
           Worksheets("B_3pp").Cells(i, 6).Value = fndRange.Offset(, 3).Value
        Else
            MsgBox ("Row " & i & " Data is null")
        End If
    Next i

    'Range(Cells(11, 1), Cells(i - 1, 6)).Copy
    'Worksheets("Final").Select
    'Worksheets("Final").Range(Cells(2, 1), Cells(i - 10, 6)).PasteSpecial xlPasteValues
    Worksheets("Final").Select
    Worksheets("Final").Activate
    finRowCount = Worksheets("Final").Range("C10").CurrentRegion.Rows.Count
    Dim tempno As Double
    
    Set srcRange = Worksheets("B_3pp").Range("A10").CurrentRegion.Columns(3)
        For j = 2 To finRowCount Step 1
            Set fndRange = srcRange.Find(what:=Cells(j, 3).Value)
            If Not fndRange Is Nothing Then
            
            Worksheets("Final").Cells(j, 10).Value = fndRange.Offset(, 2).Value
            Worksheets("Final").Cells(j, 11).Value = fndRange.Offset(, 3).Value
            'Store3pp = Cells(i, 10).Value / Cells(i, 5).Value
            'MsgBox ("i10 is " & Worksheets("Final").Cells(j, 10).Value)
            'MsgBox ("i5 is " & Worksheets("Final").Cells(j, 5).Value)
           Worksheets("Final").Cells(j, 12).Value = CLng(Worksheets("Final").Cells(j, 10).Value) / CLng(Worksheets("Final").Cells(j, 5).Value)
        
            'MsgBox (tempno)
            If Worksheets("Final").Cells(j, 11).Value <> 0 Then
            Store3pp = Worksheets("Final").Cells(j, 12).Value / Worksheets("Final").Cells(j, 11).Value
            End If
            'If Store3pp >= 0.4 Then
            ''    Worksheets("B_3pp").Cells(i, 6).Value = Cells(i, 3).Value * 0.01 * Store3pp
            'Else
            ''    Worksheets("B_3pp").Cells(i, 6).Value = 0
            'End If
            
            Select Case Store3pp
             Case Is >= 1.1
                AchRate3pp = 1.3
             Case 1 To 1.09
                AchRate3pp = 1.2
             Case 0.9 To 0.99
                AchRate3pp = 1.1
             Case 0.7 To 0.79
                AchRate3pp = 0.9
             Case 0.6 To 0.69
                AchRate3pp = 0.8
             Case 0.5 To 0.59
                AchRate3pp = 0.7
             Case Is < 0.5
                AchRate3pp = 0.5
            End Select
            Worksheets("Final").Cells(j, 13).Value = Store3pp
            Worksheets("Final").Cells(j, 14).Value = AchRate3pp

            'If Not fndRange Is Nothing Then
                
                
                'Worksheets("Final").Cells(j, 11).Value = fndRange.Offset(, 4).Value
                'Worksheets("Final").Cells(j, 12).Value = fndRange.Offset(, 5).Value
                Worksheets("Final").Cells(j, 15).Value = Worksheets("Final").Cells(j, 9).Value * Worksheets("Final").Cells(j, 14).Value
            End If
        Next j

        Worksheets("target").Select
        targetRowCount = Worksheets("target").Range("A1").CurrentRegion.Rows.Count
        Set srcRange = Worksheets("Final").Range("A1").CurrentRegion.Columns(3)
            For t = 2 To targetRowCount Step 1
                Set fndRange = srcRange.Find(what:=Cells(t, 1).Value)
                If Not fndRange Is Nothing Then
                    Worksheets("target").Cells(t, 9).Value = fndRange.Offset(, 10).Value
                    'MSGBOX("!!!!")
                End If
           Next t
        
    'MsgBox ("Done")
End Sub

Sub Staff3pp()
    Application.ScreenUpdating = False
    finRowCount = Worksheets("Final").Range("C1").CurrentRegion.Rows.Count
    Set srcRange = Worksheets("Target").Range("A1").CurrentRegion.Columns(1)

    For i = 2 To finRowCount Step 1
        Set fndRange = srcRange.Find(what:=Cells(i, 3).Value)
        If Not fndRange Is Nothing Then
            Worksheets("Final").cells(i.12).value = fndRange.Offset(,8).value

    Next i
End Sub

Sub EmpBounsfor3pp()
    Application.ScreenUpdating = False

    Dim pppRowCount As Long
    Dim emp3pp, AchRate3ppForEmp As Double
    Dim srcRange, fndRange As Range
    
    'Range(Cells(11, 1), Cells(i - 1, 6)).Copy
    'Worksheets("Final").Select
    'Worksheets("Final").Range(Cells(2, 1), Cells(i - 10, 6)).PasteSpecial xlPasteValues

    
    Worksheets("Final").Select
    Worksheets("Final").Activate
    finRowCount = Worksheets("Final").Range("A1").CurrentRegion.Rows.Count
    
    Set srcRange = Worksheets("Target").Range("A1").CurrentRegion.Columns(1)
        For j = 2 To finRowCount Step 1
            Set fndRange = srcRange.Find(what:=Cells(j, 3).Value)
            'If Worksheets("final").Cells(j, 5) <> 0 Then
            If Not fndRange Is Nothing Then
                'Worksheets("Final").Cells(j, 10).Value = fndRange.Offset(, 4).Value
                'Worksheets("Final").Cells(j, 11).Value = fndRange.Offset(, 5).Value
                Worksheets("Final").Cells(j, 13).Value = fndRange.Offset(, 8).Value
                emp3pp = Worksheets("Final").Cells(j, 13).Value = emp3pp
            'MsgBox ("Emp3pp " & emp3pp)
        
                Select Case Worksheets("Final").Cells(j, 13).Value
                 Case Is >= 1.1
                    AchRate3ppForEmp = 1.3
                 Case 1 To 1.09
                    AchRate3ppForEmp = 1.2
                 Case 0.9 To 0.99
                    AchRate3ppForEmp = 1.1
                 Case 0.7 To 0.79
                    AchRate3ppForEmp = 0.9
                 Case 0.6 To 0.69
                    AchRate3ppForEmp = 0.8
                 Case 0.5 To 0.59
                    AchRate3ppForEmp = 0.7
                 Case Is < 0.5
                    AchRate3ppForEmp = 0.5
                End Select
            'MsgBox (AchRate3ppForEmp & " %")
            
                Worksheets("Final").Cells(j, 14).Value = AchRate3ppForEmp

                Worksheets("Final").Cells(j, 15).Value = Worksheets("Final").Cells(j, 9).Value * Worksheets("Final").Cells(j, 14).Value
            'End If
            End If
        Next j
        
    'MsgBox ("Done")
End Sub

Sub CPU_StockBounsByStore()
    Application.ScreenUpdating = False

    Dim cpuRowCount As Long
    Dim StoreCpu, StoreStock, AchRateCpu, AchRateStock As Double
    Dim srcRange, fndRange As Range
    AchRateCpu = 1
    AchRateStock = 1

    cpuRowCount = Worksheets("C_Stock&CPU").Range("C10").CurrentRegion.Rows.Count
    cpuRowCount = cpuRowCount - 1
    Worksheets("C_Stock&CPU").Activate
    Set srcRange = Worksheets("Target").Range("A1").CurrentRegion.Columns(1)
    
    For i = 11 To cpuRowCount + 10 Step 1
        Set fndRange = srcRange.Find(what:=Cells(i, 3).Value)
        If Not fndRange Is Nothing Then
           Worksheets("C_Stock&CPU").Cells(i, 7).Value = fndRange.Offset(, 6).Value
           Worksheets("C_Stock&CPU").Cells(i, 10).Value = fndRange.Offset(, 4).Value
           Worksheets("C_Stock&CPU").Cells(i, 11).Value = fndRange.Offset(, 5).Value
            StoreCpu = Cells(i, 6).Value / Cells(i, 7).Value
            Worksheets("C_Stock&CPU").Cells(i, 8).Value = StoreCpu
            StoreStock = Cells(i, 11).Value / Cells(i, 10).Value
            Worksheets("C_Stock&CPU").Cells(i, 12).Value = StoreStock


            'If Store3pp >= 0.4 Then
            ''    Worksheets("B_3pp").Cells(i, 6).Value = Cells(i, 3).Value * 0.01 * Store3pp
            'Else
            ''    Worksheets("B_3pp").Cells(i, 6).Value = 0
            'End If

            Select Case StoreCpu
             Case Is >= 1.2
                AchRateCpu = 1.3
                CPUBousPcs = 13000
             Case 1.1 To 1.19
                AchRateCpu = 1.2
                CPUBousPcs = 13000
             Case 1 To 1.09
                AchRateCpu = 1.1
                CPUBousPcs = 13000
             Case 0.9 To 0.99
                AchRateCpu = 1
                CPUBousPcs = 10000
             Case 0.8 To 0.89
                AchRateCpu = 0.9
                CPUBousPcs = 10000
             Case 0.7 To 0.79
                AchRateCpu = 0.8
                CPUBousPcs = 5000
             Case 0.6 To 0.69
                AchRateCpu = 0.7
                CPUBousPcs = 5000
             Case Is < 0.6
                AchRateCpu = 0.5
                CPUBousPcs = 0
            End Select



            Select Case StoreStock
             Case Is >= 1.1
                AchRateStock = 1.3
             Case 1 To 1.09
                AchRateStock = 1.2
             Case 0.9 To 0.99
                AchRateStock = 1.1
             Case 0.8 To 0.89
                AchRateStock = 1
             Case 0.7 To 0.79
                AchRateStock = 0.9
             Case 0.6 To 0.69
                AchRateStock = 0.8
             Case 0.5 To 0.59
                AchRateStock = 0.7
             Case Is < 0.5
                AchRateStock = 0.5
            End Select


            Worksheets("C_Stock&CPU").Cells(i, 9).Value = AchRateCpu
            Worksheets("C_Stock&CPU").Cells(i, 13).Value = AchRateStock
            'Worksheets("C_Stock&CPU").Cells(i, 14).Value = CPUBousPcs
        Else
            MsgBox ("Row " & i & " Data is null")
        End If
    Next i

    Worksheets("target").Select
    targetRowCount = Worksheets("target").Range("A1").CurrentRegion.Rows.Count
    Set srcRange = Worksheets("Final").Range("A1").CurrentRegion.Columns(3)
       For t = 2 To targetRowCount Step 1
        Set fndRange = srcRange.Find(what:=Cells(t, 1).Value)
        If Not fndRange Is Nothing Then
            Worksheets("target").Cells(t, 10).Value = fndRange.Offset(, 18).Value
            Worksheets("target").Cells(t, 11).Value = fndRange.Offset(, 19).Value
            'MSGBOX("!!!!")
        End If
       Next t

    'Range(Cells(11, 1), Cells(i - 1, 6)).Copy
    'Worksheets("Final").Select
    'Worksheets("Final").Range(Cells(2, 1), Cells(i - 10, 6)).PasteSpecial xlPasteValues
    Worksheets("Final").Select
    Worksheets("Final").Activate
    finRowCount = Worksheets("Final").Range("C10").CurrentRegion.Rows.Count
    
    Set srcRange = Worksheets("C_Stock&CPU").Range("A10").CurrentRegion.Columns(3)
        For j = 2 To finRowCount Step 1
            Set fndRange = srcRange.Find(what:=Cells(j, 3).Value)

            If Not fndRange Is Nothing Then
            'for cpu stock target'
                Worksheets("Final").Cells(j, 16).Value = fndRange.Offset(, 9).Value
                Worksheets("Final").Cells(j, 17).Value = fndRange.Offset(, 10).Value
                Worksheets("Final").Cells(j, 18).Value = Worksheets("Final").Cells(j, 17).Value * Worksheets("Final").Cells(j, 15).Value
             'for cpu sale target'
                Worksheets("Final").Cells(j, 19).Value = fndRange.Offset(, 3).Value
                Worksheets("Final").Cells(j, 20).Value = fndRange.Offset(, 4).Value
                Worksheets("Final").Cells(j, 21).Value = fndRange.Offset(, 5).Value
                Worksheets("Final").Cells(j, 22).Value = fndRange.Offset(, 6).Value
                Worksheets("Final").Cells(j, 23).Value = Worksheets("Final").Cells(j, 18).Value * Worksheets("Final").Cells(j, 22).Value
             'for cpu pes'
                'Worksheets("Final").Cells(j, 22).Value = fndRange.Offset(, 3).Value
                
                CPUget = Worksheets("Final").Cells(j, 21).Value

            Select Case CPUget
             Case Is >= 1.2
                CPUB_pcs = 13000
             Case 1.1 To 1.19
                CPUBousPcs = 13000
             Case 1 To 1.09
                CPUB_pcs = 13000
             Case 0.9 To 0.99
                CPUB_pcs = 10000
             Case 0.8 To 0.89
                CPUB_pcs = 10000
             Case 0.7 To 0.79
                CPUB_pcs = 5000
             Case 0.6 To 0.69
                CPUB_pcs = 5000
             Case Is < 0.6
                CPUB_pcs = 0
            End Select

                Worksheets("Final").Cells(j, 24).Value = Worksheets("Final").Cells(j, 19).Value * CPUB_pcs
                Worksheets("Final").Cells(j, 25).Value = Worksheets("Final").Cells(j, 23).Value + Worksheets("Final").Cells(j, 24).Value
            End If
        Next j

    
    'MsgBox ("Done")
End Sub

Sub clearStoreTempForA_GP()
'Worksheets("target").Range("H2:H99").Clear
Worksheets("A_GP").Select
Range("A11:ZZ65536").EntireRow.Delete
ActiveSheet.UsedRange
Worksheets("Final").Select
Range("A2:ZZ65536").EntireRow.Delete
ActiveSheet.UsedRange
Worksheets("B_3PP").Select
Range("F11:ZZ65536").EntireRow.Delete
ActiveSheet.UsedRange
Worksheets("C_Stock&CPU").Select
Range("G11:ZZ65536").EntireRow.Delete
ActiveSheet.UsedRange
Worksheets("A_GP").Select
MsgBox ("Done")

End Sub

Sub clearB()
    Worksheets("B_3PP").Range(Cells(11, 1), Cells(9999, 20)).Clear
    Worksheets("Final").Select
    Worksheets("Final").Range(Cells(2, 9), Cells(9999, 13)).Clear
    Worksheets("B_3PP").Select
End Sub

Sub clearC()
    Worksheets("C_Stock&CPU").Range(Cells(11, 1), Cells(9999, 15)).Clear
    Worksheets("Final").Select
    Worksheets("Final").Range(Cells(2, 14), Cells(9999, 50)).Clear
    Worksheets("C_Stock&CPU").Select
End Sub

Sub CPU_StockBounsByEmp()
    Application.ScreenUpdating = False

    Dim cpuRowCount As Long
    Dim EmpCpu, EmpStock, AchRateCpu, AchRateStock As Double
    Dim srcRange, fndRange As Range
    AchRateCpu = 1
    AchRateStock = 1

    Worksheets("Final").Select
    finRowCount = Worksheets("Final").Range("A1").CurrentRegion.Rows.Count
    
    Set srcRange = Worksheets("Target").Range("A1").CurrentRegion.Columns(1)
        For j = 2 To finRowCount Step 1
            Set fndRange = srcRange.Find(what:=Cells(j, 3).Value)
            'If Worksheets("final").Cells(j, 5) <> 0 Then
            If Not fndRange Is Nothing Then
                'Worksheets("Final").Cells(j, 10).Value = fndRange.Offset(, 4).Value
                'Worksheets("Final").Cells(j, 11).Value = fndRange.Offset(, 5).Value
                Worksheets("Final").Cells(j, 16).Value = fndRange.Offset(, 9).Value
                Worksheets("Final").Cells(j, 21).Value = fndRange.Offset(, 10).Value
                EmpStock = Worksheets("Final").Cells(j, 16).Value
                EmpCpu = Worksheets("Final").Cells(j, 21).Value

                Select Case Worksheets("Final").Cells(j, 21).Value
                 Case Is >= 1.2
                    AchRateCpu = 1.3
                 Case 1.1 To 1.19
                    AchRateCpu = 1.2
                 Case 1 To 1.09
                    AchRateCpu = 1.1
                 Case 0.9 To 0.99
                    AchRateCpu = 1
                 Case 0.8 To 0.89
                    AchRateCpu = 0.9
                 Case 0.7 To 0.79
                    AchRateCpu = 0.8
                 Case 0.6 To 0.69
                    AchRateCpu = 0.7
                 Case Is < 0.6
                    AchRateCpu = 0.5
                End Select


                Select Case Worksheets("Final").Cells(j, 16).Value
                     Case Is >= 1.1
                        AchRateStock = 1.3
                     Case 1 To 1.09
                        AchRateStock = 1.2
                     Case 0.9 To 0.99
                        AchRateStock = 1.1
                     Case 0.8 To 0.89
                        AchRateStock = 1
                     Case 0.7 To 0.79
                        AchRateStock = 0.9
                     Case 0.6 To 0.69
                        AchRateStock = 0.8
                     Case 0.5 To 0.59
                        AchRateStock = 0.7
                     Case Is < 0.5
                        AchRateStock = 0.5
                    End Select
                    Worksheets("Final").Cells(j, 17).Value = AchRateStock
                    Worksheets("Final").Cells(j, 22).Value = AchRateCpu
            Else
            MsgBox ("Row " & j & " Data is null")
        End If
    Next j

    Worksheets("Final").Select
    Worksheets("Final").Activate
    finRowCount = Worksheets("Final").Range("A1").CurrentRegion.Rows.Count
    
    Set srcRange = Worksheets("C_Stock&CPU").Range("A10").CurrentRegion.Columns(3)
        For j = 2 To finRowCount Step 1
            Set fndRange = srcRange.Find(what:=Cells(j, 3).Value)

            If Not fndRange Is Nothing Then
            'for cpu stock target'
                'Worksheets("Final").Cells(j, 16).Value = fndRange.Offset(, 9).Value
                'Worksheets("Final").Cells(j, 17).Value = fndRange.Offset(, 10).Value
                Worksheets("Final").Cells(j, 18).Value = Worksheets("Final").Cells(j, 17).Value * Worksheets("Final").Cells(j, 15).Value
             'for cpu sale target'
                Worksheets("Final").Cells(j, 19).Value = fndRange.Offset(, 3).Value
                'Worksheets("Final").Cells(j, 20).Value = fndRange.Offset(, 4).Value
                'Worksheets("Final").Cells(j, 21).Value = fndRange.Offset(, 5).Value
                'Worksheets("Final").Cells(j, 22).Value = fndRange.Offset(, 6).Value
                Worksheets("Final").Cells(j, 23).Value = Worksheets("Final").Cells(j, 18).Value * Worksheets("Final").Cells(j, 22).Value
             'for cpu pes'
                'Worksheets("Final").Cells(j, 22).Value = fndRange.Offset(, 3).Valuea

                CPUget = Worksheets("Final").Cells(j, 21).Value

            Select Case CPUget
             Case Is >= 1.2
                CPUB_pcs = 13000
             Case 1.1 To 1.19
                CPUBousPcs = 13000
             Case 1 To 1.09
                CPUB_pcs = 13000
             Case 0.9 To 0.99
                CPUB_pcs = 10000
             Case 0.8 To 0.89
                CPUB_pcs = 10000
             Case 0.7 To 0.79
                CPUB_pcs = 5000
             Case 0.6 To 0.69
                CPUB_pcs = 5000
             Case Is < 0.6
                CPUB_pcs = 0
            End Select

                Worksheets("Final").Cells(j, 24).Value = Worksheets("Final").Cells(j, 19).Value * CPUB_pcs
                'Worksheets("Final").Cells(j, 24).Value = Worksheets("Final").Cells(j, 19).Value * 10000
                Worksheets("Final").Cells(j, 25).Value = Worksheets("Final").Cells(j, 23).Value + Worksheets("Final").Cells(j, 24).Value
            End If
        Next j
        
    'MsgBox ("Done")
End Sub


Sub CHECK_TYPE_3pp()

    If (Range("A11").Value = 0) And (Range("C11").Value <> 0) Then
        Bounsfor3pp
        MsgBox ("Shop 3PP attach is Done")
    ElseIf (Range("A11").Value <> 0 And Range("C11").Value <> 0) Then
            EmpBounsfor3pp
            MsgBox ("Staff 3PP attach is Done")
    Else
        MsgBox ("Please Check Data")
    End If
End Sub

Sub CHECK_TYPE_CPUStock()

    If (Range("A11").Value = 0) And (Range("C11").Value <> 0) Then
        CPU_StockBounsByStore
        copyAttch
        refreshPivot
        MsgBox ("Shop CPU_Stock is Done")
    ElseIf (Range("A11").Value <> 0 And Range("C11").Value <> 0) Then
            CPU_StockBounsByEmp
            refreshPivot
            MsgBox ("Staff CPU_Stock is Done")
    Else
        MsgBox ("Please Check Data")
    End If
End Sub


Sub refreshPivot()
    ThisWorkbook.RefreshAll
End Sub



Sub copyAttch()
    Worksheets("Target").Select
    Worksheets("Target").Activate
    finRowCount = Worksheets("Target").Range("A1").CurrentRegion.Rows.Count
    
    Set srcRange = Worksheets("Final").Range("A1").CurrentRegion.Columns(3)
        For j = 2 To finRowCount Step 1
            Set fndRange = srcRange.Find(what:=Cells(j, 1).Value)

            If Not fndRange Is Nothing Then
            
                Worksheets("Target").Cells(j, 10).Value = fndRange.Offset(, 13).Value
                Worksheets("Target").Cells(j, 11).Value = fndRange.Offset(, 18).Value
               
            End If
        Next j
End Sub








