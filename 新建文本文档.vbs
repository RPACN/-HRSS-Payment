Sub bank_file()
Dim rngCom As Range, rngICBC As Range, wbSplit As Workbook, wsSplit As Worksheet
Dim iRowCov As Integer, iColCov As Integer, iRow As Integer, iNew As Integer, iCnt As Integer, iMult As Integer, iMod As Double, iLoop As Integer

With Sheet2
    For iLoop = .Range("A65536").End(xlUp).Row To 2 Step -1
        '检查银行卡信息
        Set rng = Sheet1.Range("A:A").Find(.Range("A" & iLoop), , , xlWhole)
        If Not rng Is Nothing Then
            If rng.Offset(0, 2) <> .Range("K" & iLoop) Then
                MsgBox "'" & .Range("I" & iLoop) & "' 的银行卡错误，请更新该员工的银行卡信息！", vbError, "银行卡错误："
                Exit Sub
            End If
        End If
        Set rng = Nothing
    Next iLoop
End With

StrPt = ThisWorkbook.Path & "\网银报盘" & Format(Date, "YYYYMMDD") & "\"
If Dir(StrPt, vbDirectory) = "" Then
    MkDir StrPt
End If

With Sheet2
    .Range("A1:M" & .Range("A65536").End(xlUp).Row).Sort .Range("A2:A" & .Range("A65536").End(xlUp).Row) _
        , xlAscending, , , , , , xlYes
    Sheet1.Range("J:J").ClearContents
    .Range("A2:A" & .Range("A65536").End(xlUp).Row).Copy Sheet1.Range("J1")
    Sheet1.Range("J:J").RemoveDuplicates 1, xlNo
    For iLoop = 1 To Sheet1.Range("J65536").End(xlUp).Row
        Set rng = Sheet1.Range("A:A").Find(Sheet1.Range("J" & iLoop), , , xlWhole)
        If Not rng Is Nothing Then
            Set wb = Workbooks.Add(1)
            Set ws = wb.ActiveSheet
            iCnt = WorksheetFunction.CountIf(.Range("A:A"), rng.Value)
            Set rngCom = .Range("A:A").Find(rng.Value, , , xlWhole)
            Select Case rng.Offset(0, 3)
                Case "1" '招商银行1 北京公司
                    Sheet1.Range("U2:X2").Copy ws.Range("A1")
                    rngCom.Offset(0, 8).Resize(iCnt, 1).Copy
                    ws.Range("A2").PasteSpecial xlPasteValuesAndNumberFormats
                    rngCom.Offset(0, 11).Resize(iCnt, 1).Copy
                    ws.Range("B2").PasteSpecial xlPasteValuesAndNumberFormats
                    rngCom.Offset(0, 12).Resize(iCnt, 1).Copy
                    ws.Range("C2").PasteSpecial xlPasteValuesAndNumberFormats
                    ws.Range("D2:D" & ws.Range("A65536").End(xlUp).Row) = rng.Offset(0, 1) & rngCom.Offset(0, 9) & "工资"
                    iColCov = 3
                Case "2" '招商银行2 '广东公司
                    Sheet1.Range("U3:AA3").Copy ws.Range("A1")
                    rngCom.Offset(0, 11).Resize(iCnt, 1).Copy
                    ws.Range("A2").PasteSpecial xlPasteValuesAndNumberFormats
                    rngCom.Offset(0, 8).Resize(iCnt, 1).Copy
                    ws.Range("B2").PasteSpecial xlPasteValuesAndNumberFormats
                    rngCom.Offset(0, 12).Resize(iCnt, 1).Copy
                    ws.Range("C2").PasteSpecial xlPasteValuesAndNumberFormats
                    ws.Range("D2") = "U0001"
                    If ws.Range("A65536").End(xlUp).Row > 2 Then
                        ws.Range("D2").AutoFill ws.Range("D2:D" & ws.Range("A65536").End(xlUp).Row), xlFillSeries
                    End If
                    ws.Range("E2:E" & ws.Range("A65536").End(xlUp).Row) = rng.Offset(0, 1) & rngCom.Offset(0, 9) & "工资"
                    iColCov = 3
                Case "3" '中国农业银行
                    Sheet1.Range("U4:Y4").Copy ws.Range("A1")
'                    ws.Range("A1:E1").Merge
'                    ws.Range("A2") = "企业："
'                    ws.Range("B2") = rngCom.Value
'                    ws.Range("E2") = "单位：元"
'                    Sheet1.Range("U3:Y3").Copy ws.Range("A3")
'                    ws.UsedRange.EntireColumn.AutoFit
                    rngCom.Offset(0, 8).Resize(iCnt, 1).Copy
                    ws.Range("B2").PasteSpecial xlPasteValuesAndNumberFormats
                    rngCom.Offset(0, 11).Resize(iCnt, 1).Copy
                    ws.Range("C2").PasteSpecial xlPasteValuesAndNumberFormats
                    rngCom.Offset(0, 12).Resize(iCnt, 1).Copy
                    ws.Range("D2").PasteSpecial xlPasteValuesAndNumberFormats
                    ws.Range("A2") = 1
                    If ws.Range("B65536").End(xlUp).Row > 2 Then
                        ws.Range("A2").AutoFill ws.Range("A2:A" & ws.Range("B65536").End(xlUp).Row), xlFillSeries
                    End If
                    ws.Range("E2:E" & ws.Range("A65536").End(xlUp).Row) = rng.Offset(0, 1) & rngCom.Offset(0, 9) & "工资"
                    iColCov = 4
                Case "4" ', "7" '中国建设银行 '中国银行
                    Sheet1.Range("U5:X5").Copy ws.Range("A1")
                    rngCom.Offset(0, 11).Resize(iCnt, 1).Copy
                    ws.Range("B2").PasteSpecial xlPasteValuesAndNumberFormats
                    rngCom.Offset(0, 8).Resize(iCnt, 1).Copy
                    ws.Range("C2").PasteSpecial xlPasteValuesAndNumberFormats
                    rngCom.Offset(0, 12).Resize(iCnt, 1).Copy
                    ws.Range("D2").PasteSpecial xlPasteValuesAndNumberFormats
                    ws.Range("A2") = 1
                    If ws.Range("B65536").End(xlUp).Row > 2 Then
                        ws.Range("A2").AutoFill ws.Range("A2:A" & ws.Range("B65536").End(xlUp).Row), xlFillSeries
                    End If
                    iColCov = 4
                Case "5" '中国工商银行
                    Sheet1.Range("U6:AN6").Copy ws.Range("A1")
                    rngCom.Offset(0, 8).Resize(iCnt, 1).Copy
                    ws.Range("N2").PasteSpecial xlPasteValuesAndNumberFormats
                    rngCom.Offset(0, 11).Resize(iCnt, 1).Copy
                    ws.Range("O2").PasteSpecial xlPasteValuesAndNumberFormats
                    rngCom.Offset(0, 12).Resize(iCnt, 1).Copy
                    ws.Range("P2").PasteSpecial xlPasteValuesAndNumberFormats
                    
                    ws.Range("D2") = 1
                    If ws.Range("N65536").End(xlUp).Row > 2 Then
                        ws.Range("D2").AutoFill ws.Range("D2:D" & ws.Range("N65536").End(xlUp).Row), xlFillSeries
                    End If
                    ws.Range("A2:A" & ws.Range("N65536").End(xlUp).Row) = "RMB"
                    ws.Range("G2:G" & ws.Range("N65536").End(xlUp).Row) = "中国工商银行"
                    ws.Range("I2:I" & ws.Range("N65536").End(xlUp).Row) = rngCom.Value
                    ws.Range("J2:J" & ws.Range("N65536").End(xlUp).Row) = "中国工商银行"
                    ws.Range("Q2:Q" & ws.Range("N65536").End(xlUp).Row) = rng.Offset(0, 1) & rngCom.Offset(0, 9) & "工资"
'                    ws.Range("R2:R" & ws.Range("A65536").End(xlUp).Row) = rng.Offset(0, 1) & rngCom.Offset(0, 9) & "工资"
'                    ws.Range("R2:R" & ws.Range("N65536").End(xlUp).Row) = "代发工资"
                    ws.Range("T2:T" & ws.Range("N65536").End(xlUp).Row) = 0
                    Set rngBank = Sheet1.Range("L:L").Find(rngCom.Value)
                    If rngBank Is Nothing Then
                        MsgBox rngCom & "的银行信息未维护，请维护后重新运行！", vbInformation, "中国工商银行信息缺失："
                        Set rngCom = Nothing
                        Set rng = Nothing
                        wb.Close False
                        Set ws = Nothing
                        Set wb = nohting
                        Exit Sub
                    Else
                        ws.Range("H2:H" & ws.Range("N65536").End(xlUp).Row) = rngBank.Offset(0, 1)
                        ws.Range("K2:K" & ws.Range("N65536").End(xlUp).Row) = rngBank.Offset(0, 2)
                        ws.Range("L2:L" & ws.Range("N65536").End(xlUp).Row) = rngBank.Offset(0, 3)
                        ws.Range("M2:M" & ws.Range("N65536").End(xlUp).Row) = rngBank.Offset(0, 4)
                    End If
                    Set rngBank = Nothing
                    iColCov = 4
                Case "6" '中国交通银行
                    Sheet1.Range("U7:AA7").Copy ws.Range("A1")
                    rngCom.Offset(0, 11).Resize(iCnt, 1).Copy
                    ws.Range("A2").PasteSpecial xlPasteValuesAndNumberFormats
                    rngCom.Offset(0, 8).Resize(iCnt, 1).Copy
                    ws.Range("B2").PasteSpecial xlPasteValuesAndNumberFormats
                    rngCom.Offset(0, 12).Resize(iCnt, 1).Copy
                    ws.Range("C2").PasteSpecial xlPasteValuesAndNumberFormats
                    ws.Range("D2:D" & ws.Range("A65536").End(xlUp).Row) = 0
                    iColCov = 4
                Case "8" '中国银行
                    Sheet1.Range("AE9:AM9").Copy ws.Range("A1")
                    Sheet1.Range("U9:AC9").Copy ws.Range("A2")
                    rngCom.Offset(0, 11).Resize(iCnt, 1).Copy
                    ws.Range("B3").PasteSpecial xlPasteValuesAndNumberFormats
                    rngCom.Offset(0, 8).Resize(iCnt, 1).Copy
                    ws.Range("C3").PasteSpecial xlPasteValuesAndNumberFormats
                    rngCom.Offset(0, 12).Resize(iCnt, 1).Copy
                    ws.Range("D3").PasteSpecial xlPasteValuesAndNumberFormats
                    ws.Range("A3") = 1
                    If ws.Range("B65536").End(xlUp).Row > 3 Then
                        ws.Range("A3").AutoFill ws.Range("A3:A" & ws.Range("B65536").End(xlUp).Row), xlFillSeries
                    End If
                    ws.Range("A:A").TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
                            Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 2), TrailingMinusNumbers:=True
                    Set rngBank = Sheet1.Range("K:K").Find(rngCom.Offset(0, -1))
                    If rngBank Is Nothing Then
                        MsgBox rngCom & "的银行信息未维护，请维护后重新运行！", vbInformation, "中国银行信息缺失："
                        Set rngCom = Nothing
                        Set rng = Nothing
                        wb.Close False
                        Set ws = Nothing
                        Set wb = nohting
                        Exit Sub
                    Else
                        ws.Range("E3:E" & ws.Range("A65536").End(xlUp).Row) = rngBank.Offset(0, 2)
                        ws.Range("D1") = rngBank.Offset(0, 1)
                    End If
                    ws.Range("I3:I" & ws.Range("A65536").End(xlUp).Row) = 0
                    ws.Range("I:I").TextToColumns Destination:=Range("I1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
                            Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 2), TrailingMinusNumbers:=True
                    ws.Columns("A:I").EntireColumn.AutoFit
                    ws.UsedRange.Borders(xlEdgeBottom).LineStyle = xlContinuous
                    ws.UsedRange.Borders(xlEdgeTop).LineStyle = xlContinuous
                    ws.UsedRange.Borders(xlEdgeLeft).LineStyle = xlContinuous
                    ws.UsedRange.Borders(xlEdgeRight).LineStyle = xlContinuous
                    ws.UsedRange.Borders(xlInsideVertical).LineStyle = xlContinuous
                    ws.UsedRange.Borders(xlInsideHorizontal).LineStyle = xlContinuous
                    iColCov = 4
                Case "9" '招商银行集团网银
                    Sheet1.Range("U10:Z10").Copy ws.Range("A1")
                    rngCom.Offset(0, 11).Resize(iCnt, 1).Copy
                    ws.Range("A2").PasteSpecial xlPasteValuesAndNumberFormats
                    rngCom.Offset(0, 8).Resize(iCnt, 1).Copy
                    ws.Range("B2").PasteSpecial xlPasteValuesAndNumberFormats
                    rngCom.Offset(0, 12).Resize(iCnt, 1).Copy
                    ws.Range("C2").PasteSpecial xlPasteValuesAndNumberFormats
                    ws.Range("F2:F" & ws.Range("A65536").End(xlUp).Row) = rng.Offset(0, 1) & rngCom.Offset(0, 9) & "工资"
                    'iColCov = 4 updated by hesha on 2020-3-24
                    iColCov = 3
                Case "10" '交通银行
                    Sheet1.Range("U11:W11").Copy ws.Range("A1")
                    rngCom.Offset(0, 11).Resize(iCnt, 1).Copy
                    ws.Range("A2").PasteSpecial xlPasteValuesAndNumberFormats
                    rngCom.Offset(0, 8).Resize(iCnt, 1).Copy
                    ws.Range("B2").PasteSpecial xlPasteValuesAndNumberFormats
                    rngCom.Offset(0, 12).Resize(iCnt, 1).Copy
                    ws.Range("C2").PasteSpecial xlPasteValuesAndNumberFormats
                    iColCov = 4
            End Select
            
            If rng.Offset(0, 5) <> "" Or rng.Offset(0, 4) <> "" Then
            With ws
                '需单笔拆分
                If rng.Offset(0, 4) > 0 Then
                    For iRow = .Range("A65536 ").End(xlUp).Row To 2 Step -1
                        iMult = WorksheetFunction.Ceiling(.Cells(iRow, iColCov) / rng.Offset(0, 4), 1)
                        If iMult > 1 Then
                            iMod = .Cells(iRow, iColCov) Mod rng.Offset(0, 4)
                            For iNew = 1 To iMult - 1
                                iRowCov = .Range("A65536").End(xlUp).Row + 1
                                If iColCov = 3 Then
                                    .Range("A" & iRow & ":B" & iRow).Copy .Range("A" & iRowCov)
                                    .Range("C" & iRowCov) = rng.Offset(0, 4)
                                Else
                                    .Range("A" & iRowCov) = .Range("A" & iRowCov - 1) + 1
                                    .Range("C" & iRow & ":D" & iRow).Copy .Range("A" & iRowCov)
                                    .Range("E" & iRowCov) = rng.Offset(0, 4)
                                End If
                            Next iNew
                            
                            If iMod > 0 Then
                                .Cells(iRow, iColCov) = Round(.Cells(iRow, iColCov) - (rng.Offset(0, 4) * (iMult - 1)), 2)
                            Else
                                .Rows(iRow).Delete
                            End If
                        End If
                    Next iRow
                    .Range("F2").AutoFill .Range("F2:F" & .Range("A65536").End(xlUp).Row), xlFillCopy
                End If
'                '需整个文件拆分
'                If rng.Offset(0, 5) > 0 And WorksheetFunction.Sum(.Range(.Cells(2, iColCov), .Cells(ws.Range("A65536").End(xlUp).Row, iColCov))) > rng.Offset(0, 4) Then
'                    With ws
'                        iMult = WorksheetFunction.Ceiling(iMod / rng.Offset(0, 5), 1)
'                        If iMult > 1 Then
'                            For iNew = 1 To iMult - 1
'                                Set wbSplit = Workbooks.Add(1)
'                                Set wsSplit = wbSplit.ActiveSheet
'                                iMod = .Range("C" & .Range("A65536").End(xlUp).Row)
'                                For iRow = .Range("A65536").End(xlUp).Row - 1 To 2
'                                    iMod = iMod + .Range("C" & iRow) + .Range("C" & iRow + 1)
'
'                                Next iRow
'                            Next iNew
'                        End If
'                    End With
'                End If
            End With
            End If
            
'            If rng.Offset(0, 2) = "中国农业银行" Or rng.Offset(0, 2) = "中国工商银行" Then
'                wb.SaveAs StrPt & "报盘文件" & rng.Offset(0, 1) & "-" & rngCom.Offset(0, 9), xlCSV
'            Else
                wb.SaveAs StrPt & rng.Offset(0, 2) & "-" & rng.Offset(0, 1) & "-" & rngCom.Offset(0, 9), xlWorkbookDefault
'            End If
            wb.Close True
            Set rng = Nothing
            Set rngCom = Nothing
            Set ws = Nothing
            Set wb = Nothing
        End If
    Next iLoop
End With

With Sheet3
If Sheet3.UsedRange.Rows.Count > 1 Then
    Sheet1.Range("J:J").ClearContents
    .Range("A2:A" & .Range("A65536").End(xlUp).Row).Copy Sheet1.Range("J1")
    Sheet1.Range("J:J").RemoveDuplicates 1, xlNo
    For iLoop = 1 To Sheet1.Range("J65536").End(xlUp).Row
        Set rng = Sheet1.Range("A:A").Find(Sheet1.Range("J" & iLoop), , , xlWhole)
        If Not rng Is Nothing Then
            Set wb = Workbooks.Add(1)
            Set ws = wb.ActiveSheet
            iCnt = WorksheetFunction.CountIf(.Range("A:A"), rng.Value)
            Set rngCom = .Range("A:A").Find(rng.Value)
            
            Sheet1.Range("U2:Z2").Copy ws.Range("A1")
            rngCom.Offset(0, 8).Resize(rngCom.Row + iCnt - 1, 1).Copy
            ws.Range("A2").PasteSpecial xlPasteValuesAndNumberFormats
            rngCom.Offset(0, 11).Resize(rngCom.Row + iCnt - 1, 1).Copy
            ws.Range("B2").PasteSpecial xlPasteValuesAndNumberFormats
            rngCom.Offset(0, 12).Resize(rngCom.Row + iCnt - 1, 1).Copy
            ws.Range("C2").PasteSpecial xlPasteValuesAndNumberFormats
            
            wb.SaveAs StrPt & "外籍报盘文件" & rng.Offset(0, 1) & "-" & rngCom.Offset(0, 9), xlWorkbookDefault
            wb.Close True
            Set rng = Nothing
            Set rngCom = Nothing
            Set ws = Nothing
            Set wb = Nothing
        End If
    Next iLoop
End If
End With

MsgBox "报盘文件已全部生成！", vbInformation, "完成："
End Sub