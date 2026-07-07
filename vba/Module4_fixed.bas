Attribute VB_Name = "Module4"
Sub CreatePredictionLocationInstruction()
    Dim wsData As Worksheet
    Dim wsOut As Worksheet
    Dim lastRow As Long, i As Long

    ' --- 1. "CF" シートの取得 ---
    On Error Resume Next
    Set wsData = ActiveWorkbook.Sheets("CF")
    On Error GoTo 0

    If wsData Is Nothing Then
        Set wsData = ActiveSheet
        If MsgBox("「CF」シートが見つかりません。現在のシート（" & wsData.name & "）を処理しますか？", vbYesNo + vbQuestion) = vbNo Then
            Exit Sub
        End If
    End If

    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).row

    ' 列の特定（号機、段、列、品名、品名コード、投入回数_予測）
    Dim cMach As Integer, cDan As Integer, cRetsu As Integer, cName As Integer, cCode As Integer, cPred As Integer
    On Error Resume Next
    cMach = WorksheetFunction.Match("号機", wsData.Rows(1), 0)
    cDan = WorksheetFunction.Match("段", wsData.Rows(1), 0)
    cRetsu = WorksheetFunction.Match("列", wsData.Rows(1), 0)
    cName = WorksheetFunction.Match("品名", wsData.Rows(1), 0)
    cCode = WorksheetFunction.Match("品名コード", wsData.Rows(1), 0)
    cPred = WorksheetFunction.Match("投入回数_予測", wsData.Rows(1), 0)
    On Error GoTo 0

    If cMach = 0 Or cName = 0 Or cPred = 0 Then
        MsgBox "必須列（号機、品名、投入回数_予測）が見つかりません。1行目にヘッダーがあるか確認してください。", vbCritical
        Exit Sub
    End If

    ' --- 2. "号機回数比" シートから 号機ごとの実績比率と目標比率の乖離を取得 ---
    ' （AB01～AB46＝1～46号機に対応。乖離=実績比率(D列)-目標比率(E列)。正=目標超過(激務)、負=目標未達(暇)）
    Dim wsRatio As Worksheet
    On Error Resume Next
    Set wsRatio = ActiveWorkbook.Sheets("号機回数比")
    On Error GoTo 0

    If wsRatio Is Nothing Then
        MsgBox "「号機回数比」シートが見つかりません。先にこのシートを作成してください。", vbCritical
        Exit Sub
    End If

    Dim machDev(1 To 46) As Double
    Dim rr As Long, abLabel As String, mNum As Integer
    For rr = 3 To 48 ' AB01(1号機)～AB46(46号機)に対応する行
        abLabel = Trim(CStr(wsRatio.Cells(rr, 1).Value))
        If abLabel Like "AB##" Then
            mNum = CInt(Mid(abLabel, 3, 2))
            If mNum >= 1 And mNum <= 46 Then
                machDev(mNum) = Val(wsRatio.Cells(rr, 4).Value) - Val(wsRatio.Cells(rr, 5).Value)
            End If
        End If
    Next rr

    ' --- 出荷予測回数順位900位のカットオフ値を算出（降格対象の判定基準、全アイテムが母集団） ---
    Dim allPredArr() As Double
    Dim allCount As Long: allCount = 0
    ReDim allPredArr(1 To lastRow)
    For i = 2 To lastRow
        If IsNumeric(wsData.Cells(i, cMach).Value) Then
            Dim machChk As Integer: machChk = CInt(wsData.Cells(i, cMach).Value)
            If Not IsExcludedSlot(machChk, CInt(Val(wsData.Cells(i, cRetsu).Value))) Then
                allCount = allCount + 1
                allPredArr(allCount) = Val(wsData.Cells(i, cPred).Value)
            End If
        End If
    Next i
    If allCount > 0 Then ReDim Preserve allPredArr(1 To allCount)

    Dim demoteCutoff As Double
    If allCount >= 900 Then
        demoteCutoff = Application.WorksheetFunction.Large(allPredArr, 900)
    ElseIf allCount > 0 Then
        demoteCutoff = Application.WorksheetFunction.Min(allPredArr) ' 全アイテム件数が900未満なら全件を対象
    Else
        demoteCutoff = 0
    End If

    ' ① 号機回数比で目標を大きく超えている号機トップ3（激務＝降格元）
    Dim busyRank(1 To 3) As Integer, maxDev As Double
    Dim j As Integer, mm As Integer
    For j = 1 To 3
        maxDev = -99999
        For mm = 1 To 46
            If machDev(mm) > maxDev Then
                If mm <> busyRank(1) And mm <> busyRank(2) And mm <> busyRank(3) Then
                    maxDev = machDev(mm): busyRank(j) = mm
                End If
            End If
        Next mm
    Next j

    ' ② 号機回数比で目標を大きく下回っている号機トップ3（暇＝受け入れ先）
    Dim idleRank(1 To 3) As Integer, minDev As Double
    For j = 1 To 3
        minDev = 99999
        For mm = 1 To 46
            If machDev(mm) < minDev Then
                If mm <> idleRank(1) And mm <> idleRank(2) And mm <> idleRank(3) Then
                    minDev = machDev(mm): idleRank(j) = mm
                End If
            End If
        Next mm
    Next j
    Dim safeMachs As String
    safeMachs = idleRank(1) & "号機, " & idleRank(2) & "号機, " & idleRank(3) & "号機"

    ' --- 3. 出力シートの作成 ---
    Application.DisplayAlerts = False
    On Error Resume Next
    ActiveWorkbook.Sheets("ゾーンバランス最適").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    ' 「操作パネル」シートがあればその左側に配置する（無ければ従来通りCFシートの右隣）
    Dim wsPanel4 As Worksheet
    On Error Resume Next
    Set wsPanel4 = ActiveWorkbook.Sheets("操作パネル")
    On Error GoTo 0
    If Not wsPanel4 Is Nothing Then
        Set wsOut = ActiveWorkbook.Sheets.Add(Before:=wsPanel4)
    Else
        Set wsOut = ActiveWorkbook.Sheets.Add(After:=wsData)
    End If
    wsOut.name = "ゾーンバランス最適"
    wsOut.Cells.NumberFormat = "@" ' 日付化の完全ブロック
    wsOut.Columns("D:D").NumberFormat = "0" ' 商品コード列だけは数値表示に戻す

    Dim r As Long: r = 1
    ' タイトル・見出し・説明文はA:G列で結合し、A列だけが横に伸びないようにする
    wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, 7)).Merge
    wsOut.Cells(r, 1).Value = "【予測データ活用：事前ロケーション変更指示書（号機回数比 目標比率ベース）】"
    wsOut.Cells(r, 1).Font.Bold = True: wsOut.Cells(r, 1).Font.Size = 16: wsOut.Cells(r, 1).HorizontalAlignment = xlLeft
    r = r + 2

    wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, 7)).Merge
    wsOut.Cells(r, 1).Value = "避難先・受け入れ先（号機回数比で目標より投入回数が少ない号機トップ3）"
    wsOut.Cells(r, 1).Font.Bold = True: wsOut.Cells(r, 1).Font.Color = RGB(0, 100, 0): wsOut.Cells(r, 1).HorizontalAlignment = xlLeft
    r = r + 1
    wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, 7)).Merge
    wsOut.Cells(r, 1).Value = "以下の号機の空き間口を、移動の受け入れ先にしてください: " & safeMachs & _
        "（乖離: " & Format(machDev(idleRank(1)), "+0.0%;-0.0%") & " / " & Format(machDev(idleRank(2)), "+0.0%;-0.0%") & " / " & Format(machDev(idleRank(3)), "+0.0%;-0.0%") & "）"
    wsOut.Cells(r, 1).HorizontalAlignment = xlLeft
    r = r + 2

    ' --------------------------------------------------
    ' Part 1: 目標超過号機からの上位品の分散
    ' --------------------------------------------------
    wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, 7)).Merge
    wsOut.Cells(r, 1).Value = "1. 号機回数比 目標超過号機からの分散"
    wsOut.Cells(r, 1).Font.Bold = True: wsOut.Cells(r, 1).Font.Color = RGB(200, 0, 0): wsOut.Cells(r, 1).HorizontalAlignment = xlLeft
    r = r + 1
    wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, 7)).Merge
    wsOut.Cells(r, 1).Value = "※号機回数比シートで目標比率を大きく超えている号機（" & busyRank(1) & "号機, " & busyRank(2) & "号機, " & busyRank(3) & "号機）です。売れ筋を受け入れ先へ逃がしてください。"
    wsOut.Cells(r, 1).HorizontalAlignment = xlLeft
    r = r + 1

    wsOut.Cells(r, 1).Resize(1, 7).Value = Array("対象号機(乖離)", "現在ロケ", "品名", "商品コード", "予測回数", "現在の号機", "移動指示")
    wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, 7)).Interior.Color = RGB(255, 230, 230)
    wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, 7)).Font.Bold = True
    r = r + 1

    Dim targetItems() As Variant
    Dim itemCount As Long, k As Long
    Dim tmp1, tmp2, tmp3, tmp4, tmp5, tmp6, a As Long, b As Long
    Dim mach As Integer, pred As Long

    For j = 1 To 3
        Dim targetMach As Integer: targetMach = busyRank(j)
        itemCount = 0
        ReDim targetItems(1 To lastRow, 1 To 6)

        For i = 2 To lastRow
            If IsNumeric(wsData.Cells(i, cMach).Value) Then
                mach = CInt(wsData.Cells(i, cMach).Value)
                Dim clChk As Integer: clChk = CInt(Val(wsData.Cells(i, cRetsu).Value))
                If mach = targetMach And Not IsExcludedSlot(mach, clChk) Then
                    pred = Val(wsData.Cells(i, cPred).Value)
                    If pred >= 30 Then ' 30回以上の大物のみ対象
                        itemCount = itemCount + 1
                        targetItems(itemCount, 1) = Format(mach, "00") & "-" & Format(Val(wsData.Cells(i, cDan).Value), "00") & "-" & Format(Val(wsData.Cells(i, cRetsu).Value), "00")
                        targetItems(itemCount, 2) = wsData.Cells(i, cName).Value
                        targetItems(itemCount, 3) = GetCodeValue4(wsData, i, cCode)
                        targetItems(itemCount, 4) = pred
                        targetItems(itemCount, 5) = mach & "号機"
                        targetItems(itemCount, 6) = pred
                    End If
                End If
            End If
        Next i

        ' 降順ソート
        For a = 1 To itemCount - 1
            For b = a + 1 To itemCount
                If targetItems(a, 6) < targetItems(b, 6) Then
                    tmp1 = targetItems(a, 1): targetItems(a, 1) = targetItems(b, 1): targetItems(b, 1) = tmp1
                    tmp2 = targetItems(a, 2): targetItems(a, 2) = targetItems(b, 2): targetItems(b, 2) = tmp2
                    tmp3 = targetItems(a, 3): targetItems(a, 3) = targetItems(b, 3): targetItems(b, 3) = tmp3
                    tmp4 = targetItems(a, 4): targetItems(a, 4) = targetItems(b, 4): targetItems(b, 4) = tmp4
                    tmp5 = targetItems(a, 5): targetItems(a, 5) = targetItems(b, 5): targetItems(b, 5) = tmp5
                    tmp6 = targetItems(a, 6): targetItems(a, 6) = targetItems(b, 6): targetItems(b, 6) = tmp6
                End If
            Next b
        Next a

        For a = 1 To itemCount
            If a > 3 Then Exit For ' 各号機上位3つまで
            wsOut.Cells(r, 1).Value = targetMach & "号機(" & Format(machDev(targetMach), "+0.0%;-0.0%") & ")"
            wsOut.Cells(r, 2).Value = targetItems(a, 1)
            wsOut.Cells(r, 3).Value = targetItems(a, 2)
            wsOut.Cells(r, 4).Value = targetItems(a, 3)
            wsOut.Cells(r, 5).Value = targetItems(a, 4) & "回"
            wsOut.Cells(r, 6).Value = targetItems(a, 5)
            wsOut.Cells(r, 7).Value = safeMachs & " のいずれかへ"
            r = r + 1
        Next a
    Next j
    r = r + 1

    ' --------------------------------------------------
    ' Part 2: Cエリアからの昇格
    ' --------------------------------------------------
    wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, 7)).Merge
    wsOut.Cells(r, 1).Value = "2. 昇格（Cエリア 51-68号機 → ABエリアの目標未達号機へ）"
    wsOut.Cells(r, 1).Font.Bold = True: wsOut.Cells(r, 1).Font.Color = RGB(0, 0, 200): wsOut.Cells(r, 1).HorizontalAlignment = xlLeft
    r = r + 1
    wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, 7)).Merge
    wsOut.Cells(r, 1).Value = "※Cエリアの隠れた売れ筋をABエリアの目標未達号機に移し、全体のバランスを底上げします。"
    wsOut.Cells(r, 1).HorizontalAlignment = xlLeft
    r = r + 1

    wsOut.Cells(r, 1).Resize(1, 7).Value = Array("対象", "現在ロケ", "品名", "商品コード", "予測回数", "現在の号機", "移動指示")
    wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, 7)).Interior.Color = RGB(200, 230, 255)
    wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, 7)).Font.Bold = True
    r = r + 1

    itemCount = 0
    ReDim targetItems(1 To lastRow, 1 To 6)
    For i = 2 To lastRow
        If IsNumeric(wsData.Cells(i, cMach).Value) Then
            mach = CInt(wsData.Cells(i, cMach).Value)
            If mach >= 51 And mach <= 68 Then
                pred = Val(wsData.Cells(i, cPred).Value)
                If pred >= 10 Then
                    itemCount = itemCount + 1
                    targetItems(itemCount, 1) = Format(mach, "00") & "-" & Format(Val(wsData.Cells(i, cDan).Value), "00") & "-" & Format(Val(wsData.Cells(i, cRetsu).Value), "00")
                    targetItems(itemCount, 2) = wsData.Cells(i, cName).Value
                    targetItems(itemCount, 3) = GetCodeValue4(wsData, i, cCode)
                    targetItems(itemCount, 4) = pred
                    targetItems(itemCount, 5) = mach & "号機 (Cエリア)"
                    targetItems(itemCount, 6) = pred
                End If
            End If
        End If
    Next i

    For a = 1 To itemCount - 1
        For b = a + 1 To itemCount
            If targetItems(a, 6) < targetItems(b, 6) Then
                tmp1 = targetItems(a, 1): targetItems(a, 1) = targetItems(b, 1): targetItems(b, 1) = tmp1
                tmp2 = targetItems(a, 2): targetItems(a, 2) = targetItems(b, 2): targetItems(b, 2) = tmp2
                tmp3 = targetItems(a, 3): targetItems(a, 3) = targetItems(b, 3): targetItems(b, 3) = tmp3
                tmp4 = targetItems(a, 4): targetItems(a, 4) = targetItems(b, 4): targetItems(b, 4) = tmp4
                tmp5 = targetItems(a, 5): targetItems(a, 5) = targetItems(b, 5): targetItems(b, 5) = tmp5
                tmp6 = targetItems(a, 6): targetItems(a, 6) = targetItems(b, 6): targetItems(b, 6) = tmp6
            End If
        Next b
    Next a

    For a = 1 To itemCount
        If a > 10 Then Exit For ' 上位10アイテムまで
        wsOut.Cells(r, 1).Value = "昇格"
        wsOut.Cells(r, 2).Value = targetItems(a, 1)
        wsOut.Cells(r, 3).Value = targetItems(a, 2)
        wsOut.Cells(r, 4).Value = targetItems(a, 3)
        wsOut.Cells(r, 5).Value = targetItems(a, 4) & "回"
        wsOut.Cells(r, 6).Value = targetItems(a, 5)
        wsOut.Cells(r, 7).Value = safeMachs & " のいずれかへ"
        r = r + 1
    Next a
    r = r + 1

    ' --------------------------------------------------
    ' Part 3: 目標超過号機からの下位品の降格
    ' --------------------------------------------------
    wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, 7)).Merge
    wsOut.Cells(r, 1).Value = "3. 降格（ABエリア下位品 → Cエリア 51-68号機へ）"
    wsOut.Cells(r, 1).Font.Bold = True: wsOut.Cells(r, 1).Font.Color = RGB(150, 0, 150): wsOut.Cells(r, 1).HorizontalAlignment = xlLeft
    r = r + 1
    wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, 7)).Merge
    wsOut.Cells(r, 1).Value = "※全アイテム中の出荷予測回数順位900位以下（予測回数" & demoteCutoff & "回以下）に該当するABエリアの死に筋を抜き、Cエリアへ降格します。"
    wsOut.Cells(r, 1).HorizontalAlignment = xlLeft
    r = r + 1

    wsOut.Cells(r, 1).Resize(1, 7).Value = Array("対象", "現在ロケ", "品名", "商品コード", "予測回数", "現在の号機", "移動指示")
    wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, 7)).Interior.Color = RGB(240, 200, 240)
    wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, 7)).Font.Bold = True
    r = r + 1

    itemCount = 0
    ReDim targetItems(1 To lastRow, 1 To 6)
    For i = 2 To lastRow
        If IsNumeric(wsData.Cells(i, cMach).Value) Then
            mach = CInt(wsData.Cells(i, cMach).Value)
            Dim clChk2 As Integer: clChk2 = CInt(Val(wsData.Cells(i, cRetsu).Value))
            If mach >= 1 And mach <= 46 And Not IsExcludedSlot(mach, clChk2) Then ' ABエリア全体が対象（激務号機に限定しない、中量棚は除外）
                pred = Val(wsData.Cells(i, cPred).Value)
                If pred <= demoteCutoff Then ' 全アイテム中の出荷予測回数順位900位以下（下位）の死に筋のみ
                    itemCount = itemCount + 1
                    targetItems(itemCount, 1) = Format(mach, "00") & "-" & Format(Val(wsData.Cells(i, cDan).Value), "00") & "-" & Format(Val(wsData.Cells(i, cRetsu).Value), "00")
                    targetItems(itemCount, 2) = wsData.Cells(i, cName).Value
                    targetItems(itemCount, 3) = GetCodeValue4(wsData, i, cCode)
                    targetItems(itemCount, 4) = pred
                    targetItems(itemCount, 5) = mach & "号機"
                    targetItems(itemCount, 6) = pred
                End If
            End If
        End If
    Next i

    ' 少ない順にソート（昇順）
    For a = 1 To itemCount - 1
        For b = a + 1 To itemCount
            If targetItems(a, 6) > targetItems(b, 6) Then
                tmp1 = targetItems(a, 1): targetItems(a, 1) = targetItems(b, 1): targetItems(b, 1) = tmp1
                tmp2 = targetItems(a, 2): targetItems(a, 2) = targetItems(b, 2): targetItems(b, 2) = tmp2
                tmp3 = targetItems(a, 3): targetItems(a, 3) = targetItems(b, 3): targetItems(b, 3) = tmp3
                tmp4 = targetItems(a, 4): targetItems(a, 4) = targetItems(b, 4): targetItems(b, 4) = tmp4
                tmp5 = targetItems(a, 5): targetItems(a, 5) = targetItems(b, 5): targetItems(b, 5) = tmp5
                tmp6 = targetItems(a, 6): targetItems(a, 6) = targetItems(b, 6): targetItems(b, 6) = tmp6
            End If
        Next b
    Next a

    For a = 1 To itemCount
        If a > 10 Then Exit For ' 10アイテムまで
        wsOut.Cells(r, 1).Value = "降格"
        wsOut.Cells(r, 2).Value = targetItems(a, 1)
        wsOut.Cells(r, 3).Value = targetItems(a, 2)
        wsOut.Cells(r, 4).Value = targetItems(a, 3)
        wsOut.Cells(r, 5).Value = targetItems(a, 4) & "回"
        wsOut.Cells(r, 6).Value = targetItems(a, 5)
        wsOut.Cells(r, 7).Value = "Cエリア(51～68号機)の空き枠へ"
        r = r + 1
    Next a

    wsOut.Columns("A:G").AutoFit
    wsOut.Activate

    ' ※本マクロは予測データベースのため、KPI記録（実績データのみ対象）には記録しません

    MsgBox "号機回数比の目標比率に基づく「ゾーンバランス最適」の作成が完了しました！", vbInformation
End Sub

' 1号機・2号機の6～14列は中量棚のため、AB通常ロケーションの対象から除外する
Function IsExcludedSlot(mach As Integer, col As Integer) As Boolean
    IsExcludedSlot = (mach = 1 Or mach = 2) And col >= 6 And col <= 14
End Function

' 品名コードを数値として取得する（列が見つからない、または数値でない場合は空文字）
Function GetCodeValue4(wsData As Worksheet, ByVal rowIdx As Long, ByVal cCode As Integer) As Variant
    If cCode = 0 Then
        GetCodeValue4 = ""
        Exit Function
    End If
    Dim v As Variant: v = wsData.Cells(rowIdx, cCode).Value
    If IsNumeric(v) Then
        GetCodeValue4 = CLng(v)
    Else
        GetCodeValue4 = v
    End If
End Function
