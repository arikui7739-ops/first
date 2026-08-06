Attribute VB_Name = "Module10_石狩"
Option Explicit

' ----------------------------------------------------
' 予測データに基づくロケ変指示の作成
' 「予測データ」シート(Module8で取込済み)の機番別実績(予測)構成比を、「設定」シートの
' 「■機番別目標構成比」に近づけるためのロケーション変更(入替)指示を作成する。
' 入替の相手先は必ず同じ段(例:46-02-05なら2段目)の中から選ぶ(段をまたぐ入替は行わない)。
' 除外機番・除外ロケーション・除外品コードは「設定」シートの設定に従う(Module3と共通)。
' 候補件数は「設定」シートの「ロケ変候補件数」で個別に設定できる(AB編成動線最適化の
' 「入替候補件数」とは別の設定)。
' 既存の同名シートは削除してから作り直すため、再実行すると内容が更新される
' ----------------------------------------------------

Sub CreateRelocationPlan()
    Call EnsureRelocationPlanButton

    Dim wsData As Worksheet
    On Error Resume Next
    Set wsData = ThisWorkbook.Sheets("予測データ")
    On Error GoTo 0
    If wsData Is Nothing Then
        MsgBox "「予測データ」シートが見つかりません。先に「予測データを取り込む」を実行してください。", vbExclamation
        Exit Sub
    End If

    Call EnsureExclusionSettingsSheet

    Dim dictExcludedMach As Object: Set dictExcludedMach = CreateObject("Scripting.Dictionary")
    Dim locMach() As Long, locDanFrom() As Long, locDanTo() As Long, locColFrom() As Long, locColTo() As Long
    Dim locCount As Long
    Dim dictExcludedItemCode As Object: Set dictExcludedItemCode = CreateObject("Scripting.Dictionary")
    Dim ratioSheetName As String, maxSwapRows As Long, maxMachNum As Long, abSlotCount As Long
    Dim dictTargetRatio As Object: Set dictTargetRatio = CreateObject("Scripting.Dictionary")
    Call LoadExclusionSettings(dictExcludedMach, locMach, locDanFrom, locDanTo, locColFrom, locColTo, locCount, dictExcludedItemCode, ratioSheetName, maxSwapRows, maxMachNum, abSlotCount, dictTargetRatio)

    Dim relocCount As Long: relocCount = GetRelocationCandidateCount()

    ' 「機番別目標構成比」に数値機番キーが1つも無ければ、近づけるべき目標が無いためロケ変指示を作成できない
    Dim hasAnyTarget As Boolean: hasAnyTarget = False
    Dim chkKey As Variant
    For Each chkKey In dictTargetRatio.Keys
        If IsNumeric(chkKey) Then hasAnyTarget = True: Exit For
    Next chkKey
    If Not hasAnyTarget Then
        MsgBox "「設定」シートの「■機番別目標構成比」に機番別の目標値が入力されていないため、ロケ変指示を作成できません。", vbExclamation
        Exit Sub
    End If

    ' 「予測データ」シートは1行目=取込情報、3行目=見出し(Module8の出力形式)。
    ' 機番・段・列・品名コード・投入回数_予測の列を見出し名から探す(品名列があれば表示にも使う)
    Const HEADER_ROW As Long = 3
    Dim lastRow As Long: lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    Dim lastCol As Long: lastCol = wsData.Cells(HEADER_ROW, wsData.Columns.Count).End(xlToLeft).Column

    Dim machColIdx As Long: machColIdx = -1
    Dim danColIdx As Long: danColIdx = -1
    Dim colColIdx As Long: colColIdx = -1
    Dim itemCodeColIdx As Long: itemCodeColIdx = -1
    Dim itemNameColIdx As Long: itemNameColIdx = -1
    Dim cntColIdx As Long: cntColIdx = -1
    Dim hc As Long
    For hc = 1 To lastCol
        Dim hName As String: hName = Trim(CStr(wsData.Cells(HEADER_ROW, hc).Value))
        If hName = "号機" Then machColIdx = hc
        If hName = "段" Then danColIdx = hc
        If hName = "列" Then colColIdx = hc
        If hName = "品名コード" Then itemCodeColIdx = hc
        If hName = "品名" Then itemNameColIdx = hc
        If hName = "投入回数_予測" Then cntColIdx = hc
    Next hc
    If machColIdx = -1 Or danColIdx = -1 Or colColIdx = -1 Or itemCodeColIdx = -1 Or cntColIdx = -1 Then
        MsgBox "「予測データ」シートに「号機」「段」「列」「品名コード」「投入回数_予測」のいずれかの列が見つかりません。", vbExclamation
        Exit Sub
    End If

    ' 対象行(除外条件に該当しないAB編成対象のロケーション)を読み込む
    Dim rowMaxN As Long: rowMaxN = lastRow - HEADER_ROW
    If rowMaxN < 1 Then rowMaxN = 1
    Dim rowMach() As Long, rowDan() As Long, rowCol() As Long, rowItemCode() As String, rowItemName() As String, rowCnt() As Double, rowUsed() As Boolean
    ReDim rowMach(1 To rowMaxN)
    ReDim rowDan(1 To rowMaxN)
    ReDim rowCol(1 To rowMaxN)
    ReDim rowItemCode(1 To rowMaxN)
    ReDim rowItemName(1 To rowMaxN)
    ReDim rowCnt(1 To rowMaxN)
    ReDim rowUsed(1 To rowMaxN)
    Dim rowN As Long: rowN = 0

    Dim actualCountByMach As Object: Set actualCountByMach = CreateObject("Scripting.Dictionary")
    Dim dictByDan As Object: Set dictByDan = CreateObject("Scripting.Dictionary") ' 段→その段のロケーション行(添字)のCollection
    Dim dictByMach As Object: Set dictByMach = CreateObject("Scripting.Dictionary") ' 機番→その機番のロケーション行(添字)のCollection

    Dim r As Long
    For r = HEADER_ROW + 1 To lastRow
        If IsNumeric(wsData.Cells(r, machColIdx).Value) Then
            Dim mach As Long: mach = CLng(wsData.Cells(r, machColIdx).Value)
            If mach >= 1 And mach <= maxMachNum Then
                If Not IsExcludedSlot3(dictExcludedMach, CInt(mach)) Then
                    Dim dan As Long: dan = 0
                    If IsNumeric(wsData.Cells(r, danColIdx).Value) Then dan = CLng(wsData.Cells(r, danColIdx).Value)
                    Dim colv As Long: colv = 0
                    If IsNumeric(wsData.Cells(r, colColIdx).Value) Then colv = CLng(wsData.Cells(r, colColIdx).Value)

                    If Not IsExcludedLocation(CInt(mach), CInt(dan), CInt(colv), locMach, locDanFrom, locDanTo, locColFrom, locColTo, locCount) Then
                        Dim itemCodeStr As String: itemCodeStr = Trim(CStr(wsData.Cells(r, itemCodeColIdx).Value))
                        Dim isItemExcluded As Boolean: isItemExcluded = False
                        If itemCodeStr <> "" And dictExcludedItemCode.Count > 0 Then
                            isItemExcluded = dictExcludedItemCode.Exists(itemCodeStr)
                            If Not isItemExcluded And IsNumeric(itemCodeStr) Then isItemExcluded = dictExcludedItemCode.Exists(CStr(CLng(itemCodeStr)))
                        End If

                        If Not isItemExcluded Then
                            Dim forecastCnt As Double: forecastCnt = Val(wsData.Cells(r, cntColIdx).Value)

                            rowN = rowN + 1
                            rowMach(rowN) = mach
                            rowDan(rowN) = dan
                            rowCol(rowN) = colv
                            rowItemCode(rowN) = itemCodeStr
                            rowItemName(rowN) = IIf(itemNameColIdx > 0, Trim(CStr(wsData.Cells(r, itemNameColIdx).Value)), "")
                            rowCnt(rowN) = forecastCnt
                            rowUsed(rowN) = False

                            Dim machKey As String: machKey = CStr(mach)
                            If actualCountByMach.Exists(machKey) Then
                                actualCountByMach(machKey) = actualCountByMach(machKey) + forecastCnt
                            Else
                                actualCountByMach.Add machKey, forecastCnt
                            End If

                            Dim danKey As String: danKey = CStr(dan)
                            If Not dictByDan.Exists(danKey) Then dictByDan.Add danKey, New Collection
                            dictByDan(danKey).Add rowN

                            If Not dictByMach.Exists(machKey) Then dictByMach.Add machKey, New Collection
                            dictByMach(machKey).Add rowN
                        End If
                    End If
                End If
            End If
        End If
    Next r

    If rowN = 0 Then
        MsgBox "「予測データ」シートに、除外条件を除いた対象ロケーションが見つかりませんでした。", vbExclamation
        Exit Sub
    End If

    ' 目標構成比が入力されている機番だけを対象に、合計が100%でなくても相対バランスとして比較できるよう正規化する
    ' (機番別目標構成比の合計と、目標が設定されている機番だけの実績(予測)合計、それぞれで正規化してから比較する)
    Dim targetRatioSum As Double: targetRatioSum = 0
    Dim targetedHitTotal As Double: targetedHitTotal = 0
    Dim trKey As Variant
    For Each trKey In dictTargetRatio.Keys
        If IsNumeric(trKey) Then
            Dim trMach As Long: trMach = CLng(trKey)
            If trMach >= 1 And trMach <= maxMachNum Then
                targetRatioSum = targetRatioSum + dictTargetRatio(trKey)
                If actualCountByMach.Exists(CStr(trMach)) Then targetedHitTotal = targetedHitTotal + actualCountByMach(CStr(trMach))
            End If
        End If
    Next trKey
    If targetRatioSum <= 0 Or targetedHitTotal <= 0 Then
        MsgBox "目標構成比が設定されている機番の実績(予測)データが見つからなかったため、ロケ変指示を作成できません。", vbExclamation
        Exit Sub
    End If

    ' 入替候補の選定(目標構成比を最も上回っている機番の高頻度ロケーションを、同じ段の中から
    ' 最も目標構成比を下回っている機番の未使用ロケーションと入れ替える、を繰り返す)
    Dim outArr() As Variant
    ReDim outArr(1 To relocCount, 1 To 12)
    Dim outCnt As Long: outCnt = 0
    Dim dictGivenUpMach As Object: Set dictGivenUpMach = CreateObject("Scripting.Dictionary") ' これ以上有効な交換相手が見つからない機番

    Dim safetyCounter As Long: safetyCounter = 0
    Dim maxAttempts As Long: maxAttempts = relocCount * 20 + 100 ' 無限ループ防止の安全策

    Do While outCnt < relocCount
        safetyCounter = safetyCounter + 1
        If safetyCounter > maxAttempts Then Exit Do

        ' 現時点で目標構成比を最も上回っていて、かつまだ諦めていない機番を探す
        Dim bestOverMach As String: bestOverMach = ""
        Dim bestOverDev As Double: bestOverDev = 0.0000001 ' 誤差程度の超過は無視する
        For Each trKey In dictTargetRatio.Keys
            If IsNumeric(trKey) Then
                Dim tMach As Long: tMach = CLng(trKey)
                If tMach >= 1 And tMach <= maxMachNum Then
                    If Not dictGivenUpMach.Exists(CStr(tMach)) And dictByMach.Exists(CStr(tMach)) Then
                        If HasUnusedRow(dictByMach(CStr(tMach)), rowUsed) Then
                            Dim actCnt As Double: actCnt = 0
                            If actualCountByMach.Exists(CStr(tMach)) Then actCnt = actualCountByMach(CStr(tMach))
                            Dim actRatio As Double: actRatio = actCnt / targetedHitTotal
                            Dim tgtRatio As Double: tgtRatio = dictTargetRatio(trKey) / targetRatioSum
                            Dim devVal As Double: devVal = actRatio - tgtRatio
                            If devVal > bestOverDev Then
                                bestOverDev = devVal
                                bestOverMach = CStr(tMach)
                            End If
                        End If
                    End If
                End If
            End If
        Next trKey

        If bestOverMach = "" Then Exit Do ' これ以上、目標構成比を上回っていて動かせる機番が無い

        ' 超過機番の未使用ロケーションを予測回数の多い順に試し、同じ段で目標未達の機番の
        ' 未使用ロケーションが見つかるまで探す(超過側への影響が大きい候補から優先的に解消する)
        Dim srcRow As Long: srcRow = FindBestSourceRow(dictByMach(bestOverMach), rowUsed, rowCnt)
        Dim partnerRow As Long: partnerRow = 0
        Do While srcRow > 0 And partnerRow = 0
            partnerRow = FindBestPartnerRow(dictByDan(CStr(rowDan(srcRow))), rowUsed, rowMach, CLng(bestOverMach), actualCountByMach, dictTargetRatio, targetRatioSum, targetedHitTotal, maxMachNum)
            If partnerRow = 0 Then
                Dim tmpUsedMark As Long: tmpUsedMark = srcRow
                rowUsed(tmpUsedMark) = True ' この候補は今回使えないので一時的に使用済み扱いにして次を探す
                srcRow = FindBestSourceRow(dictByMach(bestOverMach), rowUsed, rowCnt)
                rowUsed(tmpUsedMark) = False ' 出力しない場合は使用済みフラグを戻す(以後の反復で再検討できるようにする)
            End If
        Loop

        If partnerRow = 0 Then
            dictGivenUpMach(bestOverMach) = True ' この機番からはこれ以上有効な交換相手が見つからない
        Else
            Dim srcMachV As Long: srcMachV = rowMach(srcRow)
            Dim dstMachV As Long: dstMachV = rowMach(partnerRow)
            Dim srcRatioBefore As Double: srcRatioBefore = GetMachRatio(srcMachV, actualCountByMach, targetedHitTotal)
            Dim dstRatioBefore As Double: dstRatioBefore = GetMachRatio(dstMachV, actualCountByMach, targetedHitTotal)

            ' 実績(予測)集計を更新する
            actualCountByMach(CStr(srcMachV)) = actualCountByMach(CStr(srcMachV)) - rowCnt(srcRow) + rowCnt(partnerRow)
            actualCountByMach(CStr(dstMachV)) = actualCountByMach(CStr(dstMachV)) - rowCnt(partnerRow) + rowCnt(srcRow)

            Dim srcRatioAfter As Double: srcRatioAfter = GetMachRatio(srcMachV, actualCountByMach, targetedHitTotal)
            Dim dstRatioAfter As Double: dstRatioAfter = GetMachRatio(dstMachV, actualCountByMach, targetedHitTotal)

            rowUsed(srcRow) = True
            rowUsed(partnerRow) = True

            outCnt = outCnt + 1
            outArr(outCnt, 1) = rowDan(srcRow) ' 段(移動元・移動先とも同じ)
            outArr(outCnt, 2) = srcMachV
            outArr(outCnt, 3) = rowCol(srcRow)
            outArr(outCnt, 4) = rowItemCode(srcRow)
            outArr(outCnt, 5) = rowItemName(srcRow)
            outArr(outCnt, 6) = rowCnt(srcRow)
            outArr(outCnt, 7) = dstMachV
            outArr(outCnt, 8) = rowCol(partnerRow)
            outArr(outCnt, 9) = rowItemCode(partnerRow)
            outArr(outCnt, 10) = rowItemName(partnerRow)
            outArr(outCnt, 11) = rowCnt(partnerRow)
            outArr(outCnt, 12) = Format(srcRatioBefore, "0.0%") & "→" & Format(srcRatioAfter, "0.0%") & _
                "  /  " & Format(dstRatioBefore, "0.0%") & "→" & Format(dstRatioAfter, "0.0%")
        End If
    Loop

    If outCnt = 0 Then
        MsgBox "条件を満たすロケ変指示が見つかりませんでした。除外設定や目標構成比の入力内容をご確認ください。", vbExclamation
        Exit Sub
    End If

    ' 出力
    On Error Resume Next
    ThisWorkbook.Sheets("ロケ変指示").Delete
    On Error GoTo 0

    Dim wsOut As Worksheet
    Dim wsPanel As Worksheet
    On Error Resume Next
    Set wsPanel = ThisWorkbook.Sheets("操作パネル")
    On Error GoTo 0
    If Not wsPanel Is Nothing Then
        Set wsOut = ThisWorkbook.Sheets.Add(Before:=wsPanel)
    Else
        Set wsOut = Sheets.Add
    End If
    wsOut.Name = "ロケ変指示"

    wsOut.Columns("D:D").NumberFormat = "@" ' 品コードは先頭ゼロ落ち防止
    wsOut.Columns("I:I").NumberFormat = "@"

    wsOut.Range("A1:L1").Merge
    wsOut.Cells(1, 1).Value = "【予測データに基づくロケ変指示(候補" & outCnt & "件/設定" & relocCount & "件・入替相手は同じ段のみ)】"
    wsOut.Cells(1, 1).Font.Bold = True: wsOut.Cells(1, 1).Font.Size = 14
    wsOut.Cells(1, 1).HorizontalAlignment = xlLeft

    wsOut.Range("A2:L2").Merge
    wsOut.Cells(2, 1).Value = "「設定」シートの■機番別目標構成比に近づけるよう、目標超過機番の高頻度ロケーションと目標未達機番のロケーションを、同じ段の中で入れ替える指示です。"
    wsOut.Cells(2, 1).HorizontalAlignment = xlLeft

    wsOut.Range("A4:L4").Value = Array("段", "移動元機番", "移動元列", "移動元品コード", "移動元品名", "移動元予測回数", "移動先機番", "移動先列", "移動先品コード", "移動先品名", "移動先予測回数", "構成比(移動元/移動先:変更前→変更後)")
    wsOut.Range("A5").Resize(outCnt, 12).Value = outArr

    wsOut.Range("A4:L4").Interior.Color = RGB(220, 230, 255)
    wsOut.Range("A4:L4").Font.Bold = True
    wsOut.Columns("A:L").AutoFit

    MsgBox "「ロケ変指示」シートを作成しました。(" & outCnt & "件)", vbInformation
End Sub

Private Function GetMachRatio(ByVal mach As Long, actualCountByMach As Object, ByVal targetedHitTotal As Double) As Double
    Dim c As Double: c = 0
    If actualCountByMach.Exists(CStr(mach)) Then c = actualCountByMach(CStr(mach))
    GetMachRatio = IIf(targetedHitTotal > 0, c / targetedHitTotal, 0)
End Function

Private Function HasUnusedRow(rowsCol As Collection, rowUsed() As Boolean) As Boolean
    Dim v As Variant
    For Each v In rowsCol
        If Not rowUsed(CLng(v)) Then
            HasUnusedRow = True
            Exit Function
        End If
    Next v
    HasUnusedRow = False
End Function

' 指定機番の未使用ロケーションの中で、予測回数が最も多いものを返す(0=無し)
Private Function FindBestSourceRow(rowsCol As Collection, rowUsed() As Boolean, rowCnt() As Double) As Long
    Dim bestRow As Long: bestRow = 0
    Dim bestCnt As Double: bestCnt = -1
    Dim v As Variant
    For Each v In rowsCol
        Dim idx As Long: idx = CLng(v)
        If Not rowUsed(idx) Then
            If rowCnt(idx) > bestCnt Then
                bestCnt = rowCnt(idx)
                bestRow = idx
            End If
        End If
    Next v
    FindBestSourceRow = bestRow
End Function

' 同じ段の未使用ロケーションの中から、超過機番(excludeMach)以外で最も目標構成比を下回っている機番の
' ロケーションを1件返す(0=無し)。目標未設定の機番は対象外にする
Private Function FindBestPartnerRow(rowsCol As Collection, rowUsed() As Boolean, rowMach() As Long, ByVal excludeMach As Long, actualCountByMach As Object, dictTargetRatio As Object, ByVal targetRatioSum As Double, ByVal targetedHitTotal As Double, ByVal maxMachNum As Long) As Long
    Dim bestRow As Long: bestRow = 0
    Dim bestDev As Double: bestDev = 0 ' 0未満(目標未達)のみを対象にするため、初期値0からより小さい値を探す
    Dim v As Variant
    For Each v In rowsCol
        Dim idx As Long: idx = CLng(v)
        If Not rowUsed(idx) Then
            Dim pMach As Long: pMach = rowMach(idx)
            If pMach <> excludeMach And pMach >= 1 And pMach <= maxMachNum Then
                If dictTargetRatio.Exists(CStr(pMach)) Then
                    Dim pActRatio As Double: pActRatio = GetMachRatio(pMach, actualCountByMach, targetedHitTotal)
                    Dim pTgtRatio As Double: pTgtRatio = dictTargetRatio(CStr(pMach)) / targetRatioSum
                    Dim pDev As Double: pDev = pActRatio - pTgtRatio
                    If pDev < bestDev Then
                        bestDev = pDev
                        bestRow = idx
                    End If
                End If
            End If
        End If
    Next v
    FindBestPartnerRow = bestRow
End Function

' 「設定」シートの「ロケ変候補件数」(L8)を読み込む。未入力・1未満なら既定値20を使う
Private Function GetRelocationCandidateCount() As Long
    GetRelocationCandidateCount = 20
    Dim wsSet As Worksheet
    On Error Resume Next
    Set wsSet = ThisWorkbook.Sheets("設定")
    On Error GoTo 0
    If wsSet Is Nothing Then Exit Function
    If IsNumeric(wsSet.Range("L8").Value) Then
        If CLng(wsSet.Range("L8").Value) >= 1 Then GetRelocationCandidateCount = CLng(wsSet.Range("L8").Value)
    End If
End Function

' 「操作パネル」シートにロケ変指示ボタンが無ければ追加する
' (既存のボタン・図形と重ならないよう、一番下にあるものの少し下に配置する)
Sub EnsureRelocationPlanButton()
    Dim wsPanel As Worksheet
    On Error Resume Next
    Set wsPanel = ThisWorkbook.Sheets("操作パネル")
    On Error GoTo 0
    If wsPanel Is Nothing Then Exit Sub

    Dim existing As Shape
    On Error Resume Next
    Set existing = wsPanel.Shapes("ロケ変指示ボタン")
    On Error GoTo 0
    If Not existing Is Nothing Then Exit Sub

    Dim maxBottom As Double: maxBottom = 0
    Dim shp As Shape
    For Each shp In wsPanel.Shapes
        If shp.Top + shp.Height > maxBottom Then maxBottom = shp.Top + shp.Height
    Next shp
    If maxBottom = 0 Then maxBottom = wsPanel.Range("B20").Top

    Dim btn As Button
    Set btn = wsPanel.Buttons.Add(wsPanel.Range("B2").Left, maxBottom + 16, 220, 36)
    btn.Name = "ロケ変指示ボタン"
    btn.OnAction = "CreateRelocationPlan"
    btn.Characters.Text = "ロケ変指示を作成"
    btn.Font.Size = 12
    btn.Font.Bold = True
End Sub
