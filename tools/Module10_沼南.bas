Attribute VB_Name = "Module10_沼南"
Option Explicit

' ----------------------------------------------------
' 予測データに基づくロケ変指示の作成
' 「予測データ」シート(Module8で取込済み)の号機別実績(予測)構成比を、「設定」シートの
' 「■号機別目標構成比」に近づけるためのロケーション変更(入替)指示を作成する。
' 入替の相手先は必ず同じ段(例:46-02-05なら2段目)の中から選ぶ(段をまたぐ入替は行わない)。
' 除外号機・除外ロケーション・除外品コードは「設定」シートの設定に従う(Module3と共通)。
' 対象は「■ABブロック」の範囲内の号機のみ(Cバラ・拡張は対象外)。
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
    Dim ratioSheetName As String, maxSwapRows As Long, abSlotCount As Long
    Dim abBlockFrom() As Long, abBlockTo() As Long, abBlockCount As Long
    Dim dictTargetRatio As Object: Set dictTargetRatio = CreateObject("Scripting.Dictionary")
    Dim catWeight As Double: catWeight = 0.005
    Dim sizeWeight As Double: sizeWeight = 0.01
    Dim weightWeightCoef As Double: weightWeightCoef = 0.01
    Call LoadExclusionSettings(dictExcludedMach, locMach, locDanFrom, locDanTo, locColFrom, locColTo, locCount, dictExcludedItemCode, ratioSheetName, maxSwapRows, abSlotCount, abBlockFrom, abBlockTo, abBlockCount, dictTargetRatio, catWeight, sizeWeight, weightWeightCoef)

    ' 在庫商品マスタ(任意、Module3と共通)。読み込めば、入替候補選定で入替先号機の同カテゴリー品集中度・
    ' サイズ差・重量差をソフトなペナルティとして反映する(未読込なら従来どおりの選定結果になる)
    Dim dictItemCategory As Object: Set dictItemCategory = CreateObject("Scripting.Dictionary")
    Dim dictItemWeightMaster As Object: Set dictItemWeightMaster = CreateObject("Scripting.Dictionary")
    Dim dictItemVolumeMaster As Object: Set dictItemVolumeMaster = CreateObject("Scripting.Dictionary")
    Call LoadItemAttributeMasterFromSheet(dictItemCategory, dictItemWeightMaster, dictItemVolumeMaster)

    Dim relocCount As Long: relocCount = GetRelocationCandidateCount()

    ' 「号機別目標構成比」に数値号機キーが1つも無ければ、近づけるべき目標が無いためロケ変指示を作成できない
    Dim hasAnyTarget As Boolean: hasAnyTarget = False
    Dim chkKey As Variant
    For Each chkKey In dictTargetRatio.Keys
        If IsNumeric(chkKey) Then hasAnyTarget = True: Exit For
    Next chkKey
    If Not hasAnyTarget Then
        MsgBox "「設定」シートの「■号機別目標構成比」に号機別の目標値が入力されていないため、ロケ変指示を作成できません。", vbExclamation
        Exit Sub
    End If

    ' 「予測データ」シートは1行目=取込情報、3行目=見出し(Module8の出力形式)。
    ' 号機・段・列・品名コード・投入回数_予測の列を見出し名から探す(品名列があれば表示にも使う)
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
    Dim dictByMach As Object: Set dictByMach = CreateObject("Scripting.Dictionary") ' 号機→その号機のロケーション行(添字)のCollection
    Dim dictRowMach As Object: Set dictRowMach = CreateObject("Scripting.Dictionary") ' 行番号(文字列)→号機(商品属性ペナルティ計算用)
    Dim dictRowCat As Object: Set dictRowCat = CreateObject("Scripting.Dictionary") ' 行番号→大分類コード等(在庫商品マスタ)
    Dim dictRowWt As Object: Set dictRowWt = CreateObject("Scripting.Dictionary") ' 行番号→重量(kg)
    Dim dictRowVol As Object: Set dictRowVol = CreateObject("Scripting.Dictionary") ' 行番号→体積
    Dim dictMachCatVol As Object: Set dictMachCatVol = CreateObject("Scripting.Dictionary") ' "号機|大分類コード"→その号機内の同カテゴリー品数

    Dim r As Long
    For r = HEADER_ROW + 1 To lastRow
        If IsNumeric(wsData.Cells(r, machColIdx).Value) Then
            Dim mach As Long: mach = CLng(wsData.Cells(r, machColIdx).Value)
            If IsInABBlock(CInt(mach), abBlockFrom, abBlockTo, abBlockCount) Then
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

                            ' 在庫商品マスタが読み込まれていれば、行番号をキーにカテゴリー・重量・体積を引けるようにする
                            Dim rowKey As String: rowKey = CStr(rowN)
                            dictRowMach(rowKey) = mach
                            If itemCodeStr <> "" And IsNumeric(itemCodeStr) Then
                                Dim rowCodeKey As String: rowCodeKey = CStr(CLng(itemCodeStr))
                                If dictItemCategory.Exists(rowCodeKey) Then dictRowCat(rowKey) = dictItemCategory(rowCodeKey)
                                If dictItemWeightMaster.Exists(rowCodeKey) Then dictRowWt(rowKey) = dictItemWeightMaster(rowCodeKey)
                                If dictItemVolumeMaster.Exists(rowCodeKey) Then dictRowVol(rowKey) = dictItemVolumeMaster(rowCodeKey)
                            End If

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

    ' 号機ごとのカテゴリー在庫点数を集計(入替先候補の号機に同カテゴリー品がどれだけ集中しているかの目安に使う)
    Dim rcI As Long
    For rcI = 1 To rowN
        Dim rcKey As String: rcKey = CStr(rcI)
        If dictRowCat.Exists(rcKey) Then
            Dim rcTallyKey As String: rcTallyKey = CStr(rowMach(rcI)) & "|" & dictRowCat(rcKey)
            dictMachCatVol(rcTallyKey) = dictMachCatVol(rcTallyKey) + 1
            If dictRowVol.Exists(rcKey) And dictRowVol(rcKey) > 0 Then
                Dim bIdxR As Long: bIdxR = Int(Log(dictRowVol(rcKey)) / SIZE_SIMILAR_RATIO)
                Dim bKeyR As String: bKeyR = rcTallyKey & "|B" & bIdxR
                dictMachCatVol(bKeyR) = dictMachCatVol(bKeyR) + 1
            Else
                Dim nKeyR As String: nKeyR = rcTallyKey & "|N"
                dictMachCatVol(nKeyR) = dictMachCatVol(nKeyR) + 1
            End If
        End If
    Next rcI

    ' 目標構成比が入力されている号機だけを対象に、合計が100%でなくても相対バランスとして比較できるよう正規化する
    ' (号機別目標構成比の合計と、目標が設定されている号機だけの実績(予測)合計、それぞれで正規化してから比較する)
    Dim targetRatioSum As Double: targetRatioSum = 0
    Dim targetedHitTotal As Double: targetedHitTotal = 0
    Dim trKey As Variant
    For Each trKey In dictTargetRatio.Keys
        If IsNumeric(trKey) Then
            Dim trMach As Long: trMach = CLng(trKey)
            If IsInABBlock(CInt(trMach), abBlockFrom, abBlockTo, abBlockCount) Then
                targetRatioSum = targetRatioSum + dictTargetRatio(trKey)
                If actualCountByMach.Exists(CStr(trMach)) Then targetedHitTotal = targetedHitTotal + actualCountByMach(CStr(trMach))
            End If
        End If
    Next trKey
    If targetRatioSum <= 0 Or targetedHitTotal <= 0 Then
        MsgBox "目標構成比が設定されている号機の実績(予測)データが見つからなかったため、ロケ変指示を作成できません。", vbExclamation
        Exit Sub
    End If

    ' 入替候補の選定(目標構成比を最も上回っている号機の高頻度ロケーションを、同じ段の中から
    ' 最も目標構成比を下回っている号機の未使用ロケーションと入れ替える、を繰り返す)
    Dim outArr() As Variant
    ReDim outArr(1 To relocCount, 1 To 10)
    Dim outCnt As Long: outCnt = 0
    Dim dictGivenUpMach As Object: Set dictGivenUpMach = CreateObject("Scripting.Dictionary") ' これ以上有効な交換相手が見つからない号機

    Dim safetyCounter As Long: safetyCounter = 0
    Dim maxAttempts As Long: maxAttempts = relocCount * 20 + 100 ' 無限ループ防止の安全策

    Do While outCnt < relocCount
        safetyCounter = safetyCounter + 1
        If safetyCounter > maxAttempts Then Exit Do
        ' 候補が多いと探索に時間がかかることがあるため、Excelが「応答なし」に見えないよう
        ' 一定回数ごとに制御をOSに戻す(処理自体は継続する)
        If safetyCounter Mod 10 = 0 Then DoEvents

        ' 現時点で目標構成比を最も上回っていて、かつまだ諦めていない号機を探す
        Dim bestOverMach As String: bestOverMach = ""
        Dim bestOverDev As Double: bestOverDev = 0.0000001 ' 誤差程度の超過は無視する
        For Each trKey In dictTargetRatio.Keys
            If IsNumeric(trKey) Then
                Dim tMach As Long: tMach = CLng(trKey)
                If IsInABBlock(CInt(tMach), abBlockFrom, abBlockTo, abBlockCount) Then
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

        If bestOverMach = "" Then Exit Do ' これ以上、目標構成比を上回っていて動かせる号機が無い

        ' 超過号機の未使用ロケーションを予測回数の多い順に試し、同じ段で目標未達の号機の
        ' 未使用ロケーションが見つかるまで探す(超過側への影響が大きい候補から優先的に解消する)
        Dim srcRow As Long: srcRow = FindBestSourceRow(dictByMach(bestOverMach), rowUsed, rowCnt)
        Dim partnerRow As Long: partnerRow = 0
        Dim innerRetryCounter As Long: innerRetryCounter = 0
        Const INNER_RETRY_CAP As Long = 30 ' 1機番あたりの候補探索を打ち切る上限(大きい機番でも数分待たされないようにする)
        Do While srcRow > 0 And partnerRow = 0 And innerRetryCounter < INNER_RETRY_CAP
            innerRetryCounter = innerRetryCounter + 1
            If innerRetryCounter Mod 20 = 0 Then DoEvents
            partnerRow = FindBestPartnerRow(dictByDan(CStr(rowDan(srcRow))), rowUsed, rowMach, CLng(bestOverMach), actualCountByMach, dictTargetRatio, targetRatioSum, targetedHitTotal, abBlockFrom, abBlockTo, abBlockCount, srcRow, dictRowMach, dictRowCat, dictRowWt, dictRowVol, dictMachCatVol, catWeight, sizeWeight, weightWeightCoef)
            If partnerRow = 0 Then
                Dim tmpUsedMark As Long: tmpUsedMark = srcRow
                rowUsed(tmpUsedMark) = True ' この候補は今回使えないので一時的に使用済み扱いにして次を探す
                srcRow = FindBestSourceRow(dictByMach(bestOverMach), rowUsed, rowCnt)
                rowUsed(tmpUsedMark) = False ' 出力しない場合は使用済みフラグを戻す(以後の反復で再検討できるようにする)
            End If
        Loop

        If partnerRow = 0 Then
            dictGivenUpMach(bestOverMach) = True ' この号機からはこれ以上有効な交換相手が見つからない
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
            outArr(outCnt, 1) = rowItemName(srcRow)
            outArr(outCnt, 2) = rowItemCode(srcRow)
            outArr(outCnt, 3) = FormatLocationStr(srcMachV, rowDan(srcRow), rowCol(srcRow))
            outArr(outCnt, 4) = rowCnt(srcRow)
            outArr(outCnt, 5) = "⇔"
            outArr(outCnt, 6) = rowItemName(partnerRow)
            outArr(outCnt, 7) = rowItemCode(partnerRow)
            outArr(outCnt, 8) = FormatLocationStr(dstMachV, rowDan(partnerRow), rowCol(partnerRow))
            outArr(outCnt, 9) = rowCnt(partnerRow)
            outArr(outCnt, 10) = Format(srcRatioBefore, "0.0%") & "→" & Format(srcRatioAfter, "0.0%") & _
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

    ' 品コード・ロケーションは先頭ゼロ落ち・日付誤変換防止のため文字列表示にする(AB編成動線最適化と同じ扱い)
    wsOut.Columns("B:B").NumberFormat = "@"
    wsOut.Columns("C:C").NumberFormat = "@"
    wsOut.Columns("G:G").NumberFormat = "@"
    wsOut.Columns("H:H").NumberFormat = "@"

    wsOut.Range("A1:J1").Merge
    wsOut.Cells(1, 1).Value = "【予測データに基づくロケ変指示(候補" & outCnt & "件/設定" & relocCount & "件・入替相手は同じ段のみ)】"
    wsOut.Cells(1, 1).Font.Bold = True: wsOut.Cells(1, 1).Font.Size = 14
    wsOut.Cells(1, 1).HorizontalAlignment = xlLeft

    wsOut.Range("A2:J2").Merge
    wsOut.Cells(2, 1).Value = "「設定」シートの■号機別目標構成比に近づけるよう、目標超過号機の高頻度ロケーションと目標未達号機のロケーションを、同じ段の中で入れ替える指示です。ロケーションは「号機-段-列」の表記です(AB編成動線最適化と同じ)。"
    wsOut.Cells(2, 1).HorizontalAlignment = xlLeft

    wsOut.Range("A4:J4").Value = Array("【移動元品】(交換品コード)", "移動元品コード", "移動元ロケーション", "移動元予測回数", "交換方向", "【移動先品】(交換対象品)", "移動先品コード", "移動先ロケーション", "移動先予測回数", "構成比(移動元/移動先:変更前→変更後)")
    wsOut.Range("A5").Resize(outCnt, 10).Value = outArr

    wsOut.Range("A4:J4").Interior.Color = RGB(220, 230, 255)
    wsOut.Range("A4:J4").Font.Bold = True
    wsOut.Columns("A:J").AutoFit

    MsgBox "「ロケ変指示」シートを作成しました。(" & outCnt & "件)", vbInformation
End Sub

' AB編成動線最適化と同じ「号機-段-列」形式でロケーションを表記する(例:46-02-05)
Private Function FormatLocationStr(ByVal mach As Long, ByVal dan As Long, ByVal col As Long) As String
    FormatLocationStr = mach & "-" & Format(dan, "00") & "-" & Format(col, "00")
End Function

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

' 指定号機の未使用ロケーションの中で、予測回数が最も多いものを返す(0=無し)
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

' 同じ段の未使用ロケーションの中から、超過号機(excludeMach)以外で最も目標構成比を下回っている号機の
' ロケーションを1件返す(0=無し)。目標未設定・ABブロック外の号機は対象外にする
Private Function FindBestPartnerRow(rowsCol As Collection, rowUsed() As Boolean, rowMach() As Long, ByVal excludeMach As Long, actualCountByMach As Object, dictTargetRatio As Object, ByVal targetRatioSum As Double, ByVal targetedHitTotal As Double, abBlockFrom() As Long, abBlockTo() As Long, ByVal abBlockCount As Long, ByVal srcRow As Long, dictRowMach As Object, dictRowCat As Object, dictRowWt As Object, dictRowVol As Object, dictMachCatVol As Object, ByVal catWeight As Double, ByVal sizeWeight As Double, ByVal weightWeightCoef As Double) As Long
    Dim bestRow As Long: bestRow = 0
    Dim bestScore As Double: bestScore = 0 ' 目標未達(pDev<0)の候補の中で、カテゴリー・サイズも加味した最良のものを選ぶ

    ' moverアイテム(srcRow)の属性は候補走査の前に1回だけ解決しておく(候補ごとに辞書引きし直すと、
    ' 候補数の多いロケ変指示では無駄な処理が積み重なって動作が重くなるため。Module3のResolveMoverAttrを共用)
    Dim moverHasCat As Boolean, moverCat As String
    Dim moverHasWt As Boolean, moverWt As Double
    Dim moverHasVol As Boolean, moverVol As Double
    Call ResolveMoverAttr(CStr(srcRow), dictRowCat, dictRowWt, dictRowVol, moverHasCat, moverCat, moverHasWt, moverWt, moverHasVol, moverVol)

    Dim v As Variant
    For Each v In rowsCol
        Dim idx As Long: idx = CLng(v)
        If Not rowUsed(idx) Then
            Dim pMach As Long: pMach = rowMach(idx)
            If pMach <> excludeMach And IsInABBlock(CInt(pMach), abBlockFrom, abBlockTo, abBlockCount) Then
                If dictTargetRatio.Exists(CStr(pMach)) Then
                    Dim pActRatio As Double: pActRatio = GetMachRatio(pMach, actualCountByMach, targetedHitTotal)
                    Dim pTgtRatio As Double: pTgtRatio = dictTargetRatio(CStr(pMach)) / targetRatioSum
                    Dim pDev As Double: pDev = pActRatio - pTgtRatio
                    If pDev < 0 Then
                        Dim pScore As Double
                        pScore = pDev + ComputeAttrPenalty(CStr(idx), dictRowMach, dictRowVol, dictRowWt, dictMachCatVol, moverHasCat, moverCat, moverHasWt, moverWt, moverHasVol, moverVol, catWeight, sizeWeight, weightWeightCoef)
                        If bestRow = 0 Or pScore < bestScore Then
                            bestScore = pScore
                            bestRow = idx
                        End If
                    End If
                End If
            End If
        End If
    Next v
    FindBestPartnerRow = bestRow
End Function

' 「設定」シートの「ロケ変候補件数」(L7)を読み込む。未入力・1未満なら既定値20を使う
Private Function GetRelocationCandidateCount() As Long
    GetRelocationCandidateCount = 20
    Dim wsSet As Worksheet
    On Error Resume Next
    Set wsSet = ThisWorkbook.Sheets("設定")
    On Error GoTo 0
    If wsSet Is Nothing Then Exit Function
    If IsNumeric(wsSet.Range("L7").Value) Then
        If CLng(wsSet.Range("L7").Value) >= 1 Then GetRelocationCandidateCount = CLng(wsSet.Range("L7").Value)
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
    If existing Is Nothing Then
        Dim btn As Button
        Set btn = wsPanel.Buttons.Add(wsPanel.Range("B20").Left, wsPanel.Range("B20").Top, 220, 36)
        btn.Name = "ロケ変指示ボタン"
        btn.OnAction = "CreateRelocationPlan"
        btn.Characters.Text = "ロケ変指示を作成"
        btn.Font.Size = 12
        btn.Font.Bold = True
    End If

    ' ボタンが下に伸び続けないよう、2列に並び替える(Module3の共通処理)
    Call LayoutPanelButtons
End Sub

' ----------------------------------------------------
' AB(1～46号機)・Cバラ(51～68号機)・X拡張(70号機以上)の3ゾーン間で、
' 出荷回数の順位に応じたゾーン間入替候補を作成する。
' 出荷回数は「日別ロケーション実績」の月曜列(実績)を優先し、無ければ
' 「予測データ」の「投入回数_曜日平均」で代用する。
' 各ゾーンの容量は実際にある間口数(該当ロケーション数)をそのまま使い、
' 出荷回数の多い順にAB→C→Xの優先度で「あるべきゾーン」を決める。
' 段(棚の高さ)はAB/C/Xで形状が異なる別ゾーンのため考慮せず、順位のズレ
' だけでペアを作る。除外設定(除外号機・除外ロケーション・除外品コード)は
' ロケ変指示と共通のものを使う。
' ----------------------------------------------------
Sub CreateZoneRebalancePlan()
    Call EnsureZoneRebalanceButton

    Dim wsPred As Worksheet
    On Error Resume Next
    Set wsPred = ThisWorkbook.Sheets("予測データ")
    On Error GoTo 0
    If wsPred Is Nothing Then
        MsgBox "「予測データ」シートが見つかりません。先に「予測データを取り込む」を実行してください。", vbExclamation
        Exit Sub
    End If

    Call EnsureExclusionSettingsSheet

    Dim dictExcludedMach As Object: Set dictExcludedMach = CreateObject("Scripting.Dictionary")
    Dim locMach() As Long, locDanFrom() As Long, locDanTo() As Long, locColFrom() As Long, locColTo() As Long
    Dim locCount As Long
    Dim dictExcludedItemCode As Object: Set dictExcludedItemCode = CreateObject("Scripting.Dictionary")
    Dim ratioSheetName As String, maxSwapRows As Long, abSlotCount As Long
    Dim abBlockFrom() As Long, abBlockTo() As Long, abBlockCount As Long
    Dim dictTargetRatio As Object: Set dictTargetRatio = CreateObject("Scripting.Dictionary")
    Dim catWeight As Double: catWeight = 0.005
    Dim sizeWeight As Double: sizeWeight = 0.01
    Dim weightWeightCoef As Double: weightWeightCoef = 0.01
    Call LoadExclusionSettings(dictExcludedMach, locMach, locDanFrom, locDanTo, locColFrom, locColTo, locCount, dictExcludedItemCode, ratioSheetName, maxSwapRows, abSlotCount, abBlockFrom, abBlockTo, abBlockCount, dictTargetRatio, catWeight, sizeWeight, weightWeightCoef)

    Const ZONE_AB_MAX As Long = 46
    Const ZONE_C_MIN As Long = 51
    Const ZONE_C_MAX As Long = 68
    Const ZONE_X_MIN As Long = 70
    Const TARGET_AB As Double = 0.92
    Const TARGET_C As Double = 0.07
    Const TARGET_X As Double = 0.01

    Const HEADER_ROW As Long = 3
    Dim lastRow As Long: lastRow = wsPred.Cells(wsPred.Rows.Count, 1).End(xlUp).Row
    Dim lastCol As Long: lastCol = wsPred.Cells(HEADER_ROW, wsPred.Columns.Count).End(xlToLeft).Column
    If lastRow <= HEADER_ROW Then
        MsgBox "「予測データ」シートにデータ行がありません。", vbExclamation
        Exit Sub
    End If

    Dim headerArr As Variant
    headerArr = wsPred.Range(wsPred.Cells(HEADER_ROW, 1), wsPred.Cells(HEADER_ROW, lastCol)).Value
    Dim machColIdx As Long: machColIdx = -1
    Dim danColIdx As Long: danColIdx = -1
    Dim colColIdx As Long: colColIdx = -1
    Dim itemCodeColIdx As Long: itemCodeColIdx = -1
    Dim itemNameColIdx As Long: itemNameColIdx = -1
    Dim wdAvgColIdx As Long: wdAvgColIdx = -1
    Dim hc As Long
    For hc = 1 To lastCol
        Dim hName As String: hName = Trim(CStr(headerArr(1, hc)))
        If hName = "号機" Then machColIdx = hc
        If hName = "段" Then danColIdx = hc
        If hName = "列" Then colColIdx = hc
        If hName = "品名コード" Then itemCodeColIdx = hc
        If hName = "品名" Then itemNameColIdx = hc
        If hName = "投入回数_曜日平均" Then wdAvgColIdx = hc
    Next hc
    If machColIdx = -1 Or danColIdx = -1 Or colColIdx = -1 Or itemCodeColIdx = -1 Then
        MsgBox "「予測データ」シートに「号機」「段」「列」「品名コード」のいずれかの列が見つかりません。", vbExclamation
        Exit Sub
    End If

    ' 「日別ロケーション実績」の月曜列があれば、ロケーションごとの月曜実績を読み込む
    Dim dictMondayActual As Object: Set dictMondayActual = CreateObject("Scripting.Dictionary")
    Dim hasMondayCol As Boolean: hasMondayCol = False
    Dim wsHist As Worksheet
    On Error Resume Next
    Set wsHist = ThisWorkbook.Sheets("日別ロケーション実績")
    On Error GoTo 0
    If Not wsHist Is Nothing Then
        Const HIST_DATE_COL_FIRST As Long = 7
        Const HIST_DATE_COL_LAST As Long = 16
        Dim mondayCol As Long: mondayCol = -1
        Dim hdc As Long
        For hdc = HIST_DATE_COL_FIRST To HIST_DATE_COL_LAST
            If InStr(CStr(wsHist.Cells(1, hdc).Value), "(月)") > 0 Then mondayCol = hdc
        Next hdc
        If mondayCol > 0 Then
            hasMondayCol = True
            Dim histLastRow As Long: histLastRow = wsHist.Cells(wsHist.Rows.Count, 1).End(xlUp).Row
            If histLastRow >= 2 Then
                Dim histArr As Variant
                histArr = wsHist.Range(wsHist.Cells(2, 4), wsHist.Cells(histLastRow, mondayCol)).Value
                Dim mondayRelCol As Long: mondayRelCol = mondayCol - 4 + 1
                Dim hr As Long
                For hr = 1 To UBound(histArr, 1)
                    Dim hLocKey As Variant: hLocKey = histArr(hr, 1)
                    If IsNumeric(hLocKey) Then
                        Dim hVal As Double: hVal = 0
                        Dim hMondayVal As Variant: hMondayVal = histArr(hr, mondayRelCol)
                        If IsNumeric(hMondayVal) Then hVal = CDbl(hMondayVal)
                        dictMondayActual(CStr(CLng(hLocKey))) = hVal
                    End If
                Next hr
            End If
        End If
    End If

    ' 対象ロケーション(AB・Cバラ・X拡張のいずれか、除外条件を除く)を配列にまとめる
    Dim dataArr As Variant
    dataArr = wsPred.Range(wsPred.Cells(HEADER_ROW + 1, 1), wsPred.Cells(lastRow, lastCol)).Value
    Dim n As Long: n = UBound(dataArr, 1)

    Dim rMach() As Long, rDan() As Long, rCol() As Long
    Dim rCode() As String, rName() As String, rCnt() As Double, rZone() As String
    ReDim rMach(1 To n)
    ReDim rDan(1 To n)
    ReDim rCol(1 To n)
    ReDim rCode(1 To n)
    ReDim rName(1 To n)
    ReDim rCnt(1 To n)
    ReDim rZone(1 To n)
    Dim m As Long: m = 0

    Dim i As Long
    For i = 1 To n
        If IsNumeric(dataArr(i, machColIdx)) And IsNumeric(dataArr(i, danColIdx)) And IsNumeric(dataArr(i, colColIdx)) Then
            Dim mach As Long: mach = CLng(dataArr(i, machColIdx))
            Dim zone As String: zone = ""
            If mach >= 1 And mach <= ZONE_AB_MAX Then
                zone = "AB"
            ElseIf mach >= ZONE_C_MIN And mach <= ZONE_C_MAX Then
                zone = "C"
            ElseIf mach >= ZONE_X_MIN Then
                zone = "X"
            End If
            If zone <> "" Then
                If Not IsExcludedSlot3(dictExcludedMach, CInt(mach)) Then
                    Dim dan As Long: dan = CLng(dataArr(i, danColIdx))
                    Dim colv As Long: colv = CLng(dataArr(i, colColIdx))
                    If Not IsExcludedLocation(CInt(mach), CInt(dan), CInt(colv), locMach, locDanFrom, locDanTo, locColFrom, locColTo, locCount) Then
                        Dim itemCodeStr As String: itemCodeStr = Trim(CStr(dataArr(i, itemCodeColIdx)))
                        Dim isItemExcluded As Boolean: isItemExcluded = False
                        If itemCodeStr <> "" And dictExcludedItemCode.Count > 0 Then
                            isItemExcluded = dictExcludedItemCode.Exists(itemCodeStr)
                            If Not isItemExcluded And IsNumeric(itemCodeStr) Then isItemExcluded = dictExcludedItemCode.Exists(CStr(CLng(itemCodeStr)))
                        End If
                        If Not isItemExcluded Then
                            Dim cntVal As Double: cntVal = 0
                            Dim locKeyStr As String: locKeyStr = CStr(mach * 10000& + dan * 100& + colv)
                            If hasMondayCol And dictMondayActual.Exists(locKeyStr) Then
                                cntVal = dictMondayActual(locKeyStr)
                            ElseIf wdAvgColIdx > 0 Then
                                cntVal = Val(dataArr(i, wdAvgColIdx))
                            End If

                            m = m + 1
                            rMach(m) = mach
                            rDan(m) = dan
                            rCol(m) = colv
                            rCode(m) = itemCodeStr
                            rName(m) = IIf(itemNameColIdx > 0, Trim(CStr(dataArr(i, itemNameColIdx))), "")
                            rCnt(m) = cntVal
                            rZone(m) = zone
                        End If
                    End If
                End If
            End If
        End If
    Next i

    If m = 0 Then
        MsgBox "AB・Cバラ・X拡張のいずれかの範囲に該当するロケーションが見つかりませんでした。", vbExclamation
        Exit Sub
    End If

    ' 出荷回数の多い順に並べる(QuickSort、降順)
    Dim idx() As Long: ReDim idx(1 To m)
    Dim ii As Long
    For ii = 1 To m
        idx(ii) = ii
    Next ii
    Call QuickSortIdxByCntDesc(idx, rCnt, 1, m)

    ' 各ゾーンの実際の間口数(該当ロケーション数)を容量として、出荷回数の多い順に
    ' AB→C→Xの優先度で「あるべきゾーン」を割り当てる
    Dim capAB As Long: capAB = 0
    Dim capC As Long: capC = 0
    Dim capX As Long: capX = 0
    For ii = 1 To m
        Select Case rZone(ii)
            Case "AB": capAB = capAB + 1
            Case "C": capC = capC + 1
            Case "X": capX = capX + 1
        End Select
    Next ii

    Dim idealZone() As String: ReDim idealZone(1 To m)
    Dim rank As Long
    For rank = 1 To m
        Dim origIdx As Long: origIdx = idx(rank)
        If rank <= capAB Then
            idealZone(origIdx) = "AB"
        ElseIf rank <= capAB + capC Then
            idealZone(origIdx) = "C"
        Else
            idealZone(origIdx) = "X"
        End If
    Next rank

    ' 現在のゾーンと、あるべきゾーンがズレている商品を、ズレの向きごとに集める
    ' (順位順=出荷回数の多い順に集まるので、影響の大きいズレから優先的にペアになる)
    Dim colC2AB As Collection: Set colC2AB = New Collection
    Dim colAB2C As Collection: Set colAB2C = New Collection
    Dim colX2AB As Collection: Set colX2AB = New Collection
    Dim colAB2X As Collection: Set colAB2X = New Collection
    Dim colX2C As Collection: Set colX2C = New Collection
    Dim colC2X As Collection: Set colC2X = New Collection
    For rank = 1 To m
        Dim oi As Long: oi = idx(rank)
        If idealZone(oi) <> rZone(oi) Then
            Dim dirKey As String: dirKey = rZone(oi) & ">" & idealZone(oi)
            Select Case dirKey
                Case "C>AB": colC2AB.Add oi
                Case "AB>C": colAB2C.Add oi
                Case "X>AB": colX2AB.Add oi
                Case "AB>X": colAB2X.Add oi
                Case "X>C": colX2C.Add oi
                Case "C>X": colC2X.Add oi
            End Select
        End If
    Next rank

    Dim outRows As Collection: Set outRows = New Collection
    Call AppendZonePairs(outRows, colC2AB, colAB2C, rMach, rDan, rCol, rCode, rName, rCnt, rZone)
    Call AppendZonePairs(outRows, colX2AB, colAB2X, rMach, rDan, rCol, rCode, rName, rCnt, rZone)
    Call AppendZonePairs(outRows, colX2C, colC2X, rMach, rDan, rCol, rCode, rName, rCnt, rZone)

    If outRows.Count = 0 Then
        MsgBox "現状ですでにゾーン間の入替候補はありませんでした(出荷回数順の理想配置と一致しています)。", vbInformation
        Exit Sub
    End If

    On Error Resume Next
    ThisWorkbook.Sheets("ゾーン間入替候補").Delete
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
    wsOut.Name = "ゾーン間入替候補"

    ' 現状・目標・施策後見込みのゾーン別構成比をまとめて先頭に表示する
    Dim sumAB As Double, sumC As Double, sumX As Double, sumAll As Double
    For ii = 1 To m
        Select Case rZone(ii)
            Case "AB": sumAB = sumAB + rCnt(ii)
            Case "C": sumC = sumC + rCnt(ii)
            Case "X": sumX = sumX + rCnt(ii)
        End Select
    Next ii
    sumAll = sumAB + sumC + sumX

    Dim idealAB As Double, idealC As Double, idealX As Double
    For ii = 1 To m
        Select Case idealZone(ii)
            Case "AB": idealAB = idealAB + rCnt(ii)
            Case "C": idealC = idealC + rCnt(ii)
            Case "X": idealX = idealX + rCnt(ii)
        End Select
    Next ii

    wsOut.Range("A1").Value = "ゾーン"
    wsOut.Range("B1").Value = "現状構成比"
    wsOut.Range("C1").Value = "目標構成比"
    wsOut.Range("D1").Value = "施策後見込み構成比"
    wsOut.Range("A2").Value = "AB(1～46号機)"
    wsOut.Range("A3").Value = "Cバラ(51～68号機)"
    wsOut.Range("A4").Value = "X拡張(70号機以上)"
    If sumAll > 0 Then
        wsOut.Range("B2").Value = sumAB / sumAll
        wsOut.Range("B3").Value = sumC / sumAll
        wsOut.Range("B4").Value = sumX / sumAll
        wsOut.Range("D2").Value = idealAB / sumAll
        wsOut.Range("D3").Value = idealC / sumAll
        wsOut.Range("D4").Value = idealX / sumAll
    End If
    wsOut.Range("C2").Value = TARGET_AB
    wsOut.Range("C3").Value = TARGET_C
    wsOut.Range("C4").Value = TARGET_X
    wsOut.Range("B2:D4").NumberFormat = "0.0%"
    wsOut.Range("A1:D1").Font.Bold = True
    wsOut.Range("A1:D4").Columns.AutoFit
    If Not hasMondayCol Then
        wsOut.Range("A5").Value = "※月曜実績が無いため、予測データの曜日平均で代用しています"
    End If

    Const TABLE_HEADER_ROW As Long = 7
    wsOut.Range(wsOut.Cells(TABLE_HEADER_ROW, 1), wsOut.Cells(TABLE_HEADER_ROW, 10)).Value = Array("品名(移動元)", "品コード(移動元)", "ロケーション(移動元)", "出荷回数(移動元)", "⇒", "品名(移動先)", "品コード(移動先)", "ロケーション(移動先)", "出荷回数(移動先)", "ゾーン変化")
    wsOut.Range(wsOut.Cells(TABLE_HEADER_ROW, 1), wsOut.Cells(TABLE_HEADER_ROW, 10)).Font.Bold = True
    wsOut.Range(wsOut.Cells(TABLE_HEADER_ROW, 1), wsOut.Cells(TABLE_HEADER_ROW, 10)).Interior.Color = RGB(220, 230, 255)

    Dim outR As Long: outR = TABLE_HEADER_ROW
    Dim rv As Variant
    For Each rv In outRows
        outR = outR + 1
        wsOut.Cells(outR, 1).Value = rv(5)
        wsOut.Cells(outR, 2).Value = rv(4)
        wsOut.Cells(outR, 3).Value = rv(1) & "-" & Format(rv(2), "00") & "-" & Format(rv(3), "00")
        wsOut.Cells(outR, 4).Value = rv(6)
        wsOut.Cells(outR, 5).Value = "⇒"
        wsOut.Cells(outR, 6).Value = rv(12)
        wsOut.Cells(outR, 7).Value = rv(11)
        wsOut.Cells(outR, 8).Value = rv(8) & "-" & Format(rv(9), "00") & "-" & Format(rv(10), "00")
        wsOut.Cells(outR, 9).Value = rv(13)
        wsOut.Cells(outR, 10).Value = rv(0) & "→" & rv(7) & " / " & rv(7) & "→" & rv(0)
    Next rv

    wsOut.Range(wsOut.Cells(TABLE_HEADER_ROW, 1), wsOut.Cells(outR, 10)).Columns.AutoFit
    wsOut.Range(wsOut.Cells(TABLE_HEADER_ROW, 1), wsOut.Cells(TABLE_HEADER_ROW, 10)).AutoFilter

    MsgBox "「ゾーン間入替候補」シートを作成しました。(" & outRows.Count & "件)", vbInformation
End Sub

' idx()を、rCnt()の値が大きい順(降順)に並べ替える(QuickSort)
Private Sub QuickSortIdxByCntDesc(idx() As Long, rCnt() As Double, ByVal lo As Long, ByVal hi As Long)
    If lo >= hi Then Exit Sub
    Dim pivot As Double: pivot = rCnt(idx((lo + hi) \ 2))
    Dim i As Long: i = lo
    Dim j As Long: j = hi
    Do While i <= j
        Do While rCnt(idx(i)) > pivot
            i = i + 1
        Loop
        Do While rCnt(idx(j)) < pivot
            j = j - 1
        Loop
        If i <= j Then
            Dim tmp As Long: tmp = idx(i)
            idx(i) = idx(j)
            idx(j) = tmp
            i = i + 1
            j = j - 1
        End If
    Loop
    If lo < j Then Call QuickSortIdxByCntDesc(idx, rCnt, lo, j)
    If i < hi Then Call QuickSortIdxByCntDesc(idx, rCnt, i, hi)
End Sub

' 「本来こちらへ移りたい」候補(colInto)と「本来あちらへ移りたい」候補(colOutOf)を
' 先頭(出荷回数が多い方)から順にペアにしてswapリスト(outRows)に追加する
' (件数が多い方の余りは、対になる相手が無いため今回は対象外とする)
Private Sub AppendZonePairs(outRows As Collection, colInto As Collection, colOutOf As Collection, rMach() As Long, rDan() As Long, rCol() As Long, rCode() As String, rName() As String, rCnt() As Double, rZone() As String)
    Dim n As Long: n = colInto.Count
    If colOutOf.Count < n Then n = colOutOf.Count
    Dim k As Long
    For k = 1 To n
        Dim iA As Long: iA = colInto(k)
        Dim iB As Long: iB = colOutOf(k)
        outRows.Add Array(rZone(iA), rMach(iA), rDan(iA), rCol(iA), rCode(iA), rName(iA), rCnt(iA), rZone(iB), rMach(iB), rDan(iB), rCol(iB), rCode(iB), rName(iB), rCnt(iB))
    Next k
End Sub

' 「操作パネル」シートにゾーン間入替候補作成ボタンが無ければ追加する
Sub EnsureZoneRebalanceButton()
    Dim wsPanel As Worksheet
    On Error Resume Next
    Set wsPanel = ThisWorkbook.Sheets("操作パネル")
    On Error GoTo 0
    If wsPanel Is Nothing Then Exit Sub

    Dim existing As Shape
    On Error Resume Next
    Set existing = wsPanel.Shapes("ゾーン間入替候補ボタン")
    On Error GoTo 0
    If existing Is Nothing Then
        Dim btn As Button
        Set btn = wsPanel.Buttons.Add(wsPanel.Range("B24").Left, wsPanel.Range("B24").Top, 220, 36)
        btn.Name = "ゾーン間入替候補ボタン"
        btn.OnAction = "CreateZoneRebalancePlan"
        btn.Characters.Text = "ゾーン間入替候補作成"
        btn.Font.Size = 12
        btn.Font.Bold = True
    End If

    Call LayoutPanelButtons
End Sub
