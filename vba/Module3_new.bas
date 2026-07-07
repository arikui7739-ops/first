Attribute VB_Name = "Module3"
Sub OptimizeABFormationFlow()
    Dim fd As Office.FileDialog
    Dim filePath As String
    Dim fileNo As Integer, textLine As String

    Dim dictItemLoc As Object, dictItemHit As Object, dictItemMach As Object, dictItemZone As Object
    Set dictItemLoc = CreateObject("Scripting.Dictionary")
    Set dictItemHit = CreateObject("Scripting.Dictionary")
    Set dictItemMach = CreateObject("Scripting.Dictionary")
    Set dictItemZone = CreateObject("Scripting.Dictionary")

    Dim dictPairs As Object: Set dictPairs = CreateObject("Scripting.Dictionary") ' 同一編成・同一ゾーンの共起回数
    Dim dictCrossFace As Object: Set dictCrossFace = CreateObject("Scripting.Dictionary") ' そのペアが対面(異なる号機)かどうか
    Dim currentFormationItems As Object: Set currentFormationItems = CreateObject("Scripting.Dictionary")
    Dim orderCountInFormation As Long: orderCountInFormation = 0
    Dim dictAllHit As Object: Set dictAllHit = CreateObject("Scripting.Dictionary") ' AB占有率スコア用：全ゾーンのヒット数

    ' ヒートマップ表示専用（スワップ対象外の1～4号機も含めた全AB間口のゾーン・対面情報）
    ' ※スワップ候補探索・AB号機使用比率スコアからは1～4号機を除外する方針は変えず、ヒートマップ表示だけ実態を反映する
    Dim dictItemZoneAll As Object: Set dictItemZoneAll = CreateObject("Scripting.Dictionary")
    Dim dictItemMachAll As Object: Set dictItemMachAll = CreateObject("Scripting.Dictionary")
    Dim dictPairsAll As Object: Set dictPairsAll = CreateObject("Scripting.Dictionary")
    Dim dictCrossFaceAll As Object: Set dictCrossFaceAll = CreateObject("Scripting.Dictionary")
    Dim currentFormationItemsAll As Object: Set currentFormationItemsAll = CreateObject("Scripting.Dictionary")

    ' 0. CFシートから ロケーション→品名・商品コード の対応表を先に作っておく
    Dim dictLocName As Object: Set dictLocName = CreateObject("Scripting.Dictionary")
    Dim dictLocCode As Object: Set dictLocCode = CreateObject("Scripting.Dictionary")
    Dim wsCF As Worksheet
    On Error Resume Next
    Set wsCF = ActiveWorkbook.Sheets("CF")
    On Error GoTo 0
    If Not wsCF Is Nothing Then
        Dim lastCF As Long: lastCF = wsCF.Cells(wsCF.Rows.Count, "B").End(xlUp).row
        Dim cf As Long
        For cf = 2 To lastCF
            Dim locCode As String: locCode = Trim(CStr(wsCF.Cells(cf, 2).Value)) ' B列:ロケーション(号機*10000+段*100+列)
            If locCode <> "" And Not dictLocName.Exists(locCode) Then
                dictLocName.Add locCode, CStr(wsCF.Cells(cf, 9).Value) ' I列:品名
                Dim codeVal As Variant
                If IsNumeric(wsCF.Cells(cf, 8).Value) Then
                    codeVal = CLng(wsCF.Cells(cf, 8).Value)
                Else
                    codeVal = wsCF.Cells(cf, 8).Value
                End If
                dictLocCode.Add locCode, codeVal ' H列:品名コード
            End If
        Next cf
    End If

    ' 1. ファイル選択（複数選択可・全ファイル形式）
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "ピッキング実績ファイルを選択（複数選択可）"
        .Filters.Clear
        .Filters.Add "すべてのファイル", "*.*"
        .AllowMultiSelect = True
        If .Show = False Then Exit Sub
    End With

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    ' 2. データの読み込み（H行6件を1編成として区切り、AB(1～46号機、中量棚除く)を対象に集計）
    ' ※編成はファイルをまたがない前提のため、ファイルが変わるたびに前ファイルの端数編成を締めてリセットする
    Dim fIdx As Long
    Dim latestFileDate As Date: latestFileDate = DateSerial(1900, 1, 1) ' ファイル更新日時（B行から日付が読めない場合のフォールバック）
    Dim latestBDate As Date: latestBDate = DateSerial(1900, 1, 1) ' B行(先頭"B"+8桁日付)から読み取った実績日
    For fIdx = 1 To fd.SelectedItems.Count
        If fIdx > 1 Then
            Call RecordZonePairs(currentFormationItems, dictPairs, dictCrossFace, dictItemZone, dictItemMach)
            currentFormationItems.RemoveAll
            Call RecordZonePairs(currentFormationItemsAll, dictPairsAll, dictCrossFaceAll, dictItemZoneAll, dictItemMachAll)
            currentFormationItemsAll.RemoveAll
            orderCountInFormation = 0
        End If

        filePath = fd.SelectedItems(fIdx)
        Dim thisFileDate As Date: thisFileDate = FileDateTime(filePath)
        If thisFileDate > latestFileDate Then latestFileDate = thisFileDate
        fileNo = FreeFile
        Dim skipMode As Boolean: skipMode = False ' H99999(各ロケの在庫数サマリー行)配下は読み飛ばす
        Open filePath For Input As #fileNo
        Do While Not EOF(fileNo)
            Line Input #fileNo, textLine
            If Left(textLine, 1) = "B" And Len(textLine) >= 9 Then
                ' B行の2～9文字目(8桁)が実績日(YYYYMMDD)
                Dim bDateStr As String: bDateStr = Mid(textLine, 2, 8)
                If IsNumeric(bDateStr) Then
                    Dim bDate As Date
                    On Error Resume Next
                    bDate = DateSerial(CInt(Left(bDateStr, 4)), CInt(Mid(bDateStr, 5, 2)), CInt(Mid(bDateStr, 7, 2)))
                    On Error GoTo 0
                    If bDate > latestBDate Then latestBDate = bDate
                End If
            ElseIf Left(textLine, 1) = "H" Then
                If Mid(textLine, 2, 5) = "99999" Then
                    ' 在庫数サマリー行。直前の編成を確定し、以降のE行(在庫数)はオーダーとして扱わない
                    Call RecordZonePairs(currentFormationItems, dictPairs, dictCrossFace, dictItemZone, dictItemMach)
                    currentFormationItems.RemoveAll
                    Call RecordZonePairs(currentFormationItemsAll, dictPairsAll, dictCrossFaceAll, dictItemZoneAll, dictItemMachAll)
                    currentFormationItemsAll.RemoveAll
                    skipMode = True
                Else
                    skipMode = False
                    orderCountInFormation = orderCountInFormation + 1
                    If orderCountInFormation > 6 Then
                        Call RecordZonePairs(currentFormationItems, dictPairs, dictCrossFace, dictItemZone, dictItemMach)
                        currentFormationItems.RemoveAll
                        Call RecordZonePairs(currentFormationItemsAll, dictPairsAll, dictCrossFaceAll, dictItemZoneAll, dictItemMachAll)
                        currentFormationItemsAll.RemoveAll
                        orderCountInFormation = 1
                    End If
                End If
            ElseIf Left(textLine, 1) = "E" And Len(textLine) >= 10 And Not skipMode Then
                Dim slotStart As Long
                For slotStart = 2 To Len(textLine) - 8 Step 13
                    Dim rec As String: rec = Mid(textLine, slotStart, 9)
                    If Trim(rec) <> "" And Len(Trim(rec)) = 9 And IsNumeric(rec) Then
                        Dim mach As Integer, dan As Integer, retsu As Integer
                        mach = Val(Mid(rec, 1, 2))
                        dan = Val(Mid(rec, 3, 2))
                        retsu = Val(Mid(rec, 5, 2))

                        ' AB占有率スコア用：全ゾーン(号機の範囲を問わず)のヒット数を集計（中量棚のみ除く）
                        ' ※未使用スロット（号機00等のパディング）が実データに混ざっていると全体回数が水増しされ、
                        '   理論比率・実績比率とも本来の値からズレるため、mach=0（実在しない号機）は除外する
                        If mach > 0 And Not ((mach = 1 Or mach = 2) And retsu >= 6 And retsu <= 14) Then
                            Dim allLocKey As String: allLocKey = "M" & Format(mach, "000") & Format(dan, "00") & Format(retsu, "00")
                            dictAllHit(allLocKey) = dictAllHit(allLocKey) + 1
                        End If

                        If mach >= 1 And mach <= 46 Then
                            Dim zoneNum As Integer: zoneNum = Int((mach - 1) / 2) + 1 ' 1&2→1, 3&4→2 ... 45&46→23
                            Dim locKey As String: locKey = Format(mach, "00") & Format(dan, "00") & Format(retsu, "00")

                            ' ヒートマップ用：1～4号機も含めた全AB間口（中量棚のみ除く）でゾーン・対面情報を記録
                            If Not ((mach = 1 Or mach = 2) And retsu >= 6 And retsu <= 14) Then
                                dictItemZoneAll(locKey) = zoneNum
                                dictItemMachAll(locKey) = mach
                                currentFormationItemsAll(locKey) = 1
                            End If

                            If Not IsExcludedSlot3(mach, retsu) Then
                                dictItemLoc(locKey) = mach & "-" & Format(dan, "00") & "-" & Format(retsu, "00")
                                dictItemMach(locKey) = mach
                                dictItemZone(locKey) = zoneNum
                                dictItemHit(locKey) = dictItemHit(locKey) + 1
                                currentFormationItems(locKey) = 1
                            End If
                        End If
                    End If
                Next slotStart
            End If
        Loop
        Close #fileNo
    Next fIdx
    ' 最後の編成（6件に満たない端数も含む）を締める
    Call RecordZonePairs(currentFormationItems, dictPairs, dictCrossFace, dictItemZone, dictItemMach)
    Call RecordZonePairs(currentFormationItemsAll, dictPairsAll, dictCrossFaceAll, dictItemZoneAll, dictItemMachAll)

    ' 2.5 現在の奇数号機側・偶数号機側の合計ヒット数を算出（バランス調整の基準値。以降スワップのたびに更新する）
    Dim oddTotal As Double, evenTotal As Double
    oddTotal = 0: evenTotal = 0
    Dim hitKey As Variant
    For Each hitKey In dictItemHit.Keys
        If dictItemMach(hitKey) Mod 2 = 1 Then
            oddTotal = oddTotal + dictItemHit(hitKey)
        Else
            evenTotal = evenTotal + dictItemHit(hitKey)
        End If
    Next hitKey
    Dim oddTotalStart As Double, evenTotalStart As Double
    oddTotalStart = oddTotal: evenTotalStart = evenTotal

    ' 3. ペアスコアの計算と配列化
    Dim pairArr() As Variant
    Dim maxPairs As Long: maxPairs = dictPairs.Count
    If maxPairs = 0 Then maxPairs = 1
    ReDim pairArr(1 To maxPairs, 1 To 8)
    Dim pCnt As Long: pCnt = 0

    Dim pairKey As Variant
    For Each pairKey In dictPairs.Keys
        Dim items() As String: items = Split(pairKey, ",")
        Dim itemA As String: itemA = items(0)
        Dim itemB As String: itemB = items(1)

        pCnt = pCnt + 1
        Dim coCount As Long: coCount = dictPairs(pairKey)
        Dim isCross As Boolean: isCross = dictCrossFace.Exists(pairKey)
        Dim scoreVal As Long: scoreVal = coCount * IIf(isCross, 2, 1) ' 対面ペアを優先(重み2倍)

        ' 総ヒット数が多い方をアンカー(据え置き)、少ない方をムーバー(移動対象)とする
        Dim anchorItem As String, moverItem As String
        If dictItemHit(itemA) >= dictItemHit(itemB) Then
            anchorItem = itemA: moverItem = itemB
        Else
            anchorItem = itemB: moverItem = itemA
        End If

        pairArr(pCnt, 1) = dictItemZone(anchorItem)
        pairArr(pCnt, 2) = scoreVal
        pairArr(pCnt, 3) = coCount
        pairArr(pCnt, 4) = IIf(isCross, "対面", "同面")
        pairArr(pCnt, 5) = anchorItem
        pairArr(pCnt, 6) = dictItemLoc(anchorItem)
        pairArr(pCnt, 7) = moverItem
        pairArr(pCnt, 8) = dictItemLoc(moverItem)
    Next pairKey

    If pCnt = 0 Then GoTo RestoreSettings

    ' ゼロ落ち防止処置をして一時シートでスコア降順にソート
    Dim wsTemp As Worksheet: Set wsTemp = Sheets.Add
    wsTemp.Columns("E:H").NumberFormat = "@"
    wsTemp.Range("A1").Resize(pCnt, 8).Value = pairArr
    wsTemp.Sort.SortFields.Clear
    wsTemp.Sort.SortFields.Add Key:=wsTemp.Range("B1:B" & pCnt), Order:=xlDescending
    wsTemp.Sort.SetRange wsTemp.Range("A1:H" & pCnt)
    wsTemp.Sort.Apply
    pairArr = wsTemp.Range("A1:H" & pCnt).Value

    ' 4. アイテムを総ヒット数昇順に整理し、ゾーンごとにコレクション化（入替対象探索の高速化）
    Dim itemArr() As Variant, iCnt As Long
    iCnt = dictItemHit.Count
    ReDim itemArr(1 To iCnt, 1 To 3)
    Dim r As Long: r = 1
    Dim k As Variant
    For Each k In dictItemHit.Keys
        itemArr(r, 1) = k
        itemArr(r, 2) = dictItemZone(k)
        itemArr(r, 3) = dictItemHit(k)
        r = r + 1
    Next k

    wsTemp.Cells.Clear
    wsTemp.Columns("A:A").NumberFormat = "@"
    wsTemp.Range("A1").Resize(iCnt, 3).Value = itemArr
    wsTemp.Sort.SortFields.Clear
    wsTemp.Sort.SortFields.Add Key:=wsTemp.Range("C1:C" & iCnt), Order:=xlAscending
    wsTemp.Sort.SetRange wsTemp.Range("A1:C" & iCnt)
    wsTemp.Sort.Apply
    itemArr = wsTemp.Range("A1:C" & iCnt).Value
    wsTemp.Delete

    Dim zoneItems As Object: Set zoneItems = CreateObject("Scripting.Dictionary")
    For r = 1 To iCnt
        Dim zStr As String: zStr = CStr(itemArr(r, 2))
        If Not zoneItems.Exists(zStr) Then zoneItems.Add zStr, New Collection
        zoneItems(zStr).Add CStr(itemArr(r, 1))
    Next r

    ' 5. 交換候補の決定（アンカーとは別ゾーンの低頻度アイテムを入替対象にする）
    Dim outArr() As Variant
    ReDim outArr(1 To pCnt, 1 To 13)
    Dim outCnt As Long: outCnt = 0
    Dim dictSwapped As Object: Set dictSwapped = CreateObject("Scripting.Dictionary")
    Dim dictZoneUsedCount As Object: Set dictZoneUsedCount = CreateObject("Scripting.Dictionary") ' 入替先ゾーンの採用回数（分散のため上限を設ける）
    Const MAX_PER_ZONE As Integer = 2 ' 同じゾーンを入替先に採用できる回数の上限

    For r = 1 To pCnt
        If outCnt >= 10 Then Exit For ' 相関度(スコア)上位10件まで

        Dim aItem As String: aItem = CStr(pairArr(r, 5))
        Dim mItem As String: mItem = CStr(pairArr(r, 7))

        If Not dictSwapped.Exists(aItem) And Not dictSwapped.Exists(mItem) Then
            Dim anchorZone As Integer: anchorZone = dictItemZone(aItem)
            Dim targetItem As String: targetItem = ""

            ' 奇数・偶数バランスを踏まえた入替先の希望サイドを決定
            ' ムーバーが「多い側」にいるなら反対側(少ない側)へ、「少ない側」にいるなら同じ側で入れ替えて是正を妨げない
            Dim moverSide As Integer: moverSide = dictItemMach(mItem) Mod 2 ' 1=奇数, 0=偶数
            Dim desiredSide As Integer
            If Abs(oddTotal - evenTotal) <= 0.001 Then
                desiredSide = -1 ' ほぼ均衡しているのでサイドにこだわらない
            ElseIf (oddTotal > evenTotal And moverSide = 1) Or (evenTotal > oddTotal And moverSide = 0) Then
                desiredSide = 1 - moverSide
            Else
                desiredSide = moverSide
            End If

            ' パス1:サイド希望＋ゾーン上限を両方満たす／パス2:ゾーン上限のみ／パス3:制約なし(最終手段)
            Dim passNum As Integer
            For passNum = 1 To 3
                If targetItem <> "" Then Exit For
                Dim zKey As Variant
                For Each zKey In zoneItems.Keys
                    If CInt(zKey) <> anchorZone Then
                        Dim zoneUsed As Integer
                        If dictZoneUsedCount.Exists(zKey) Then zoneUsed = dictZoneUsedCount(zKey) Else zoneUsed = 0
                        If passNum = 3 Or zoneUsed < MAX_PER_ZONE Then
                            Dim candidate As Variant
                            For Each candidate In zoneItems(zKey)
                                Dim candStr As String: candStr = CStr(candidate)
                                If candStr <> aItem And candStr <> mItem And Not dictSwapped.Exists(candStr) Then
                                    If passNum = 1 And desiredSide <> -1 Then
                                        If dictItemMach(candStr) Mod 2 = desiredSide Then
                                            targetItem = candStr
                                            Exit For
                                        End If
                                    Else
                                        targetItem = candStr
                                        Exit For
                                    End If
                                End If
                            Next candidate
                        End If
                    End If
                    If targetItem <> "" Then Exit For
                Next zKey
            Next passNum

            If targetItem <> "" Then
                ' このゾーンの入替先採用回数をカウント（分散の判定に使用）
                Dim usedZoneKey As String: usedZoneKey = CStr(dictItemZone(targetItem))
                If dictZoneUsedCount.Exists(usedZoneKey) Then
                    dictZoneUsedCount(usedZoneKey) = dictZoneUsedCount(usedZoneKey) + 1
                Else
                    dictZoneUsedCount.Add usedZoneKey, 1
                End If

                ' 奇数・偶数の合計を更新（サイドが異なる場合のみバランスが動く）
                Dim targetSide As Integer: targetSide = dictItemMach(targetItem) Mod 2
                If moverSide <> targetSide Then
                    Dim moverHits As Double: moverHits = dictItemHit(mItem)
                    Dim targetHits As Double: targetHits = dictItemHit(targetItem)
                    If moverSide = 1 Then
                        oddTotal = oddTotal - moverHits + targetHits
                        evenTotal = evenTotal - targetHits + moverHits
                    Else
                        evenTotal = evenTotal - moverHits + targetHits
                        oddTotal = oddTotal - targetHits + moverHits
                    End If
                End If

                outCnt = outCnt + 1
                outArr(outCnt, 1) = anchorZone
                outArr(outCnt, 2) = pairArr(r, 4) ' 対面／同面
                outArr(outCnt, 3) = pairArr(r, 3) ' 編成内同時回数
                outArr(outCnt, 4) = GetLocName3(dictLocName, dictItemMach(aItem), aItem)
                outArr(outCnt, 5) = GetLocCode3(dictLocCode, dictItemMach(aItem), aItem)
                outArr(outCnt, 6) = dictItemLoc(aItem)
                outArr(outCnt, 7) = GetLocName3(dictLocName, dictItemMach(mItem), mItem)
                outArr(outCnt, 8) = GetLocCode3(dictLocCode, dictItemMach(mItem), mItem)
                outArr(outCnt, 9) = dictItemLoc(mItem)
                outArr(outCnt, 10) = "⇔"
                outArr(outCnt, 11) = GetLocName3(dictLocName, dictItemMach(targetItem), targetItem)
                outArr(outCnt, 12) = GetLocCode3(dictLocCode, dictItemMach(targetItem), targetItem)
                outArr(outCnt, 13) = dictItemLoc(targetItem)

                dictSwapped(mItem) = True
                dictSwapped(targetItem) = True
            End If
        End If
    Next r

    ' 6. 一括出力
    If outCnt > 0 Then
        Dim wsOut As Worksheet
        On Error Resume Next
        Sheets("AB編成流れ最適化").Delete
        On Error GoTo 0

        ' 「操作パネル」シートがあればその左側に配置する
        Dim wsPanel3 As Worksheet
        On Error Resume Next
        Set wsPanel3 = ThisWorkbook.Sheets("操作パネル")
        On Error GoTo 0
        If Not wsPanel3 Is Nothing Then
            Set wsOut = ThisWorkbook.Sheets.Add(Before:=wsPanel3)
        Else
            Set wsOut = Sheets.Add
        End If
        wsOut.name = "AB編成流れ最適化"

        wsOut.Columns("F:F").NumberFormat = "@"
        wsOut.Columns("I:I").NumberFormat = "@"
        wsOut.Columns("M:M").NumberFormat = "@"

        ' タイトル・サマリーはA:M列で結合し、A列だけが横に伸びないようにする
        wsOut.Range("A1:M1").Merge
        wsOut.Cells(1, 1).Value = "【AB編成流れ最適化（相関上位10件）】"
        wsOut.Cells(1, 1).Font.Bold = True: wsOut.Cells(1, 1).Font.Size = 14
        wsOut.Cells(1, 1).HorizontalAlignment = xlLeft

        wsOut.Range("A2:M2").Merge
        wsOut.Cells(2, 1).Value = "奇数号機合計ヒット数: " & Format(oddTotalStart, "0") & " → " & Format(oddTotal, "0") & _
            "　／　偶数号機合計ヒット数: " & Format(evenTotalStart, "0") & " → " & Format(evenTotal, "0") & _
            "（乖離: " & Format(Abs(oddTotalStart - evenTotalStart), "0") & " → " & Format(Abs(oddTotal - evenTotal), "0") & "）"
        wsOut.Cells(2, 1).HorizontalAlignment = xlLeft

        wsOut.Range("A4:M4").Value = Array("ゾーン", "対面区分", "編成内同時回数", "【起点品名】(動かさない)", "起点商品コード", "起点ロケーション", "【相関品名】(ずらしたい)", "相関商品コード", "相関ロケーション", "⇒交換⇒", "【入替対象品名】(別ゾーンの低頻度品)", "入替対象商品コード", "入替ロケーション")
        wsOut.Range("A5").Resize(outCnt, 13).Value = outArr

        wsOut.Range("A4:M4").Interior.Color = RGB(220, 230, 255)
        wsOut.Range("A4:M4").Font.Bold = True
        wsOut.Columns("A:M").AutoFit

        ' 対面同時ヒット状況のヒートマップ（1&2号機～45&46号機の物理配置順に23ゾーンを帯状に表示）
        ' ※1～4号機はスワップ候補・AB号機使用比率スコアの対象外だが、ヒートマップは実態を見るためdictPairsAll(全AB間口)を使う
        Dim zoneCrossHit(1 To 23) As Double
        Dim pk As Variant, pkParts() As String
        For Each pk In dictPairsAll.Keys
            If dictCrossFaceAll.Exists(pk) Then
                pkParts = Split(CStr(pk), ",")
                If dictItemZoneAll.Exists(pkParts(0)) Then
                    Dim pZone As Integer: pZone = CInt(dictItemZoneAll(pkParts(0)))
                    If pZone >= 1 And pZone <= 23 Then
                        zoneCrossHit(pZone) = zoneCrossHit(pZone) + dictPairsAll(pk)
                    End If
                End If
            End If
        Next pk

        Dim maxCross As Double: maxCross = 0
        Dim zi As Integer
        For zi = 1 To 23
            If zoneCrossHit(zi) > maxCross Then maxCross = zoneCrossHit(zi)
        Next zi

        Dim heatTitleRow As Long: heatTitleRow = 4 + outCnt + 3
        Dim heatLabelRow As Long: heatLabelRow = heatTitleRow + 1
        Dim heatValueRow As Long: heatValueRow = heatTitleRow + 2

        ' 本表(A:M)と列を共有すると列幅が本表側に引っ張られて広くなるため、O列(15列目)以降の未使用列にコンパクトな幅で配置する
        Const HEAT_COL_OFFSET As Long = 14 ' 15列目(O)から開始
        Dim heatFirstCol As Long: heatFirstCol = HEAT_COL_OFFSET + 1
        Dim heatLastCol As Long: heatLastCol = HEAT_COL_OFFSET + 23

        wsOut.Range(wsOut.Cells(heatTitleRow, heatFirstCol), wsOut.Cells(heatTitleRow, heatLastCol)).Merge
        wsOut.Cells(heatTitleRow, heatFirstCol).Value = "【対面同時ヒット状況（ゾーン別ヒートマップ）】　色が濃いほど対面での同時出荷（同じ編成内での競合）が多い"
        wsOut.Cells(heatTitleRow, heatFirstCol).Font.Bold = True: wsOut.Cells(heatTitleRow, heatFirstCol).Font.Size = 12
        wsOut.Cells(heatTitleRow, heatFirstCol).HorizontalAlignment = xlLeft

        wsOut.Range(wsOut.Cells(heatLabelRow, heatFirstCol), wsOut.Cells(heatLabelRow, heatLastCol)).EntireColumn.ColumnWidth = 6

        For zi = 1 To 23
            Dim heatCol As Long: heatCol = HEAT_COL_OFFSET + zi
            wsOut.Cells(heatLabelRow, heatCol).Value = (zi * 2 - 1) & "&" & (zi * 2)
            wsOut.Cells(heatLabelRow, heatCol).Font.Size = 8
            wsOut.Cells(heatLabelRow, heatCol).HorizontalAlignment = xlCenter

            wsOut.Cells(heatValueRow, heatCol).Value = zoneCrossHit(zi)
            wsOut.Cells(heatValueRow, heatCol).HorizontalAlignment = xlCenter
            wsOut.Cells(heatValueRow, heatCol).Font.Bold = True

            Dim crossRatio As Double
            If maxCross > 0 Then crossRatio = zoneCrossHit(zi) / maxCross Else crossRatio = 0
            Dim gb As Integer: gb = 255 - CInt(155 * crossRatio) ' 0件=白、最大件数=濃い赤
            wsOut.Cells(heatValueRow, heatCol).Interior.Color = RGB(255, gb, gb)
        Next zi
        wsOut.Range(wsOut.Cells(heatLabelRow, heatFirstCol), wsOut.Cells(heatValueRow, heatLastCol)).Borders.LineStyle = xlContinuous

        ' KPI記録：AB号機使用比率スコア（号機回数比の目標比率＝理論値 と、実績ファイル集計＝実績値 の近さ）
        Dim abRatioScore As Variant: abRatioScore = ""
        Dim wsRatio3 As Worksheet
        On Error Resume Next
        Set wsRatio3 = ActiveWorkbook.Sheets("号機回数比")
        On Error GoTo 0
        If Not wsRatio3 Is Nothing Then
            Dim machHit(1 To 46) As Double, machTarget(1 To 46) As Double
            Dim hk As Variant
            For Each hk In dictItemHit.Keys
                Dim hm As Integer: hm = dictItemMach(hk)
                If hm >= 1 And hm <= 46 Then machHit(hm) = machHit(hm) + dictItemHit(hk)
            Next hk

            Dim rr3 As Long, abLabel3 As String, mNum3 As Integer
            For rr3 = 3 To 48 ' AB01(1号機)～AB46(46号機)に対応する行
                abLabel3 = Trim(CStr(wsRatio3.Cells(rr3, 1).Value))
                If abLabel3 Like "AB##" Then
                    mNum3 = CInt(Mid(abLabel3, 3, 2))
                    If mNum3 >= 1 And mNum3 <= 46 Then machTarget(mNum3) = Val(wsRatio3.Cells(rr3, 5).Value)
                End If
            Next rr3

            ' 実績で対象となった号機(5～46、1～4号機は集計対象外)だけで正規化して比較する
            Dim hitTotal As Double, targetTotal As Double, mIdx As Integer
            hitTotal = 0: targetTotal = 0
            For mIdx = 5 To 46
                hitTotal = hitTotal + machHit(mIdx)
                targetTotal = targetTotal + machTarget(mIdx)
            Next mIdx

            If hitTotal > 0 And targetTotal > 0 Then
                Dim sumAbsDiff As Double: sumAbsDiff = 0
                For mIdx = 5 To 46
                    sumAbsDiff = sumAbsDiff + Abs((machHit(mIdx) / hitTotal) - (machTarget(mIdx) / targetTotal))
                Next mIdx
                ' 乖離合計(sumAbsDiff)が0.6(理論上の最大2.0の約1/3)以上で0点、0で100点、その間は比例配分
                abRatioScore = Application.WorksheetFunction.Max(0, 100 * (1 - sumAbsDiff / 0.6))
            End If
        End If

        ' KPI記録：AB占有率スコア（理論値：実績全体の回数上位900アイテムの回数比率／実績値：AB間口の実績回数比率）
        Dim abOccupancyScore As Variant: abOccupancyScore = ""
        Dim abTheoreticalRatioOut As Variant: abTheoreticalRatioOut = ""
        Dim abActualRatioOut As Variant: abActualRatioOut = ""
        If dictAllHit.Count > 0 Then
            Dim wsTempAll As Worksheet: Set wsTempAll = Sheets.Add
            Dim allKey As Variant, ar As Long: ar = 1
            Dim grandTotal As Double: grandTotal = 0
            Dim abActualTotal As Double: abActualTotal = 0
            For Each allKey In dictAllHit.Keys
                Dim hitVal As Double: hitVal = dictAllHit(allKey)
                wsTempAll.Cells(ar, 1).Value = hitVal
                grandTotal = grandTotal + hitVal
                Dim keyMach As Integer: keyMach = CInt(Mid(CStr(allKey), 2, 3))
                If keyMach >= 1 And keyMach <= 46 Then abActualTotal = abActualTotal + hitVal
                ar = ar + 1
            Next allKey
            Dim allN As Long: allN = ar - 1

            Dim topSum As Double
            If allN >= 900 Then
                ' 回数の多い順に並べ替えてから、ちょうど上位900件だけを合計する
                ' （Large+SumIf(">=")だと同数タイのロケーションが全部含まれてしまい、900件を超えて合計される不具合があったため修正）
                wsTempAll.Range("A1:A" & allN).Sort Key1:=wsTempAll.Range("A1"), Order1:=xlDescending, Header:=xlNo
                topSum = Application.WorksheetFunction.Sum(wsTempAll.Range("A1:A900"))
            Else
                topSum = grandTotal
            End If
            wsTempAll.Delete

            If grandTotal > 0 Then
                Dim abTheoreticalRatio As Double: abTheoreticalRatio = topSum / grandTotal
                Dim abActualRatio As Double: abActualRatio = abActualTotal / grandTotal
                abTheoreticalRatioOut = abTheoreticalRatio
                abActualRatioOut = abActualRatio
                Dim occDeviation As Double: occDeviation = Abs(abTheoreticalRatio - abActualRatio)
                If occDeviation <= 0.01 Then
                    abOccupancyScore = 100
                ElseIf occDeviation >= 0.06 Then
                    abOccupancyScore = 0
                Else
                    abOccupancyScore = 100 * (0.06 - occDeviation) / 0.05
                End If
            End If
        End If

        ' KPI記録：奇数偶数比率スコア・対面同時ヒットスコア・実績日（B行の日付があれば優先、無ければファイル更新日時）
        Dim reportDate As Date
        If latestBDate > DateSerial(1900, 1, 1) Then
            reportDate = latestBDate
        Else
            reportDate = DateSerial(Year(latestFileDate), Month(latestFileDate), Day(latestFileDate))
        End If
        On Error Resume Next
        Module7.LogFormationScore oddTotalStart, evenTotalStart, oddTotal, evenTotal, dictCrossFace.Count, dictPairs.Count, abRatioScore, abOccupancyScore, abTheoreticalRatioOut, abActualRatioOut, reportDate
        On Error GoTo 0

        MsgBox "「AB編成流れ最適化」の作成が完了しました。（" & fd.SelectedItems.Count & "ファイル読込／" & outCnt & "件の交換候補）" & vbCrLf & _
            "奇数偶数の乖離: " & Format(Abs(oddTotalStart - evenTotalStart), "0") & " → " & Format(Abs(oddTotal - evenTotal), "0"), vbInformation
    Else
        MsgBox "交換候補が見つかりませんでした。", vbExclamation
    End If

RestoreSettings:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' ----------------------------------------------------
' 補助関数群
' ----------------------------------------------------

' 編成内で同一ゾーンとなるアイテムペアを記録し、対面(異なる号機)かどうかも記録する
Sub RecordZonePairs(currentItems As Object, dictPairs As Object, dictCrossFace As Object, dictZone As Object, dictMach As Object)
    If currentItems.Count < 2 Then Exit Sub
    Dim itemsArr() As Variant: itemsArr = currentItems.Keys
    Dim i As Long, j As Long
    For i = 0 To UBound(itemsArr) - 1
        For j = i + 1 To UBound(itemsArr)
            Dim item1 As String: item1 = CStr(itemsArr(i))
            Dim item2 As String: item2 = CStr(itemsArr(j))
            If dictZone.Exists(item1) And dictZone.Exists(item2) Then
                If dictZone(item1) = dictZone(item2) Then
                    Dim pairKey As String
                    If item1 < item2 Then pairKey = item1 & "," & item2 Else pairKey = item2 & "," & item1
                    dictPairs(pairKey) = dictPairs(pairKey) + 1
                    If dictMach(item1) <> dictMach(item2) Then
                        If Not dictCrossFace.Exists(pairKey) Then dictCrossFace.Add pairKey, True
                    End If
                End If
            End If
        Next j
    Next i
End Sub

' 1～4号機はサイズが異なるアイテムを格納しているため対象から除外する
Function IsExcludedSlot3(mach As Integer, col As Integer) As Boolean
    IsExcludedSlot3 = (mach >= 1 And mach <= 4)
End Function

Function GetLocName3(dictLocName As Object, ByVal mach As Long, locKey As String) As String
    Dim dan As String, retsu As String
    dan = Mid(locKey, 3, 2)
    retsu = Mid(locKey, 5, 2)
    Dim locCode As String
    locCode = CStr(CLng(mach) * 10000& + CLng(Val(dan)) * 100& + CLng(Val(retsu)))
    If dictLocName.Exists(locCode) Then
        GetLocName3 = dictLocName(locCode)
    Else
        GetLocName3 = "(品名不明)"
    End If
End Function

Function GetLocCode3(dictLocCode As Object, ByVal mach As Long, locKey As String) As Variant
    Dim dan As String, retsu As String
    dan = Mid(locKey, 3, 2)
    retsu = Mid(locKey, 5, 2)
    Dim locCode As String
    locCode = CStr(CLng(mach) * 10000& + CLng(Val(dan)) * 100& + CLng(Val(retsu)))
    If dictLocCode.Exists(locCode) Then
        GetLocCode3 = dictLocCode(locCode)
    Else
        GetLocCode3 = ""
    End If
End Function
