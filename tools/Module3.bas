Attribute VB_Name = "Module3"
Option Explicit

Sub OptimizeABFormationFlow()
    Dim fd As Office.FileDialog
    Dim filePath As String
    Dim fileNo As Integer, textLine As String

    Dim dictItemLoc As Object, dictItemHit As Object, dictItemMach As Object, dictItemZone As Object
    Set dictItemLoc = CreateObject("Scripting.Dictionary")
    Set dictItemHit = CreateObject("Scripting.Dictionary")
    Set dictItemMach = CreateObject("Scripting.Dictionary")
    Set dictItemZone = CreateObject("Scripting.Dictionary")

    Dim dictPairs As Object: Set dictPairs = CreateObject("Scripting.Dictionary") ' 同一ゾーン内アイテムペアの共起回数
    Dim dictCrossFace As Object: Set dictCrossFace = CreateObject("Scripting.Dictionary") ' そのペアが対面(異なる機番)かどうか
    Dim currentFormationItems As Object: Set currentFormationItems = CreateObject("Scripting.Dictionary")
    Dim orderCountInFormation As Long: orderCountInFormation = 0
    Dim dictAllHit As Object: Set dictAllHit = CreateObject("Scripting.Dictionary") ' AB稼働率スコア用:全ゾーンのヒット数

    ' ヒートマップ表示用(スワップ対象外の1～4番機も含めた全AB番号のゾーン・面情報)
    ' ※スワップ候補・AB稼働率スコアの対象(除外ロケーション設定を反映)は変わらないが、ヒートマップは実際の稼働状況を反映する
    Dim dictItemZoneAll As Object: Set dictItemZoneAll = CreateObject("Scripting.Dictionary")
    Dim dictItemMachAll As Object: Set dictItemMachAll = CreateObject("Scripting.Dictionary")
    Dim dictPairsAll As Object: Set dictPairsAll = CreateObject("Scripting.Dictionary")
    Dim dictCrossFaceAll As Object: Set dictCrossFaceAll = CreateObject("Scripting.Dictionary")
    Dim currentFormationItemsAll As Object: Set currentFormationItemsAll = CreateObject("Scripting.Dictionary")

    ' 拠点カスタマイズ設定:除外機番・除外ロケーション・除外品コードを「設定」シートから読み込む
    Dim dictExcludedMach As Object: Set dictExcludedMach = CreateObject("Scripting.Dictionary") ' スワップ対象外にする機番
    Dim excludedLocMach() As Long, excludedLocDanFrom() As Long, excludedLocDanTo() As Long, excludedLocColFrom() As Long, excludedLocColTo() As Long
    Dim excludedLocCount As Long: excludedLocCount = 0
    Dim dictExcludedItemCode As Object: Set dictExcludedItemCode = CreateObject("Scripting.Dictionary") ' 全ての集計・スワップ対象から除外する品コード

    ' 0. CFシート読込 ロケーション番号と品名・品コードの対応表を作る
    Dim dictLocName As Object: Set dictLocName = CreateObject("Scripting.Dictionary")
    Dim dictLocCode As Object: Set dictLocCode = CreateObject("Scripting.Dictionary")
    Dim wsCF As Worksheet
    On Error Resume Next
    Set wsCF = ActiveWorkbook.Sheets("CF")
    On Error GoTo 0
    If Not wsCF Is Nothing Then
        Dim lastCF As Long: lastCF = wsCF.Cells(wsCF.Rows.Count, "B").End(xlUp).Row
        Dim cf As Long
        For cf = 2 To lastCF
            Dim locCode As String: locCode = Trim(CStr(wsCF.Cells(cf, 2).Value)) ' B列:ロケーション番号(機番*10000+段*100+列)
            If locCode <> "" And Not dictLocName.Exists(locCode) Then
                dictLocName.Add locCode, CStr(wsCF.Cells(cf, 9).Value) ' I列:品名
                Dim codeVal As Variant
                If IsNumeric(wsCF.Cells(cf, 8).Value) Then
                    codeVal = CLng(wsCF.Cells(cf, 8).Value)
                Else
                    codeVal = wsCF.Cells(cf, 8).Value
                End If
                dictLocCode.Add locCode, codeVal ' H列:品コード
            End If
        Next cf
    End If

    ' 0.5 拠点カスタマイズ設定の読込(「設定」シートが無ければ従来どおりの初期値で自動生成)
    Dim ratioSheetName As String: ratioSheetName = "機番回数比"
    Dim maxSwapRows As Long: maxSwapRows = 15
    Call EnsureExclusionSettingsSheet
    Call LoadExclusionSettings(dictExcludedMach, excludedLocMach, excludedLocDanFrom, excludedLocDanTo, excludedLocColFrom, excludedLocColTo, excludedLocCount, dictExcludedItemCode, ratioSheetName, maxSwapRows)

    ' 1. ファイル選択(複数選択・全ファイル形式)
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "ピッキング実績ファイルを選択(複数選択可)"
        .Filters.Clear
        .Filters.Add "すべてのファイル", "*.*"
        .AllowMultiSelect = True
        If .Show = False Then Exit Sub
    End With

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    ' 2. データの読込(H行6件を1編成として区切り、AB(1～46番機、除外設定を反映)を対象に集計)
    ' 複数編成がファイルをまたがない前提のため、ファイルが変わるたびに前ファイルの端数編成を確定させてリセットする
    Dim fIdx As Long
    Dim latestFileDate As Date: latestFileDate = DateSerial(1900, 1, 1) ' ファイル更新日時(B行から日付が読めない場合のフォールバック)
    Dim latestBDate As Date: latestBDate = DateSerial(1900, 1, 1) ' B行(先頭"B"+8桁日付)から読み取った最も新しい日付
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
        Dim skipMode As Boolean: skipMode = False ' H99999(棚卸等の在庫サマリー行)配下は読み飛ばす
        Open filePath For Input As #fileNo
        Do While Not EOF(fileNo)
            Line Input #fileNo, textLine
            If Left(textLine, 1) = "B" And Len(textLine) >= 9 Then
                ' B行の2～9文字目(8桁)が集計日(YYYYMMDD)
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
                    ' 在庫サマリー行。直前の編成を確定させ、以降のE行(在庫全数)はオーダーとして扱わない
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

                        ' 除外品コード(設定シートで指定)に該当する品は、格納場所を問わず全ての集計・スワップ対象から除く
                        Dim itemCodeExcluded As Boolean
                        itemCodeExcluded = IsExcludedItemCode(dictLocCode, dictExcludedItemCode, mach, dan, retsu)

                        ' AB稼働率スコア用:全ゾーン(機番の範囲を問わず)のヒット数を集計(実在番のみ対象)
                        ' 除外ロケーション(常時使用の固定スロットなど)がデータに混ざっていると全体回数が水増しされ、
                        '   理論比率・実績比率とも本来の値からズレるため、拠点設定に応じて除外する
                        If mach > 0 And Not itemCodeExcluded And Not IsExcludedLocation(mach, dan, retsu, excludedLocMach, excludedLocDanFrom, excludedLocDanTo, excludedLocColFrom, excludedLocColTo, excludedLocCount) Then
                            Dim allLocKey As String: allLocKey = "M" & Format(mach, "000") & Format(dan, "00") & Format(retsu, "00")
                            dictAllHit(allLocKey) = dictAllHit(allLocKey) + 1
                        End If

                        If mach >= 1 And mach <= 46 Then
                            Dim zoneNum As Integer: zoneNum = Int((mach - 1) / 2) + 1 ' 1&2番機は1、3&4番機は2 … 45&46番機は23
                            Dim locKey As String: locKey = Format(mach, "00") & Format(dan, "00") & Format(retsu, "00")

                            ' ヒートマップ用:除外機番も含めた全AB番号(除外ロケーション・除外品コードのみ除く)でゾーン・面情報を記録
                            If Not itemCodeExcluded And Not IsExcludedLocation(mach, dan, retsu, excludedLocMach, excludedLocDanFrom, excludedLocDanTo, excludedLocColFrom, excludedLocColTo, excludedLocCount) Then
                                dictItemZoneAll(locKey) = zoneNum
                                dictItemMachAll(locKey) = mach
                                currentFormationItemsAll(locKey) = 1
                            End If

                            If Not itemCodeExcluded And Not IsExcludedSlot3(dictExcludedMach, mach) And Not IsExcludedLocation(mach, dan, retsu, excludedLocMach, excludedLocDanFrom, excludedLocDanTo, excludedLocColFrom, excludedLocColTo, excludedLocCount) Then
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
    ' 最後の編成(6件に満たない端数を含む)を確定させる
    Call RecordZonePairs(currentFormationItems, dictPairs, dictCrossFace, dictItemZone, dictItemMach)
    Call RecordZonePairs(currentFormationItemsAll, dictPairsAll, dictCrossFaceAll, dictItemZoneAll, dictItemMachAll)

    ' 2.5 現在の奇数機番・偶数機番の合計ヒット数を算出(左右バランスの基準値。以降スワップのたびに更新する)
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

        ' よりヒット数が多い方をアンカー(据え置き)、少ない方をムーバー(移動対象)とする
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

    ' 端数配列を防止しつつ一時シートでスコア降順にソート
    Dim wsTemp As Worksheet: Set wsTemp = Sheets.Add
    wsTemp.Columns("E:H").NumberFormat = "@"
    wsTemp.Range("A1").Resize(pCnt, 8).Value = pairArr
    wsTemp.Sort.SortFields.Clear
    wsTemp.Sort.SortFields.Add Key:=wsTemp.Range("B1:B" & pCnt), Order:=xlDescending
    wsTemp.Sort.SetRange wsTemp.Range("A1:H" & pCnt)
    wsTemp.Sort.Apply
    pairArr = wsTemp.Range("A1:H" & pCnt).Value

    ' 4. アイテムを総ヒット数昇順に整列し、ゾーンごとにコレクション化(交換対象探索の高速化)
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

    ' 5. 入替案の決定(アンカーとは別ゾーンの低頻度アイテムを交換対象とする)
    Dim outArr() As Variant
    ReDim outArr(1 To pCnt, 1 To 13)
    Dim outCnt As Long: outCnt = 0
    Dim dictSwapped As Object: Set dictSwapped = CreateObject("Scripting.Dictionary")
    Dim dictZoneUsedCount As Object: Set dictZoneUsedCount = CreateObject("Scripting.Dictionary") ' 交換先ゾーンの採用回数(偏りを防ぐため)
    Const MAX_PER_ZONE As Integer = 2 ' 同一ゾーンから交換先に採用できる回数の上限

    For r = 1 To pCnt
        If outCnt >= maxSwapRows Then Exit For ' 入替候補(スコア順)は設定件数まで

        Dim aItem As String: aItem = CStr(pairArr(r, 5))
        Dim mItem As String: mItem = CStr(pairArr(r, 7))

        If Not dictSwapped.Exists(aItem) And Not dictSwapped.Exists(mItem) Then
            Dim anchorZone As Integer: anchorZone = dictItemZone(aItem)
            Dim targetItem As String: targetItem = ""

            ' 奇数・偶数バランスを踏まえた交換先の希望サイドを決定
            ' ムーバーが「奇数側」にいるなら反対側(偶数)へ、「偶数側」にいるなら同様の側で入替えて偏りを広げないようにする
            Dim moverSide As Integer: moverSide = dictItemMach(mItem) Mod 2 ' 1=奇数, 0=偶数
            Dim desiredSide As Integer
            If Abs(oddTotal - evenTotal) <= 0.001 Then
                desiredSide = -1 ' ほぼ均衡しているのでサイドにこだわらない
            ElseIf (oddTotal > evenTotal And moverSide = 1) Or (evenTotal > oddTotal And moverSide = 0) Then
                desiredSide = 1 - moverSide
            Else
                desiredSide = moverSide
            End If

            ' パス1:サイド指定+ゾーン利用上限あり両方満たす/パス2:ゾーン利用上限のみ/パス3:制限なし(最終手段)
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
                ' この交換先ゾーンの利用回数をカウント(偏りの判定に使用)
                Dim usedZoneKey As String: usedZoneKey = CStr(dictItemZone(targetItem))
                If dictZoneUsedCount.Exists(usedZoneKey) Then
                    dictZoneUsedCount(usedZoneKey) = dictZoneUsedCount(usedZoneKey) + 1
                Else
                    dictZoneUsedCount.Add usedZoneKey, 1
                End If

                ' 奇数・偶数の合計を更新(サイドが異なる場合のみバランスが変化する)
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
                outArr(outCnt, 2) = pairArr(r, 4) ' 対面/同面
                outArr(outCnt, 3) = pairArr(r, 3) ' 編成内共起回数
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
        Sheets("AB編成動線最適化").Delete
        Sheets("対面化促進スワップ指示").Delete ' 旧バージョンで作成された出力シートが残っていれば削除する
        On Error GoTo 0

        ' 「操作パネル」シートがあればその左隣に配置する
        Dim wsPanel3 As Worksheet
        On Error Resume Next
        Set wsPanel3 = ThisWorkbook.Sheets("操作パネル")
        On Error GoTo 0
        If Not wsPanel3 Is Nothing Then
            Set wsOut = ThisWorkbook.Sheets.Add(Before:=wsPanel3)
        Else
            Set wsOut = Sheets.Add
        End If
        wsOut.Name = "AB編成動線最適化"

        wsOut.Columns("F:F").NumberFormat = "@"
        wsOut.Columns("I:I").NumberFormat = "@"
        wsOut.Columns("M:M").NumberFormat = "@"

        ' タイトル・サマリー行はA:M列で結合し、A列だけが横に伸びないようにする
        wsOut.Range("A1:M1").Merge
        wsOut.Cells(1, 1).Value = "【AB編成動線最適化(入替候補" & maxSwapRows & "件)】"
        wsOut.Cells(1, 1).Font.Bold = True: wsOut.Cells(1, 1).Font.Size = 14
        wsOut.Cells(1, 1).HorizontalAlignment = xlLeft

        wsOut.Range("A2:M2").Merge
        wsOut.Cells(2, 1).Value = "奇数機番合計ヒット数: " & Format(oddTotalStart, "0") & " → " & Format(oddTotal, "0") & _
            "　／　偶数機番合計ヒット数: " & Format(evenTotalStart, "0") & " → " & Format(evenTotal, "0") & _
            "(差: " & Format(Abs(oddTotalStart - evenTotalStart), "0") & " → " & Format(Abs(oddTotal - evenTotal), "0") & ")"
        wsOut.Cells(2, 1).HorizontalAlignment = xlLeft

        wsOut.Range("A4:M4").Value = Array("ゾーン", "対面区分", "編成内共起回数", "【起点品】(動かさない)", "起点品コード", "起点ロケーション", "【交換品】(こちらを動かす)", "交換品コード", "交換元ロケーション", "交換方向", "【交換対象品】(別ゾーンの低頻度品)", "交換対象品コード", "交換先ロケーション")
        wsOut.Range("A5").Resize(outCnt, 13).Value = outArr

        wsOut.Range("A4:M4").Interior.Color = RGB(220, 230, 255)
        wsOut.Range("A4:M4").Font.Bold = True
        wsOut.Columns("A:M").AutoFit

        ' 対面同士のヒット状況ヒートマップ(1&2番機～45&46番機を機番配置順に23ゾーン表示)
        ' ※1～4番機はスワップ候補・AB稼働率スコアの対象外だが、ヒートマップは実態を反映するためdictPairsAll(除外ロケーションのみ反映)を使う
        Dim zoneCrossHit(1 To 23) As Double
        Dim zoneTotalHit(1 To 23) As Double ' そのゾーンの同面込みヒット数(対面比率(%)の分母)
        Dim pk As Variant, pkParts() As String
        For Each pk In dictPairsAll.Keys
            pkParts = Split(CStr(pk), ",")
            If dictItemZoneAll.Exists(pkParts(0)) Then
                Dim pZone As Integer: pZone = CInt(dictItemZoneAll(pkParts(0)))
                If pZone >= 1 And pZone <= 23 Then
                    zoneTotalHit(pZone) = zoneTotalHit(pZone) + dictPairsAll(pk)
                    If dictCrossFaceAll.Exists(pk) Then
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
        Dim heatPctRow As Long: heatPctRow = heatTitleRow + 3

        ' 本表(A:M)と列幅を揃えると本表側の列幅が崩れるため、O列(15列目)以降の未使用領域にコンパクトな幅で配置する
        Const HEAT_COL_OFFSET As Long = 14 ' 15列目(O)から開始
        Dim heatFirstCol As Long: heatFirstCol = HEAT_COL_OFFSET + 1
        Dim heatLastCol As Long: heatLastCol = HEAT_COL_OFFSET + 23

        wsOut.Range(wsOut.Cells(heatTitleRow, heatFirstCol), wsOut.Cells(heatTitleRow, heatLastCol)).Merge
        wsOut.Cells(heatTitleRow, heatFirstCol).Value = "【対面同士のヒット状況(ゾーン別ヒートマップ)】※濃いほど対面での同時出庫(同一編成での共起)が多い。下段はそのゾーン内の同時ヒットのうち対面が占める割合"
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

            ' ゾーン内対面比率(%):そのゾーンの同時ヒットのうち対面が占める割合(機番選定の精度を見る指標)
            Dim zonePct As Double
            If zoneTotalHit(zi) > 0 Then zonePct = zoneCrossHit(zi) / zoneTotalHit(zi) * 100 Else zonePct = 0
            wsOut.Cells(heatPctRow, heatCol).Value = zonePct / 100
            wsOut.Cells(heatPctRow, heatCol).NumberFormat = "0%"
            wsOut.Cells(heatPctRow, heatCol).Font.Size = 8
            wsOut.Cells(heatPctRow, heatCol).HorizontalAlignment = xlCenter
        Next zi
        wsOut.Range(wsOut.Cells(heatLabelRow, heatFirstCol), wsOut.Cells(heatPctRow, heatLastCol)).Borders.LineStyle = xlContinuous

        ' KPI記録:AB稼働率スコア(機番回数比の目標比率実績値と、今回ファイル集計結果との近さ)
        Dim abRatioScore As Variant: abRatioScore = ""
        Dim abRatioScoreNote As String: abRatioScoreNote = ""
        Dim wsRatio3 As Worksheet
        On Error Resume Next
        Set wsRatio3 = ActiveWorkbook.Sheets(ratioSheetName)
        On Error GoTo 0
        If wsRatio3 Is Nothing Then
            abRatioScoreNote = "「" & ratioSheetName & "」シートが見つからないため、AB稼働率スコアは算出されていません(「設定」シートL4でシート名を確認してください)"
        Else
            Dim machHit(1 To 46) As Double, machTarget(1 To 46) As Double
            Dim hk As Variant
            For Each hk In dictItemHit.Keys
                Dim hm As Integer: hm = dictItemMach(hk)
                If hm >= 1 And hm <= 46 Then machHit(hm) = machHit(hm) + dictItemHit(hk)
            Next hk

            Dim rr3 As Long, abLabel3 As String, mNum3 As Integer
            For rr3 = 3 To 48 ' AB01(1番機)～AB46(46番機)に対応する行
                abLabel3 = Trim(CStr(wsRatio3.Cells(rr3, 1).Value))
                If abLabel3 Like "AB##" Then
                    mNum3 = CInt(Mid(abLabel3, 3, 2))
                    If mNum3 >= 1 And mNum3 <= 46 Then machTarget(mNum3) = Val(wsRatio3.Cells(rr3, 5).Value)
                End If
            Next rr3

            ' 1号機～46号機を対象に、設定シートの除外機番だけを動的に除いて正規化して比較する
            Dim hitTotal As Double, targetTotal As Double, mIdx As Integer
            hitTotal = 0: targetTotal = 0
            For mIdx = 1 To 46
                If Not dictExcludedMach.Exists(CStr(mIdx)) Then
                    hitTotal = hitTotal + machHit(mIdx)
                    targetTotal = targetTotal + machTarget(mIdx)
                End If
            Next mIdx

            If hitTotal <= 0 Then
                abRatioScoreNote = "実績データが除外設定によりすべて対象外のため、AB稼働率スコアは算出されていません"
            ElseIf targetTotal <= 0 Then
                abRatioScoreNote = "「" & ratioSheetName & "」シートにAB01～AB46の目標比率(A列ラベル・E列数値、3～48行目)が見つからないため、AB稼働率スコアは算出されていません"
            Else
                Dim sumAbsDiff As Double: sumAbsDiff = 0
                For mIdx = 1 To 46
                    If Not dictExcludedMach.Exists(CStr(mIdx)) Then
                        sumAbsDiff = sumAbsDiff + Abs((machHit(mIdx) / hitTotal) - (machTarget(mIdx) / targetTotal))
                    End If
                Next mIdx
                ' 差の合計(sumAbsDiff)が0.6(理論上の最大2.0の約1/3)以上で0点、0で100点、その間は線形
                abRatioScore = Application.WorksheetFunction.Max(0, 100 * (1 - sumAbsDiff / 0.6))
            End If
        End If

        ' KPI記録:AB得意先スコア(理論値:全体の回数上位900アイテムの回数比率／実績値:AB番機の実回数比率)
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
                ' 回数の多い順に並べ替えて、ちょうど上位900件だけ合計する
                ' (LARGE+SUMIF(">=")式だと同着タイのロケーションが全部含まれてしまい、900件を超えて合計されることがあるため補正)
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

        ' KPI記録:対面化ヒットスコア(ゾーンごとに機番の奇数/偶数の組み方まで含めて最適配置した場合の
        ' 「理論上最小の対面ヒット数」に対して、実績の対面ヒット数がどれだけ近いかで評価する)
        Dim actualCrossFaceHits As Double: actualCrossFaceHits = 0
        Dim theoreticalMinCrossFace As Double: theoreticalMinCrossFace = 0
        Dim zoneItemsDict As Object: Set zoneItemsDict = CreateObject("Scripting.Dictionary") ' zone -> Dictionary(item->1)
        Dim zoneWeightDict As Object: Set zoneWeightDict = CreateObject("Scripting.Dictionary") ' zone -> Dictionary(pairKey->weight)
        Dim pk3 As Variant, pkParts3() As String
        For Each pk3 In dictPairs.Keys
            pkParts3 = Split(CStr(pk3), ",")
            If dictItemZone.Exists(pkParts3(0)) Then
                Dim zz3 As Integer: zz3 = dictItemZone(pkParts3(0))
                If Not zoneItemsDict.Exists(zz3) Then Set zoneItemsDict(zz3) = CreateObject("Scripting.Dictionary")
                zoneItemsDict(zz3)(pkParts3(0)) = 1
                zoneItemsDict(zz3)(pkParts3(1)) = 1
                If Not zoneWeightDict.Exists(zz3) Then Set zoneWeightDict(zz3) = CreateObject("Scripting.Dictionary")
                zoneWeightDict(zz3)(CStr(pk3)) = dictPairs(pk3)
                If dictCrossFace.Exists(pk3) Then actualCrossFaceHits = actualCrossFaceHits + dictPairs(pk3)
            End If
        Next pk3

        Dim zoneKeyMC As Variant
        For Each zoneKeyMC In zoneItemsDict.Keys
            Dim zItemsArr() As Variant: zItemsArr = zoneItemsDict(zoneKeyMC).Keys
            theoreticalMinCrossFace = theoreticalMinCrossFace + ComputeZoneMinCut(zItemsArr, dictItemMach, zoneWeightDict(zoneKeyMC))
        Next zoneKeyMC

        Dim crossFaceScore As Variant: crossFaceScore = ""
        If actualCrossFaceHits > 0 Then
            crossFaceScore = Application.WorksheetFunction.Min(100, 100 * theoreticalMinCrossFace / actualCrossFaceHits)
        Else
            crossFaceScore = 100
        End If

        ' KPI記録:実施日は、B行の日付があればそれを優先し、無ければファイル更新日時を使う
        Dim reportDate As Date
        If latestBDate > DateSerial(1900, 1, 1) Then
            reportDate = latestBDate
        Else
            reportDate = DateSerial(Year(latestFileDate), Month(latestFileDate), Day(latestFileDate))
        End If
        On Error Resume Next
        Module7.LogFormationScore oddTotalStart, evenTotalStart, oddTotal, evenTotal, crossFaceScore, abRatioScore, abOccupancyScore, abTheoreticalRatioOut, abActualRatioOut, reportDate
        On Error GoTo 0

        Dim completeMsg As String
        completeMsg = "「AB編成動線最適化」の作成が完了しました。(" & fd.SelectedItems.Count & "ファイル読込／" & outCnt & "件の入替案)" & vbCrLf & _
            "左右機番の差: " & Format(Abs(oddTotalStart - evenTotalStart), "0") & " → " & Format(Abs(oddTotal - evenTotal), "0")
        If abRatioScoreNote <> "" Then completeMsg = completeMsg & vbCrLf & "※" & abRatioScoreNote
        MsgBox completeMsg, vbInformation
    Else
        MsgBox "入替候補が見つかりませんでした。", vbExclamation
    End If

RestoreSettings:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' ----------------------------------------------------
' 補助関数群
' ----------------------------------------------------

' 編成内で同一ゾーンとなるアイテムペアを記録し、対面(異なる機番)かどうかも記録する
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

' 1つのゾーン内で、奇数機番・偶数機番の組み方まで含めて最適配置した場合の
' 「理論上最小の対面ヒット数」を局所探索(Kernighan-Linに近い2分割法)で求める。
' itemsArr: そのゾーンに属するアイテムキーの配列／dictMach: アイテム→機番／weightDict: "item1,item2"(ソート済)→編成内共起回数
Function ComputeZoneMinCut(itemsArr() As Variant, dictMach As Object, weightDict As Object) As Double
    Dim n As Long: n = UBound(itemsArr) - LBound(itemsArr) + 1
    If n <= 1 Then ComputeZoneMinCut = 0: Exit Function

    Dim lo As Long: lo = LBound(itemsArr)
    Dim hi As Long: hi = UBound(itemsArr)
    Dim i As Long, j As Long, k As Long

    ' 重み参照を高速化するため、隣接リストを事前構築
    Dim adjKeys() As String, adjVals() As Double, adjCount() As Long
    ReDim adjCount(lo To hi)
    Dim maxAdj As Long: maxAdj = n
    ReDim adjKeys(lo To hi, 0 To maxAdj - 1)
    ReDim adjVals(lo To hi, 0 To maxAdj - 1)
    For i = lo To hi
        adjCount(i) = 0
    Next i
    For i = lo To hi - 1
        For j = i + 1 To hi
            Dim wKey As String
            If itemsArr(i) < itemsArr(j) Then wKey = itemsArr(i) & "," & itemsArr(j) Else wKey = itemsArr(j) & "," & itemsArr(i)
            If weightDict.Exists(wKey) Then
                Dim wVal As Double: wVal = weightDict(wKey)
                adjKeys(i, adjCount(i)) = CStr(j): adjVals(i, adjCount(i)) = wVal: adjCount(i) = adjCount(i) + 1
                adjKeys(j, adjCount(j)) = CStr(i): adjVals(j, adjCount(j)) = wVal: adjCount(j) = adjCount(j) + 1
            End If
        Next j
    Next i

    Dim nSide1 As Long: nSide1 = 0
    For i = lo To hi
        If dictMach(itemsArr(i)) Mod 2 = 1 Then nSide1 = nSide1 + 1
    Next i

    Randomize
    Dim bestCut As Double: bestCut = -1
    Dim restartIdx As Long
    For restartIdx = 1 To 4 ' 1回目=実際の配置、2～4回目=ランダム配置から局所探索
        Dim assign() As Integer: ReDim assign(lo To hi)
        If restartIdx = 1 Then
            For i = lo To hi
                assign(i) = IIf(dictMach(itemsArr(i)) Mod 2 = 1, 1, 0)
            Next i
        Else
            Dim order() As Long: ReDim order(lo To hi)
            For i = lo To hi: order(i) = i: Next i
            For i = hi To lo + 1 Step -1
                Dim rIdx As Long: rIdx = Int(Rnd() * (i - lo + 1)) + lo
                Dim tmp As Long: tmp = order(i): order(i) = order(rIdx): order(rIdx) = tmp
            Next i
            For i = lo To hi
                assign(order(i)) = IIf((i - lo) < nSide1, 1, 0)
            Next i
        End If

        Dim curCut As Double: curCut = 0
        For i = lo To hi
            For k = 0 To adjCount(i) - 1
                j = CLng(adjKeys(i, k))
                If j > i Then
                    If assign(i) <> assign(j) Then curCut = curCut + adjVals(i, k)
                End If
            Next k
        Next i

        Dim improved As Boolean: improved = True
        Do While improved
            improved = False
            For i = lo To hi - 1
                For j = i + 1 To hi
                    If assign(i) <> assign(j) Then
                        Dim delta As Double: delta = 0
                        For k = 0 To adjCount(i) - 1
                            Dim nb As Long: nb = CLng(adjKeys(i, k))
                            If nb <> j Then
                                Dim oldV As Integer, newV As Integer
                                oldV = IIf(assign(i) <> assign(nb), 1, 0)
                                newV = IIf(assign(j) <> assign(nb), 1, 0)
                                delta = delta + (newV - oldV) * adjVals(i, k)
                            End If
                        Next k
                        For k = 0 To adjCount(j) - 1
                            nb = CLng(adjKeys(j, k))
                            If nb <> i Then
                                oldV = IIf(assign(j) <> assign(nb), 1, 0)
                                newV = IIf(assign(i) <> assign(nb), 1, 0)
                                delta = delta + (newV - oldV) * adjVals(j, k)
                            End If
                        Next k
                        If delta < -0.0001 Then
                            Dim tmpA As Integer: tmpA = assign(i): assign(i) = assign(j): assign(j) = tmpA
                            curCut = curCut + delta
                            improved = True
                        End If
                    End If
                Next j
            Next i
        Loop

        If bestCut < 0 Or curCut < bestCut Then bestCut = curCut
    Next restartIdx

    ComputeZoneMinCut = bestCut
End Function

' ----------------------------------------------------
' 拠点カスタマイズ設定(除外機番・除外ロケーション)
' ----------------------------------------------------

' 「設定」シートが無い場合、これまでの固定値(機番1～4を完全除外、1・2番機の列6～14を除外)を
' 初期値として自動生成する。他拠点ではこのシートの値を書き換えるだけで動作を変更できる。
' 列幅は用途ごとに固定値で設定する(説明文の長さに引っ張られて横に広がらないようにするため、AutoFitは使わない)。
Sub EnsureExclusionSettingsSheet()
    Dim wsSet As Worksheet
    On Error Resume Next
    Set wsSet = ThisWorkbook.Sheets("設定")
    On Error GoTo 0
    If Not wsSet Is Nothing Then Exit Sub

    Set wsSet = ThisWorkbook.Sheets.Add
    wsSet.Name = "設定"

    wsSet.Columns("A:A").ColumnWidth = 10  ' 除外機番
    wsSet.Columns("B:B").ColumnWidth = 3   ' 区切り
    wsSet.Columns("C:G").ColumnWidth = 8   ' 除外ロケーション(機番/段From/段To/列From/列To)
    wsSet.Columns("H:H").ColumnWidth = 3   ' 区切り
    wsSet.Columns("I:I").ColumnWidth = 14  ' 除外品コード
    wsSet.Columns("I:I").NumberFormat = "@" ' 品コードは先頭0落ち・数値化を防ぐため文字列扱いにする
    wsSet.Columns("J:J").ColumnWidth = 3   ' 区切り
    wsSet.Columns("K:K").ColumnWidth = 20  ' シート名設定ラベル
    wsSet.Columns("L:L").ColumnWidth = 16  ' シート名設定値

    wsSet.Range("A1:I1").Merge
    wsSet.Range("A1").Value = "AB編成動線最適化で除外する条件をここで設定します。①除外機番:スワップ対象・AB稼働率スコアから機番ごと除外。②除外ロケーション:常時使用スロットなど機番×段×列の範囲を、スワップ対象・稼働率・ヒートマップ集計のすべてから除外(段From/To・列From/Toはそれぞれ空欄にすると「全段」「全列」扱いになる)。③除外品コード:その品コードを格納場所を問わず全ての集計・スワップ対象から除外(CFシートの品コード列と同じ値で指定)。各表の5行目以降に追加・削除して使ってください。"
    wsSet.Range("A1").Font.Bold = True
    wsSet.Range("A1").WrapText = True
    wsSet.Range("A1").VerticalAlignment = xlTop
    wsSet.Rows(1).RowHeight = 60

    wsSet.Range("A3").Value = "■除外機番"
    wsSet.Range("A3").Font.Bold = True
    wsSet.Range("A4").Value = "機番"
    wsSet.Range("A4").Font.Bold = True
    wsSet.Range("A5").Value = 1
    wsSet.Range("A6").Value = 2
    wsSet.Range("A7").Value = 3
    wsSet.Range("A8").Value = 4

    wsSet.Range("C3").Value = "■除外ロケーション"
    wsSet.Range("C3").Font.Bold = True
    wsSet.Range("C4").Value = "機番": wsSet.Range("D4").Value = "段From": wsSet.Range("E4").Value = "段To": wsSet.Range("F4").Value = "列From": wsSet.Range("G4").Value = "列To"
    wsSet.Range("C4:G4").Font.Bold = True
    wsSet.Range("C5").Value = 1: wsSet.Range("F5").Value = 6: wsSet.Range("G5").Value = 14 ' 段From/Toは空欄=全段
    wsSet.Range("C6").Value = 2: wsSet.Range("F6").Value = 6: wsSet.Range("G6").Value = 14

    wsSet.Range("I3").Value = "■除外品コード"
    wsSet.Range("I3").Font.Bold = True
    wsSet.Range("I4").Value = "品コード"
    wsSet.Range("I4").Font.Bold = True

    wsSet.Range("K3").Value = "■シート名設定"
    wsSet.Range("K3").Font.Bold = True
    wsSet.Range("K4").Value = "機番回数比シート名"
    wsSet.Range("K4").Font.Bold = True
    wsSet.Range("L4").Value = "機番回数比" ' AB稼働率スコアの目標比率を読むシート名。拠点によって名前が違う場合はここを書き換える

    wsSet.Range("K5").Value = "入替候補件数"
    wsSet.Range("K5").Font.Bold = True
    wsSet.Range("L5").Value = 15 ' 「AB編成動線最適化」に出力する入替候補の最大行数
End Sub

' 「設定」シートの内容を読み込み、除外機番・除外品コードの辞書と除外ロケーションの配列、シート名・件数設定を組み立てる
Sub LoadExclusionSettings(dictExcludedMach As Object, ByRef locMach() As Long, ByRef locDanFrom() As Long, ByRef locDanTo() As Long, ByRef locColFrom() As Long, ByRef locColTo() As Long, ByRef locCount As Long, dictExcludedItemCode As Object, ByRef ratioSheetName As String, ByRef maxSwapRows As Long)
    locCount = 0
    ReDim locMach(1 To 1)
    ReDim locDanFrom(1 To 1)
    ReDim locDanTo(1 To 1)
    ReDim locColFrom(1 To 1)
    ReDim locColTo(1 To 1)
    ratioSheetName = "機番回数比"
    maxSwapRows = 15

    Dim wsSet As Worksheet
    On Error Resume Next
    Set wsSet = ThisWorkbook.Sheets("設定")
    On Error GoTo 0
    If wsSet Is Nothing Then Exit Sub

    ' 機番回数比シート名(L4)。空欄ならデフォルト名のまま
    If Trim(CStr(wsSet.Range("L4").Value)) <> "" Then ratioSheetName = Trim(CStr(wsSet.Range("L4").Value))

    ' 入替候補件数(L5)。1以上の数値が入っていればそれを使う
    If IsNumeric(wsSet.Range("L5").Value) Then
        If CLng(wsSet.Range("L5").Value) >= 1 Then maxSwapRows = CLng(wsSet.Range("L5").Value)
    End If

    ' 除外機番リスト(A列、5行目以降)
    Dim lastA As Long: lastA = wsSet.Cells(wsSet.Rows.Count, "A").End(xlUp).Row
    Dim rA As Long
    For rA = 5 To lastA
        If IsNumeric(wsSet.Cells(rA, 1).Value) And Trim(CStr(wsSet.Cells(rA, 1).Value)) <> "" Then
            dictExcludedMach(CStr(CLng(wsSet.Cells(rA, 1).Value))) = True
        End If
    Next rA

    ' 除外ロケーションリスト(C:G列=機番/段From/段To/列From/列To、5行目以降)
    Dim lastC As Long: lastC = wsSet.Cells(wsSet.Rows.Count, "C").End(xlUp).Row
    If lastC >= 5 Then
        ReDim locMach(1 To lastC - 4)
        ReDim locDanFrom(1 To lastC - 4)
        ReDim locDanTo(1 To lastC - 4)
        ReDim locColFrom(1 To lastC - 4)
        ReDim locColTo(1 To lastC - 4)
        Dim rC As Long
        For rC = 5 To lastC
            If IsNumeric(wsSet.Cells(rC, 3).Value) And Trim(CStr(wsSet.Cells(rC, 3).Value)) <> "" Then
                locCount = locCount + 1
                locMach(locCount) = CLng(wsSet.Cells(rC, 3).Value)
                locDanFrom(locCount) = IIf(IsNumeric(wsSet.Cells(rC, 4).Value), CLng(wsSet.Cells(rC, 4).Value), 0)
                locDanTo(locCount) = IIf(IsNumeric(wsSet.Cells(rC, 5).Value), CLng(wsSet.Cells(rC, 5).Value), 0)
                locColFrom(locCount) = IIf(IsNumeric(wsSet.Cells(rC, 6).Value), CLng(wsSet.Cells(rC, 6).Value), 0)
                locColTo(locCount) = IIf(IsNumeric(wsSet.Cells(rC, 7).Value), CLng(wsSet.Cells(rC, 7).Value), 0)
            End If
        Next rC
    End If

    ' 除外品コードリスト(I列、5行目以降)
    Dim lastI As Long: lastI = wsSet.Cells(wsSet.Rows.Count, "I").End(xlUp).Row
    Dim rI As Long
    For rI = 5 To lastI
        Dim codeStr As String: codeStr = Trim(CStr(wsSet.Cells(rI, 9).Value))
        If codeStr <> "" Then dictExcludedItemCode(codeStr) = True
    Next rI
End Sub

' 指定の機番・段・列が「除外ロケーション」設定に該当するか判定する(段From/To・列From/Toはそれぞれ両方0なら「全段」「全列」の意味になる)
Function IsExcludedLocation(mach As Integer, dan As Integer, col As Integer, locMach() As Long, locDanFrom() As Long, locDanTo() As Long, locColFrom() As Long, locColTo() As Long, locCount As Long) As Boolean
    Dim i As Long
    For i = 1 To locCount
        If locMach(i) = mach Then
            Dim danMatch As Boolean, colMatch As Boolean
            danMatch = (locDanFrom(i) = 0 And locDanTo(i) = 0) Or (dan >= locDanFrom(i) And dan <= locDanTo(i))
            colMatch = (locColFrom(i) = 0 And locColTo(i) = 0) Or (col >= locColFrom(i) And col <= locColTo(i))
            If danMatch And colMatch Then
                IsExcludedLocation = True
                Exit Function
            End If
        End If
    Next i
    IsExcludedLocation = False
End Function

' 指定の機番が「除外機番」設定に該当するか判定する(1～4番機のような、サイズが異なる品を格納する機番など)
Function IsExcludedSlot3(dictExcludedMach As Object, mach As Integer) As Boolean
    IsExcludedSlot3 = dictExcludedMach.Exists(CStr(mach))
End Function

' 指定の機番・段・列にある品が「除外品コード」設定に該当するか判定する(CFシートのロケーション⇔品コード対応表を使って引く)
Function IsExcludedItemCode(dictLocCode As Object, dictExcludedItemCode As Object, ByVal mach As Integer, ByVal dan As Integer, ByVal retsu As Integer) As Boolean
    If dictExcludedItemCode.Count = 0 Then Exit Function
    Dim locCodeKey As String: locCodeKey = CStr(CLng(mach) * 10000& + CLng(dan) * 100& + CLng(retsu))
    If dictLocCode.Exists(locCodeKey) Then
        IsExcludedItemCode = dictExcludedItemCode.Exists(Trim(CStr(dictLocCode(locCodeKey))))
    End If
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
