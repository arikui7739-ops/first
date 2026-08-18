Attribute VB_Name = "Module3_石狩"
Option Explicit

' 実績データのキャッシュ(1度読み込んだピッキング実績ファイルを、「設定」シートの条件を変えながら
' 何度も使い回せるようにする。ブックを閉じるかVBAプロジェクトが再初期化されると空に戻る)
Private g_DataLoaded As Boolean
Private g_CachedFormations As Collection ' 編成ごとのCollection。各要素は"機番,段,列"形式の文字列
Private g_CachedLatestFileDate As Date
Private g_CachedLatestBDate As Date
Private g_CachedFileCount As Long
Private g_CachedDictLocName As Object
Private g_CachedDictLocCode As Object
Private g_CachedDictItemCategory As Object
Private g_CachedDictItemWeightMaster As Object
Private g_CachedDictItemVolumeMaster As Object

' カテゴリー集中ペナルティで「サイズが近い」と判定する閾値(体積比lnの絶対値。0.4は1.5倍以内)
Public Const SIZE_SIMILAR_RATIO As Double = 0.4

Sub OptimizeABFormationFlow()
    Dim fd As Office.FileDialog
    Dim filePath As String
    Dim fileNo As Integer, textLine As String

    Dim dictItemLoc As Object, dictItemHit As Object, dictItemMach As Object, dictItemZone As Object
    Set dictItemLoc = CreateObject("Scripting.Dictionary")
    Set dictItemHit = CreateObject("Scripting.Dictionary")
    Set dictItemMach = CreateObject("Scripting.Dictionary")
    Set dictItemZone = CreateObject("Scripting.Dictionary")

    ' 在庫データ(任意、WF021L1形式)による入替候補の絞り込み用。属性マスタが未読込なら常に空のままで、
    ' 挙動は従来どおり(スコアに影響しない)になる
    Dim dictItemCat As Object: Set dictItemCat = CreateObject("Scripting.Dictionary") ' locKey→大分類コード
    Dim dictItemWt As Object: Set dictItemWt = CreateObject("Scripting.Dictionary") ' locKey→重量(kg)
    Dim dictItemVol As Object: Set dictItemVol = CreateObject("Scripting.Dictionary") ' locKey→体積(縦×横×高)
    Dim dictMachCatVol As Object: Set dictMachCatVol = CreateObject("Scripting.Dictionary") ' "機番|大分類コード"→その機番内の同カテゴリー品数

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
    Dim dictTargetRatio As Object: Set dictTargetRatio = CreateObject("Scripting.Dictionary") ' 機番別目標構成比(機番→0～1の比率)

    Dim dictLocName As Object, dictLocCode As Object
    Dim dictItemCategory As Object, dictItemWeightMaster As Object, dictItemVolumeMaster As Object

    ' 生データ読込用(キャッシュ再利用時は使わないが、Dimはプロシージャ全体で有効なのでここでまとめて宣言する)
    Dim wsCF As Worksheet
    Dim lastCF As Long, cf As Long
    Dim locCode As String, codeVal As Variant
    Dim fIdx As Long
    Dim thisFileDate As Date
    Dim skipMode As Boolean
    Dim bDateStr As String, bDate As Date
    Dim slotStart As Long, rec As String
    Dim mach As Integer, dan As Integer, retsu As Integer
    Dim itemCodeExcluded As Boolean
    Dim allLocKey As String
    Dim zoneNum As Integer
    Dim locKey As String
    Dim currentFormationRaw As Collection

    ' 0.3 前回読み込んだ実績データが残っていれば、再利用するか確認する
    ' (「設定」シートの条件だけを変えて何度も試したいときに、ファイル選択をやり直さずに済む)
    Dim useCache As Boolean: useCache = False
    Dim latestFileDate As Date, latestBDate As Date, selectedFileCount As Long
    If g_DataLoaded Then
        Dim reuseResp As VbMsgBoxResult
        reuseResp = MsgBox("前回読み込んだ実績データ(" & g_CachedFileCount & "ファイル分)があります。" & vbCrLf & vbCrLf & _
            "「はい」: ファイルを選び直さず、「設定」シートの現在の内容を反映して再集計する" & vbCrLf & _
            "「いいえ」: ファイルを選び直す(実績データを更新する場合)", _
            vbQuestion + vbYesNoCancel, "実績データの再利用")
        If reuseResp = vbCancel Then Exit Sub
        useCache = (reuseResp = vbYes)
    End If

    ' 0.4 「操作パネル」シート(説明・実行ボタン)が無ければ自動生成する
    Call EnsureOperationPanelSheet

    ' 0.5 拠点カスタマイズ設定の読込(「設定」シートが無ければ従来どおりの初期値で自動生成)
    ' ※キャッシュ再利用時も、設定の変更を反映するため必ず読み直す
    Dim ratioSheetName As String: ratioSheetName = "機番回数比"
    Dim maxSwapRows As Long: maxSwapRows = 15
    Dim maxMachNum As Long: maxMachNum = 46 ' この機番までを集計・スワップ対象の範囲とする(拠点のラック総数に合わせて設定シートで変更可能)
    Dim abSlotCount As Long: abSlotCount = 900 ' ABの間口数(AB得意先スコアの理論値算出に使う上位件数)
    Dim catWeight As Double: catWeight = 0.005 ' 入替先号機の同カテゴリー品1件あたりの減点係数(在庫データ読込時のみ有効)
    Dim sizeWeight As Double: sizeWeight = 0.01 ' サイズ(体積)差1桁(対数比)あたりの減点係数
    Dim weightWeightCoef As Double: weightWeightCoef = 0.01 ' 重量差1桁(対数比)あたりの減点係数
    Call EnsureExclusionSettingsSheet
    Call LoadExclusionSettings(dictExcludedMach, excludedLocMach, excludedLocDanFrom, excludedLocDanTo, excludedLocColFrom, excludedLocColTo, excludedLocCount, dictExcludedItemCode, ratioSheetName, maxSwapRows, maxMachNum, abSlotCount, dictTargetRatio, catWeight, sizeWeight, weightWeightCoef)
    Dim maxZoneNum As Long: maxZoneNum = Int((maxMachNum - 1) / 2) + 1 ' 機番を2台単位で束ねたゾーン数

    If useCache Then
        ' --- キャッシュされた実績データをそのまま使う ---
        Set dictLocName = g_CachedDictLocName
        Set dictLocCode = g_CachedDictLocCode
        Set dictItemCategory = g_CachedDictItemCategory
        Set dictItemWeightMaster = g_CachedDictItemWeightMaster
        Set dictItemVolumeMaster = g_CachedDictItemVolumeMaster
        latestFileDate = g_CachedLatestFileDate
        latestBDate = g_CachedLatestBDate
        selectedFileCount = g_CachedFileCount

        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False
        Application.DisplayAlerts = False
    Else
        ' --- ファイルを選び直して読み込む ---
        Set dictLocName = CreateObject("Scripting.Dictionary")
        Set dictLocCode = CreateObject("Scripting.Dictionary")
        Set dictItemCategory = CreateObject("Scripting.Dictionary")
        Set dictItemWeightMaster = CreateObject("Scripting.Dictionary")
        Set dictItemVolumeMaster = CreateObject("Scripting.Dictionary")

        ' 0. CFシート読込 ロケーション番号と品名・品コードの対応表を作る
        On Error Resume Next
        Set wsCF = ActiveWorkbook.Sheets("CF")
        On Error GoTo 0
        If Not wsCF Is Nothing Then
            lastCF = wsCF.Cells(wsCF.Rows.Count, "B").End(xlUp).Row
            For cf = 2 To lastCF
                locCode = Trim(CStr(wsCF.Cells(cf, 2).Value)) ' B列:ロケーション番号(機番*10000+段*100+列)
                If locCode <> "" And Not dictLocName.Exists(locCode) Then
                    dictLocName.Add locCode, CStr(wsCF.Cells(cf, 9).Value) ' I列:品名
                    If IsNumeric(wsCF.Cells(cf, 8).Value) Then
                        codeVal = CLng(wsCF.Cells(cf, 8).Value)
                    Else
                        codeVal = wsCF.Cells(cf, 8).Value
                    End If
                    dictLocCode.Add locCode, codeVal ' H列:品コード
                End If
            Next cf
        End If

        ' 0.7 品名マスタ・ロケーションマスタ(任意)の読込。選べばCFシートの品名・品コードをこちらで上書き・補完する
        Call LoadItemMasterFilesIfSelected(dictLocCode, dictLocName)

        ' 0.8 在庫データ(任意、在庫状況ダウンロード=WF021L1形式)の読込。
        ' サイズ・重量・カテゴリーが分かれば、入替候補選定でこれまで人が目視判断していた
        ' 「同カテゴリーが集中しない」「サイズ・重量が近い」をスコアの目安に反映できる
        Call LoadItemAttributeMasterFromSheet(dictItemCategory, dictItemWeightMaster, dictItemVolumeMaster)

        ' 1. ファイル選択(複数選択・全ファイル形式)
        Set fd = Application.FileDialog(msoFileDialogFilePicker)
        With fd
            .Title = "ピッキング実績ファイル(S71で始まるファイル)を選択(複数選択可)"
            .Filters.Clear
            .Filters.Add "すべてのファイル", "*.*"
            .AllowMultiSelect = True
            If .Show = False Then Exit Sub
        End With

        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False
        Application.DisplayAlerts = False

        ' 2. データの読込(H行6件を1編成として区切り、除外設定は適用せず機番・段・列の生データのまま
        ' 編成ごとにキャッシュする。除外設定の適用は、この後のキャッシュ集計ステップで毎回行う)
        ' 複数編成がファイルをまたがない前提のため、ファイルが変わるたびに前ファイルの端数編成を確定させてリセットする
        Set g_CachedFormations = New Collection
        Set currentFormationRaw = New Collection
        orderCountInFormation = 0
        latestFileDate = DateSerial(1900, 1, 1) ' ファイル更新日時(B行から日付が読めない場合のフォールバック)
        latestBDate = DateSerial(1900, 1, 1) ' B行(先頭"B"+8桁日付)から読み取った最も新しい日付
        For fIdx = 1 To fd.SelectedItems.Count
            If fIdx > 1 Then
                If currentFormationRaw.Count > 0 Then g_CachedFormations.Add currentFormationRaw
                Set currentFormationRaw = New Collection
                orderCountInFormation = 0
            End If

            filePath = fd.SelectedItems(fIdx)
            thisFileDate = FileDateTime(filePath)
            If thisFileDate > latestFileDate Then latestFileDate = thisFileDate
            fileNo = FreeFile
            skipMode = False ' H99999(棚卸等の在庫サマリー行)配下は読み飛ばす
            Open filePath For Input As #fileNo
            Do While Not EOF(fileNo)
                Line Input #fileNo, textLine
                If Left(textLine, 1) = "B" And Len(textLine) >= 9 Then
                    ' B行の2～9文字目(8桁)が集計日(YYYYMMDD)
                    bDateStr = Mid(textLine, 2, 8)
                    If IsNumeric(bDateStr) Then
                        On Error Resume Next
                        bDate = DateSerial(CInt(Left(bDateStr, 4)), CInt(Mid(bDateStr, 5, 2)), CInt(Mid(bDateStr, 7, 2)))
                        On Error GoTo 0
                        If bDate > latestBDate Then latestBDate = bDate
                    End If
                ElseIf Left(textLine, 1) = "H" Then
                    If Mid(textLine, 2, 5) = "99999" Then
                        ' 在庫サマリー行。直前の編成を確定させ、以降のE行(在庫全数)はオーダーとして扱わない
                        If currentFormationRaw.Count > 0 Then g_CachedFormations.Add currentFormationRaw
                        Set currentFormationRaw = New Collection
                        skipMode = True
                    Else
                        skipMode = False
                        orderCountInFormation = orderCountInFormation + 1
                        If orderCountInFormation > 6 Then
                            If currentFormationRaw.Count > 0 Then g_CachedFormations.Add currentFormationRaw
                            Set currentFormationRaw = New Collection
                            orderCountInFormation = 1
                        End If
                    End If
                ElseIf Left(textLine, 1) = "E" And Len(textLine) >= 10 And Not skipMode Then
                    ' E行は13文字おきに最大3件のレコード(先頭9文字=機番2桁+段2桁+列2桁+3桁)が
                    ' 詰められていることがある。末尾の余白がスペース埋め(新形式)またはゼロ埋め
                    ' (旧形式)のいずれかで、ゼロ埋めの場合は余白がたまたま数字のみになり、
                    ' 「機番0・段0・列0」という実在しないレコードが混じることがあるため、
                    ' 機番0のレコードは明示的に除外する
                    For slotStart = 2 To Len(textLine) - 8 Step 13
                        rec = Mid(textLine, slotStart, 9)
                        If Trim(rec) <> "" And Len(Trim(rec)) = 9 And IsNumeric(rec) Then
                            mach = Val(Mid(rec, 1, 2))
                            dan = Val(Mid(rec, 3, 2))
                            retsu = Val(Mid(rec, 5, 2))
                            If mach >= 1 Then
                                currentFormationRaw.Add mach & "," & dan & "," & retsu
                            End If
                        End If
                    Next slotStart
                End If
            Loop
            Close #fileNo
        Next fIdx
        If currentFormationRaw.Count > 0 Then g_CachedFormations.Add currentFormationRaw

        ' 次回実行時に再利用できるよう、生データをキャッシュしておく(ブックを閉じるかVBAが再初期化されるまで有効)
        selectedFileCount = fd.SelectedItems.Count
        g_CachedLatestFileDate = latestFileDate
        g_CachedLatestBDate = latestBDate
        g_CachedFileCount = selectedFileCount
        Set g_CachedDictLocName = dictLocName
        Set g_CachedDictLocCode = dictLocCode
        Set g_CachedDictItemCategory = dictItemCategory
        Set g_CachedDictItemWeightMaster = dictItemWeightMaster
        Set g_CachedDictItemVolumeMaster = dictItemVolumeMaster
        g_DataLoaded = True
    End If

    ' 2.3 キャッシュされた編成データ(機番・段・列の生データ)に、現在の「設定」シートの除外条件・最大機番を
    ' 適用しながら集計する。ここで初めて除外設定を反映するため、キャッシュ再利用時も設定変更が正しく反映される
    Dim formationIter As Variant, recIter As Variant
    Dim recParts() As String
    For Each formationIter In g_CachedFormations
        currentFormationItems.RemoveAll
        currentFormationItemsAll.RemoveAll
        For Each recIter In formationIter
            recParts = Split(CStr(recIter), ",")
            mach = CInt(recParts(0))
            dan = CInt(recParts(1))
            retsu = CInt(recParts(2))

            ' 除外品コード(設定シートで指定)に該当する品は、格納場所を問わず全ての集計・スワップ対象から除く
            itemCodeExcluded = IsExcludedItemCode(dictLocCode, dictExcludedItemCode, mach, dan, retsu)

            ' AB稼働率スコア用:全ゾーン(機番の範囲を問わず)のヒット数を集計(実在番のみ対象)
            ' 除外ロケーション(常時使用の固定スロットなど)がデータに混ざっていると全体回数が水増しされ、
            '   理論比率・実績比率とも本来の値からズレるため、拠点設定に応じて除外する
            If mach > 0 And Not itemCodeExcluded And Not IsExcludedLocation(mach, dan, retsu, excludedLocMach, excludedLocDanFrom, excludedLocDanTo, excludedLocColFrom, excludedLocColTo, excludedLocCount) Then
                allLocKey = "M" & Format(mach, "000") & Format(dan, "00") & Format(retsu, "00")
                dictAllHit(allLocKey) = dictAllHit(allLocKey) + 1
            End If

            If mach >= 1 And mach <= maxMachNum Then
                zoneNum = Int((mach - 1) / 2) + 1 ' 1&2番機は1、3&4番機は2 … 45&46番機は23
                locKey = Format(mach, "00") & Format(dan, "00") & Format(retsu, "00")

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
                    If Not dictItemCat.Exists(locKey) Then
                        Call ResolveItemAttr(locKey, mach, dan, retsu, dictLocCode, dictItemCategory, dictItemWeightMaster, dictItemVolumeMaster, dictItemCat, dictItemWt, dictItemVol)
                    End If
                End If
            End If
        Next recIter
        Call RecordZonePairs(currentFormationItems, dictPairs, dictCrossFace, dictItemZone, dictItemMach)
        Call RecordZonePairs(currentFormationItemsAll, dictPairsAll, dictCrossFaceAll, dictItemZoneAll, dictItemMachAll)
    Next formationIter

    ' 2.5 現在の奇数機番・偶数機番の合計ヒット数を算出(左右バランスの基準値。以降スワップのたびに更新する)
    ' 併せて、機番ごとのヒット数(machHitStart)も算出しておく(目標構成比を考慮した交換先選定に使う)
    Dim oddTotal As Double, evenTotal As Double
    oddTotal = 0: evenTotal = 0
    Dim machHitStart() As Double
    ReDim machHitStart(1 To maxMachNum)
    Dim hitKey As Variant
    For Each hitKey In dictItemHit.Keys
        Dim hkMach As Long: hkMach = dictItemMach(hitKey)
        If hkMach Mod 2 = 1 Then
            oddTotal = oddTotal + dictItemHit(hitKey)
        Else
            evenTotal = evenTotal + dictItemHit(hitKey)
        End If
        machHitStart(hkMach) = machHitStart(hkMach) + dictItemHit(hitKey)
    Next hitKey

    ' 号機ごとのカテゴリー在庫点数を集計(入替先候補の号機に同カテゴリー品がどれだけ集中しているかの目安に使う)
    Dim catHitKey As Variant
    For Each catHitKey In dictItemHit.Keys
        If dictItemCat.Exists(catHitKey) Then
            Dim catTallyKey As String: catTallyKey = CStr(dictItemMach(catHitKey)) & "|" & dictItemCat(catHitKey)
            dictMachCatVol(catTallyKey) = dictMachCatVol(catTallyKey) + 1
            If dictItemVol.Exists(catHitKey) And dictItemVol(catHitKey) > 0 Then
                Dim bIdx As Long: bIdx = Int(Log(dictItemVol(catHitKey)) / SIZE_SIMILAR_RATIO)
                Dim bKey As String: bKey = catTallyKey & "|B" & bIdx
                dictMachCatVol(bKey) = dictMachCatVol(bKey) + 1
            Else
                Dim nKey As String: nKey = catTallyKey & "|N"
                dictMachCatVol(nKey) = dictMachCatVol(nKey) + 1
            End If
        End If
    Next catHitKey
    Dim oddTotalStart As Double, evenTotalStart As Double
    oddTotalStart = oddTotal: evenTotalStart = evenTotal

    ' 目標構成比は「設定」シートの合計が100%になっていなくても機番どうしの相対バランスとして扱えるよう、
    ' 目標比率の合計(targetRatioSum)と、目標が設定されている機番だけの実績ヒット合計(targetedHitStart)で
    ' それぞれ正規化してから比較する(Cバラ等AB以外への出荷分による絶対値のズレの影響を受けないようにする)。
    ' 「機番別目標構成比」にはC01・C02・Xのような機番以外のカテゴリ行(構成比グラフ用)が混在することがあるため、
    ' 数値の機番キーだけをスワップ判定の対象にする(数値以外のキーはCLngでエラーになるため必ず判定してから使う)
    Dim targetRatioSum As Double: targetRatioSum = 0
    Dim targetedHitStart As Double: targetedHitStart = 0
    Dim trKey As Variant
    For Each trKey In dictTargetRatio.Keys
        If IsNumeric(trKey) Then
            targetRatioSum = targetRatioSum + dictTargetRatio(trKey)
            Dim trMach As Long: trMach = CLng(trKey)
            If trMach >= 1 And trMach <= maxMachNum Then targetedHitStart = targetedHitStart + machHitStart(trMach)
        End If
    Next trKey
    Dim hasTargetRatioData As Boolean: hasTargetRatioData = (dictTargetRatio.Count > 0 And targetRatioSum > 0 And targetedHitStart > 0)

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

    ' 5. 入替案の決定(アンカーとは別ゾーンの低頻度アイテムを交換対象とする。
    ' 「設定」シートに機番別目標構成比が入力されていれば目標比率への近さを優先し、未入力なら奇数偶数バランスを優先する)
    Dim outArr() As Variant
    ReDim outArr(1 To pCnt, 1 To 13)
    Dim outCnt As Long: outCnt = 0
    Dim dictSwapped As Object: Set dictSwapped = CreateObject("Scripting.Dictionary")
    Dim dictZoneUsedCount As Object: Set dictZoneUsedCount = CreateObject("Scripting.Dictionary") ' 交換先ゾーンの採用回数(偏りを防ぐため)
    Const MAX_PER_ZONE As Integer = 2 ' 同一ゾーンから交換先に採用できる回数の上限

    ' 目標構成比の実績追跡用(このセクション内のスワップのたびに更新する)。目標構成比が未入力なら
    ' 従来どおり奇数・偶数バランス優先にフォールバックするため、machHitStartをコピーするだけで初期化する
    Dim machHitLive() As Double
    ReDim machHitLive(1 To maxMachNum)
    Dim mIdx2 As Long
    For mIdx2 = 1 To maxMachNum
        machHitLive(mIdx2) = machHitStart(mIdx2)
    Next mIdx2
    Dim targetedHitTotal As Double: targetedHitTotal = targetedHitStart

    For r = 1 To pCnt
        If outCnt >= maxSwapRows Then Exit For ' 入替候補(スコア順)は設定件数まで
        ' 候補ペアが多いと探索に時間がかかることがあるため、Excelが「応答なし」に見えないよう
        ' 一定回数ごとに制御をOSに戻す(処理自体は継続する)
        If r Mod 200 = 0 Then DoEvents

        Dim aItem As String: aItem = CStr(pairArr(r, 5))
        Dim mItem As String: mItem = CStr(pairArr(r, 7))

        If Not dictSwapped.Exists(aItem) And Not dictSwapped.Exists(mItem) Then
            Dim anchorZone As Integer: anchorZone = dictItemZone(aItem)
            Dim targetItem As String: targetItem = ""

            ' 奇数・偶数バランスを踏まえた交換先の希望サイドを決定(目標構成比が未入力の場合のフォールバック用)
            ' ムーバーが「奇数側」にいるなら反対側(偶数)へ、「偶数側」にいるなら同様の側で入替えて偏りを広げないようにする
            Dim moverSide As Integer: moverSide = dictItemMach(mItem) Mod 2 ' 1=奇数, 0=偶数

            ' パス1:目標構成比が入力されていれば、最も比率が不足している機番の候補をゾーン利用上限内で探す
            If hasTargetRatioData Then
                targetItem = FindBestUnderTargetCandidate(zoneItems, anchorZone, aItem, mItem, dictSwapped, dictItemMach, dictTargetRatio, machHitLive, targetedHitTotal, targetRatioSum, dictZoneUsedCount, True, MAX_PER_ZONE, dictItemCat, dictItemWt, dictItemVol, dictMachCatVol, catWeight, sizeWeight, weightWeightCoef)
            End If
            ' パス1':目標構成比が未入力なら、従来どおり希望サイド+ゾーン利用上限で探す
            If targetItem = "" And Not hasTargetRatioData Then
                Dim desiredSide As Integer
                If Abs(oddTotal - evenTotal) <= 0.001 Then
                    desiredSide = -1 ' ほぼ均衡しているのでサイドにこだわらない
                ElseIf (oddTotal > evenTotal And moverSide = 1) Or (evenTotal > oddTotal And moverSide = 0) Then
                    desiredSide = 1 - moverSide
                Else
                    desiredSide = moverSide
                End If
                If desiredSide <> -1 Then
                    Dim zKey As Variant
                    For Each zKey In zoneItems.Keys
                        If CInt(zKey) <> anchorZone Then
                            Dim zoneUsed As Integer
                            If dictZoneUsedCount.Exists(zKey) Then zoneUsed = dictZoneUsedCount(zKey) Else zoneUsed = 0
                            If zoneUsed < MAX_PER_ZONE Then
                                Dim candidate As Variant
                                For Each candidate In zoneItems(zKey)
                                    Dim candStr As String: candStr = CStr(candidate)
                                    If candStr <> aItem And candStr <> mItem And Not dictSwapped.Exists(candStr) Then
                                        If dictItemMach(candStr) Mod 2 = desiredSide Then
                                            targetItem = candStr
                                            Exit For
                                        End If
                                    End If
                                Next candidate
                            End If
                        End If
                        If targetItem <> "" Then Exit For
                    Next zKey
                End If
            End If
            ' パス2:ゾーン利用上限内で、比率・サイドを問わず最初に見つかった候補(在庫データがあればその中で一番属性が近い候補)
            If targetItem = "" Then
                targetItem = FindFirstCandidate(zoneItems, anchorZone, aItem, mItem, dictSwapped, dictZoneUsedCount, True, MAX_PER_ZONE, dictItemMach, dictItemCat, dictItemWt, dictItemVol, dictMachCatVol, catWeight, sizeWeight, weightWeightCoef)
            End If
            ' パス3:制限なしで、最初に見つかった候補(最終手段。在庫データがあればその中で一番属性が近い候補)
            If targetItem = "" Then
                targetItem = FindFirstCandidate(zoneItems, anchorZone, aItem, mItem, dictSwapped, dictZoneUsedCount, False, MAX_PER_ZONE, dictItemMach, dictItemCat, dictItemWt, dictItemVol, dictMachCatVol, catWeight, sizeWeight, weightWeightCoef)
            End If

            If targetItem <> "" Then
                ' この交換先ゾーンの利用回数をカウント(偏りの判定に使用)
                Dim usedZoneKey As String: usedZoneKey = CStr(dictItemZone(targetItem))
                If dictZoneUsedCount.Exists(usedZoneKey) Then
                    dictZoneUsedCount(usedZoneKey) = dictZoneUsedCount(usedZoneKey) + 1
                Else
                    dictZoneUsedCount.Add usedZoneKey, 1
                End If

                Dim moverHits As Double: moverHits = dictItemHit(mItem)
                Dim targetHits As Double: targetHits = dictItemHit(targetItem)

                ' 奇数・偶数の合計を更新(サイドが異なる場合のみバランスが変化する。KPI記録用に維持する)
                Dim targetSide As Integer: targetSide = dictItemMach(targetItem) Mod 2
                If moverSide <> targetSide Then
                    If moverSide = 1 Then
                        oddTotal = oddTotal - moverHits + targetHits
                        evenTotal = evenTotal - targetHits + moverHits
                    Else
                        evenTotal = evenTotal - moverHits + targetHits
                        oddTotal = oddTotal - targetHits + moverHits
                    End If
                End If

                ' 機番別の実績ヒット数を更新(目標構成比を考慮した交換先選定に使う)。
                ' targetedHitTotalは目標が設定されている機番だけの合計なので、対象機番が
                ' 目標未設定→設定 (またはその逆)に移る場合も正しく増減させる
                Dim mMach As Long: mMach = dictItemMach(mItem)
                Dim tMach As Long: tMach = dictItemMach(targetItem)
                If dictTargetRatio.Exists(CStr(mMach)) Then targetedHitTotal = targetedHitTotal - moverHits + targetHits
                If dictTargetRatio.Exists(CStr(tMach)) Then targetedHitTotal = targetedHitTotal - targetHits + moverHits
                machHitLive(mMach) = machHitLive(mMach) - moverHits + targetHits
                machHitLive(tMach) = machHitLive(tMach) - targetHits + moverHits

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
        Sheets("AB対面分散").Delete
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
        wsOut.Name = "AB対面分散"

        wsOut.Columns("F:F").NumberFormat = "@"
        wsOut.Columns("I:I").NumberFormat = "@"
        wsOut.Columns("M:M").NumberFormat = "@"

        ' タイトル・サマリー行はA:M列で結合し、A列だけが横に伸びないようにする
        wsOut.Range("A1:M1").Merge
        wsOut.Cells(1, 1).Value = "【AB対面分散(入替候補" & maxSwapRows & "件)】"
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

        ' 対面同士のヒット状況ヒートマップ(1&2番機～maxMachNum番機を機番配置順にゾーン表示)
        ' ※除外機番はスワップ候補・AB稼働率スコアの対象外だが、ヒートマップは実態を反映するためdictPairsAll(除外ロケーションのみ反映)を使う
        Dim zoneCrossHit() As Double, zoneTotalHit() As Double ' zoneTotalHitはそのゾーンの同面込みヒット数(対面比率(%)の分母)
        ReDim zoneCrossHit(1 To maxZoneNum)
        ReDim zoneTotalHit(1 To maxZoneNum)
        Dim pk As Variant, pkParts() As String
        For Each pk In dictPairsAll.Keys
            pkParts = Split(CStr(pk), ",")
            If dictItemZoneAll.Exists(pkParts(0)) Then
                Dim pZone As Long: pZone = CLng(dictItemZoneAll(pkParts(0)))
                If pZone >= 1 And pZone <= maxZoneNum Then
                    zoneTotalHit(pZone) = zoneTotalHit(pZone) + dictPairsAll(pk)
                    If dictCrossFaceAll.Exists(pk) Then
                        zoneCrossHit(pZone) = zoneCrossHit(pZone) + dictPairsAll(pk)
                    End If
                End If
            End If
        Next pk

        Dim maxCross As Double: maxCross = 0
        Dim zi As Long
        For zi = 1 To maxZoneNum
            If zoneCrossHit(zi) > maxCross Then maxCross = zoneCrossHit(zi)
        Next zi

        Dim heatTitleRow As Long: heatTitleRow = 4 + outCnt + 3
        Dim heatLabelRow As Long: heatLabelRow = heatTitleRow + 1
        Dim heatValueRow As Long: heatValueRow = heatTitleRow + 2
        Dim heatPctRow As Long: heatPctRow = heatTitleRow + 3

        ' 本表(A:M)と列幅を揃えると本表側の列幅が崩れるため、O列(15列目)以降の未使用領域にコンパクトな幅で配置する
        Const HEAT_COL_OFFSET As Long = 14 ' 15列目(O)から開始
        Dim heatFirstCol As Long: heatFirstCol = HEAT_COL_OFFSET + 1
        Dim heatLastCol As Long: heatLastCol = HEAT_COL_OFFSET + maxZoneNum

        wsOut.Range(wsOut.Cells(heatTitleRow, heatFirstCol), wsOut.Cells(heatTitleRow, heatLastCol)).Merge
        wsOut.Cells(heatTitleRow, heatFirstCol).Value = "【対面同士のヒット状況(ゾーン別ヒートマップ)】※濃いほど対面での同時出庫(同一編成での共起)が多い。下段はそのゾーン内の同時ヒットのうち対面が占める割合"
        wsOut.Cells(heatTitleRow, heatFirstCol).Font.Bold = True: wsOut.Cells(heatTitleRow, heatFirstCol).Font.Size = 12
        wsOut.Cells(heatTitleRow, heatFirstCol).HorizontalAlignment = xlLeft

        wsOut.Range(wsOut.Cells(heatLabelRow, heatFirstCol), wsOut.Cells(heatLabelRow, heatLastCol)).EntireColumn.ColumnWidth = 6

        For zi = 1 To maxZoneNum
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

        ' 7. 同号機分散(同号機内(対面を除く)ペアのみを対象にする。対面ヒットは対象外)
        Dim maxSameMachPairs As Long: maxSameMachPairs = dictPairs.Count - dictCrossFace.Count
        ' 対象ペアが無い/入替案が1件も出ない場合でも均衡化スコアが算出できるよう、既定値を変更前と同じにしておく
        Dim oddTotalSM As Double, evenTotalSM As Double
        oddTotalSM = oddTotalStart: evenTotalSM = evenTotalStart
        If maxSameMachPairs > 0 Then
            Dim smPairArr() As Variant
            ReDim smPairArr(1 To maxSameMachPairs, 1 To 6)
            Dim smCnt As Long: smCnt = 0
            Dim smKey As Variant
            For Each smKey In dictPairs.Keys
                If Not dictCrossFace.Exists(smKey) Then
                    Dim smItems() As String: smItems = Split(CStr(smKey), ",")
                    Dim smItemA As String: smItemA = smItems(0)
                    Dim smItemB As String: smItemB = smItems(1)
                    smCnt = smCnt + 1
                    Dim smCoCount As Long: smCoCount = dictPairs(smKey)
                    Dim smAnchor As String, smMover As String
                    If dictItemHit(smItemA) >= dictItemHit(smItemB) Then
                        smAnchor = smItemA: smMover = smItemB
                    Else
                        smAnchor = smItemB: smMover = smItemA
                    End If
                    smPairArr(smCnt, 1) = dictItemZone(smAnchor)
                    smPairArr(smCnt, 2) = smCoCount
                    smPairArr(smCnt, 3) = smAnchor
                    smPairArr(smCnt, 4) = dictItemLoc(smAnchor)
                    smPairArr(smCnt, 5) = smMover
                    smPairArr(smCnt, 6) = dictItemLoc(smMover)
                End If
            Next smKey

            ' 編成内共起回数(列2)の降順でソート
            Dim wsTempSM As Worksheet: Set wsTempSM = Sheets.Add
            wsTempSM.Columns("C:F").NumberFormat = "@" ' C列(起点品キー)も先頭ゼロ付きの機番を含むため、数値変換されないよう文字列扱いにする
            wsTempSM.Range("A1").Resize(smCnt, 6).Value = smPairArr
            wsTempSM.Sort.SortFields.Clear
            wsTempSM.Sort.SortFields.Add Key:=wsTempSM.Range("B1:B" & smCnt), Order:=xlDescending
            wsTempSM.Sort.SetRange wsTempSM.Range("A1:F" & smCnt)
            wsTempSM.Sort.Apply
            smPairArr = wsTempSM.Range("A1:F" & smCnt).Value
            wsTempSM.Delete

            ' 入替案の決定(このシート専用に、奇数/偶数の合計を独立して再計算する)
            Dim outArrSM() As Variant
            ReDim outArrSM(1 To smCnt, 1 To 13)
            Dim outCntSM As Long: outCntSM = 0
            Dim dictSwappedSM As Object: Set dictSwappedSM = CreateObject("Scripting.Dictionary")
            Dim dictZoneUsedCountSM As Object: Set dictZoneUsedCountSM = CreateObject("Scripting.Dictionary")

            ' このシート専用に、機番別ヒット数もmachHitStartから独立してコピーし直す(本表側のスワップの影響を受けない)
            Dim machHitLiveSM() As Double
            ReDim machHitLiveSM(1 To maxMachNum)
            Dim mIdx3 As Long
            For mIdx3 = 1 To maxMachNum
                machHitLiveSM(mIdx3) = machHitStart(mIdx3)
            Next mIdx3
            Dim targetedHitTotalSM As Double: targetedHitTotalSM = targetedHitStart

            Dim rsm As Long
            For rsm = 1 To smCnt
                If outCntSM >= maxSwapRows Then Exit For ' 入替候補(スコア順)は設定件数まで
                ' 候補ペアが多いと探索に時間がかかることがあるため、Excelが「応答なし」に見えないよう
                ' 一定回数ごとに制御をOSに戻す(処理自体は継続する)
                If rsm Mod 200 = 0 Then DoEvents

                Dim aItemSM As String: aItemSM = CStr(smPairArr(rsm, 3))
                Dim mItemSM As String: mItemSM = CStr(smPairArr(rsm, 5))

                If Not dictSwappedSM.Exists(aItemSM) And Not dictSwappedSM.Exists(mItemSM) Then
                    Dim anchorZoneSM As Integer: anchorZoneSM = dictItemZone(aItemSM)
                    Dim targetItemSM As String: targetItemSM = ""

                    Dim moverSideSM As Integer: moverSideSM = dictItemMach(mItemSM) Mod 2

                    ' パス1:目標構成比が入力されていれば、最も比率が不足している機番の候補をゾーン利用上限内で探す
                    If hasTargetRatioData Then
                        targetItemSM = FindBestUnderTargetCandidate(zoneItems, anchorZoneSM, aItemSM, mItemSM, dictSwappedSM, dictItemMach, dictTargetRatio, machHitLiveSM, targetedHitTotalSM, targetRatioSum, dictZoneUsedCountSM, True, MAX_PER_ZONE, dictItemCat, dictItemWt, dictItemVol, dictMachCatVol, catWeight, sizeWeight, weightWeightCoef)
                    End If
                    ' パス1':目標構成比が未入力なら、従来どおり希望サイド+ゾーン利用上限で探す
                    If targetItemSM = "" And Not hasTargetRatioData Then
                        Dim desiredSideSM As Integer
                        If Abs(oddTotalSM - evenTotalSM) <= 0.001 Then
                            desiredSideSM = -1
                        ElseIf (oddTotalSM > evenTotalSM And moverSideSM = 1) Or (evenTotalSM > oddTotalSM And moverSideSM = 0) Then
                            desiredSideSM = 1 - moverSideSM
                        Else
                            desiredSideSM = moverSideSM
                        End If
                        If desiredSideSM <> -1 Then
                            Dim zKeySM As Variant
                            For Each zKeySM In zoneItems.Keys
                                If CInt(zKeySM) <> anchorZoneSM Then
                                    Dim zoneUsedSM As Integer
                                    If dictZoneUsedCountSM.Exists(zKeySM) Then zoneUsedSM = dictZoneUsedCountSM(zKeySM) Else zoneUsedSM = 0
                                    If zoneUsedSM < MAX_PER_ZONE Then
                                        Dim candidateSM As Variant
                                        For Each candidateSM In zoneItems(zKeySM)
                                            Dim candStrSM As String: candStrSM = CStr(candidateSM)
                                            If candStrSM <> aItemSM And candStrSM <> mItemSM And Not dictSwappedSM.Exists(candStrSM) Then
                                                If dictItemMach(candStrSM) Mod 2 = desiredSideSM Then
                                                    targetItemSM = candStrSM
                                                    Exit For
                                                End If
                                            End If
                                        Next candidateSM
                                    End If
                                End If
                                If targetItemSM <> "" Then Exit For
                            Next zKeySM
                        End If
                    End If
                    ' パス2:ゾーン利用上限内で、比率・サイドを問わず最初に見つかった候補(在庫データがあればその中で一番属性が近い候補)
                    If targetItemSM = "" Then
                        targetItemSM = FindFirstCandidate(zoneItems, anchorZoneSM, aItemSM, mItemSM, dictSwappedSM, dictZoneUsedCountSM, True, MAX_PER_ZONE, dictItemMach, dictItemCat, dictItemWt, dictItemVol, dictMachCatVol, catWeight, sizeWeight, weightWeightCoef)
                    End If
                    ' パス3:制限なしで、最初に見つかった候補(最終手段。在庫データがあればその中で一番属性が近い候補)
                    If targetItemSM = "" Then
                        targetItemSM = FindFirstCandidate(zoneItems, anchorZoneSM, aItemSM, mItemSM, dictSwappedSM, dictZoneUsedCountSM, False, MAX_PER_ZONE, dictItemMach, dictItemCat, dictItemWt, dictItemVol, dictMachCatVol, catWeight, sizeWeight, weightWeightCoef)
                    End If

                    If targetItemSM <> "" Then
                        Dim usedZoneKeySM As String: usedZoneKeySM = CStr(dictItemZone(targetItemSM))
                        If dictZoneUsedCountSM.Exists(usedZoneKeySM) Then
                            dictZoneUsedCountSM(usedZoneKeySM) = dictZoneUsedCountSM(usedZoneKeySM) + 1
                        Else
                            dictZoneUsedCountSM.Add usedZoneKeySM, 1
                        End If

                        Dim moverHitsSM As Double: moverHitsSM = dictItemHit(mItemSM)
                        Dim targetHitsSM As Double: targetHitsSM = dictItemHit(targetItemSM)

                        Dim targetSideSM As Integer: targetSideSM = dictItemMach(targetItemSM) Mod 2
                        If moverSideSM <> targetSideSM Then
                            If moverSideSM = 1 Then
                                oddTotalSM = oddTotalSM - moverHitsSM + targetHitsSM
                                evenTotalSM = evenTotalSM - targetHitsSM + moverHitsSM
                            Else
                                evenTotalSM = evenTotalSM - moverHitsSM + targetHitsSM
                                oddTotalSM = oddTotalSM - targetHitsSM + moverHitsSM
                            End If
                        End If

                        Dim mMachSM As Long: mMachSM = dictItemMach(mItemSM)
                        Dim tMachSM As Long: tMachSM = dictItemMach(targetItemSM)
                        If dictTargetRatio.Exists(CStr(mMachSM)) Then targetedHitTotalSM = targetedHitTotalSM - moverHitsSM + targetHitsSM
                        If dictTargetRatio.Exists(CStr(tMachSM)) Then targetedHitTotalSM = targetedHitTotalSM - targetHitsSM + moverHitsSM
                        machHitLiveSM(mMachSM) = machHitLiveSM(mMachSM) - moverHitsSM + targetHitsSM
                        machHitLiveSM(tMachSM) = machHitLiveSM(tMachSM) - targetHitsSM + moverHitsSM

                        outCntSM = outCntSM + 1
                        outArrSM(outCntSM, 1) = anchorZoneSM
                        outArrSM(outCntSM, 2) = dictItemMach(aItemSM) & "号機内"
                        outArrSM(outCntSM, 3) = smPairArr(rsm, 2) ' 編成内共起回数
                        outArrSM(outCntSM, 4) = GetLocName3(dictLocName, dictItemMach(aItemSM), aItemSM)
                        outArrSM(outCntSM, 5) = GetLocCode3(dictLocCode, dictItemMach(aItemSM), aItemSM)
                        outArrSM(outCntSM, 6) = dictItemLoc(aItemSM)
                        outArrSM(outCntSM, 7) = GetLocName3(dictLocName, dictItemMach(mItemSM), mItemSM)
                        outArrSM(outCntSM, 8) = GetLocCode3(dictLocCode, dictItemMach(mItemSM), mItemSM)
                        outArrSM(outCntSM, 9) = dictItemLoc(mItemSM)
                        outArrSM(outCntSM, 10) = "⇔"
                        outArrSM(outCntSM, 11) = GetLocName3(dictLocName, dictItemMach(targetItemSM), targetItemSM)
                        outArrSM(outCntSM, 12) = GetLocCode3(dictLocCode, dictItemMach(targetItemSM), targetItemSM)
                        outArrSM(outCntSM, 13) = dictItemLoc(targetItemSM)

                        dictSwappedSM(mItemSM) = True
                        dictSwappedSM(targetItemSM) = True
                    End If
                End If
            Next rsm

            If outCntSM > 0 Then
                Dim wsOutSM As Worksheet
                On Error Resume Next
                Sheets("同号機分散").Delete
                On Error GoTo 0

                Dim wsPanelSM As Worksheet
                On Error Resume Next
                Set wsPanelSM = ThisWorkbook.Sheets("操作パネル")
                On Error GoTo 0
                If Not wsPanelSM Is Nothing Then
                    Set wsOutSM = ThisWorkbook.Sheets.Add(Before:=wsPanelSM)
                Else
                    Set wsOutSM = Sheets.Add
                End If
                wsOutSM.Name = "同号機分散"

                wsOutSM.Columns("F:F").NumberFormat = "@"
                wsOutSM.Columns("I:I").NumberFormat = "@"
                wsOutSM.Columns("M:M").NumberFormat = "@"

                wsOutSM.Range("A1:M1").Merge
                wsOutSM.Cells(1, 1).Value = "【同号機分散(同号機内・対面を除くペアのみ・入替候補" & maxSwapRows & "件)】"
                wsOutSM.Cells(1, 1).Font.Bold = True: wsOutSM.Cells(1, 1).Font.Size = 14
                wsOutSM.Cells(1, 1).HorizontalAlignment = xlLeft

                wsOutSM.Range("A2:M2").Merge
                wsOutSM.Cells(2, 1).Value = "同一号機内で同時ピッキングされている組み合わせを対象に、別ゾーンへ分散させる入替案です(対面(異なる号機)のペアは対象外)。奇数機番合計ヒット数: " & _
                    Format(oddTotalStart, "0") & " → " & Format(oddTotalSM, "0") & _
                    "　／　偶数機番合計ヒット数: " & Format(evenTotalStart, "0") & " → " & Format(evenTotalSM, "0")
                wsOutSM.Cells(2, 1).HorizontalAlignment = xlLeft

                wsOutSM.Range("A4:M4").Value = Array("ゾーン", "区分", "編成内共起回数", "【起点品】(動かさない)", "起点品コード", "起点ロケーション", "【交換品】(こちらを動かす)", "交換品コード", "交換元ロケーション", "交換方向", "【交換対象品】(別ゾーンの低頻度品)", "交換対象品コード", "交換先ロケーション")
                wsOutSM.Range("A5").Resize(outCntSM, 13).Value = outArrSM

                wsOutSM.Range("A4:M4").Interior.Color = RGB(230, 245, 225)
                wsOutSM.Range("A4:M4").Font.Bold = True
                wsOutSM.Columns("A:M").AutoFit
            End If
        End If

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
            Dim machHit() As Double, machTarget() As Double
            ReDim machHit(1 To maxMachNum)
            ReDim machTarget(1 To maxMachNum)
            Dim hk As Variant
            For Each hk In dictItemHit.Keys
                Dim hm As Long: hm = dictItemMach(hk)
                If hm >= 1 And hm <= maxMachNum Then machHit(hm) = machHit(hm) + dictItemHit(hk)
            Next hk

            Dim rr3 As Long, abLabel3 As String, mNum3 As Integer
            For rr3 = 3 To 2 + maxMachNum ' AB01(1番機)～maxMachNum番機に対応する行
                abLabel3 = Trim(CStr(wsRatio3.Cells(rr3, 1).Value))
                If abLabel3 Like "AB##" Then
                    mNum3 = CInt(Mid(abLabel3, 3, 2))
                    If mNum3 >= 1 And mNum3 <= maxMachNum Then machTarget(mNum3) = Val(wsRatio3.Cells(rr3, 5).Value)
                End If
            Next rr3

            ' 1号機～maxMachNum号機を対象に、設定シートの除外機番だけを動的に除いて正規化して比較する
            Dim hitTotal As Double, targetTotal As Double, mIdx As Long
            hitTotal = 0: targetTotal = 0
            For mIdx = 1 To maxMachNum
                If Not dictExcludedMach.Exists(CStr(mIdx)) Then
                    hitTotal = hitTotal + machHit(mIdx)
                    targetTotal = targetTotal + machTarget(mIdx)
                End If
            Next mIdx

            If hitTotal <= 0 Then
                abRatioScoreNote = "実績データが除外設定によりすべて対象外のため、AB稼働率スコアは算出されていません"
            ElseIf targetTotal <= 0 Then
                abRatioScoreNote = "「" & ratioSheetName & "」シートにAB01～AB" & Format(maxMachNum, "00") & "の目標比率(A列ラベル・E列数値、3～" & (2 + maxMachNum) & "行目)が見つからないため、AB稼働率スコアは算出されていません"
            Else
                Dim sumAbsDiff As Double: sumAbsDiff = 0
                For mIdx = 1 To maxMachNum
                    If Not dictExcludedMach.Exists(CStr(mIdx)) Then
                        sumAbsDiff = sumAbsDiff + Abs((machHit(mIdx) / hitTotal) - (machTarget(mIdx) / targetTotal))
                    End If
                Next mIdx
                ' 差の合計(sumAbsDiff)が0.6(理論上の最大2.0の約1/3)以上で0点、0で100点、その間は線形
                abRatioScore = Application.WorksheetFunction.Max(0, 100 * (1 - sumAbsDiff / 0.6))
            End If
        End If

        ' KPI記録:AB得意先スコア(理論値:全体の回数上位abSlotCount件(AB間口数)の回数比率／実績値:AB番機の実回数比率)
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
                Dim keyMach As Long: keyMach = CLng(Mid(CStr(allKey), 2, 3))
                If keyMach >= 1 And keyMach <= maxMachNum Then abActualTotal = abActualTotal + hitVal
                ar = ar + 1
            Next allKey
            Dim allN As Long: allN = ar - 1

            Dim topSum As Double
            If allN >= abSlotCount Then
                ' 回数の多い順に並べ替えて、ちょうど上位abSlotCount件(AB間口数)だけ合計する
                ' (LARGE+SUMIF(">=")式だと同着タイのロケーションが全部含まれてしまい、間口数を超えて合計されることがあるため補正)
                wsTempAll.Range("A1:A" & allN).Sort Key1:=wsTempAll.Range("A1"), Order1:=xlDescending, Header:=xlNo
                topSum = Application.WorksheetFunction.Sum(wsTempAll.Range("A1:A" & abSlotCount))
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
        ' Module7が無いブックでもコンパイルエラーにならないよう、Application.Runで実行時に解決する
        Application.Run "Module7.LogFormationScore", oddTotalStart, evenTotalStart, oddTotal, evenTotal, crossFaceScore, abRatioScore, abOccupancyScore, abTheoreticalRatioOut, abActualRatioOut, reportDate
        On Error GoTo 0

        Dim completeMsg As String
        completeMsg = "「AB対面分散」の作成が完了しました。(" & selectedFileCount & "ファイル読込／" & outCnt & "件の入替案)" & vbCrLf & _
            "左右機番の差: " & Format(Abs(oddTotalStart - evenTotalStart), "0") & " → " & Format(Abs(oddTotal - evenTotal), "0")
        If useCache Then completeMsg = completeMsg & vbCrLf & "(前回読み込んだ実績データを再利用しました)"
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

' 交換先候補の中から、目標構成比(設定シート「■機番別目標構成比」)に対して最も不足している
' (現在の実績比率と目標比率の差=deviationが最小=マイナス方向に最も大きい)機番の候補を探す。
' 実績比率・目標比率とも「目標が設定されている機番だけ」の合計(targetedHitTotal/targetRatioSum)で
' それぞれ正規化してから比較するため、設定シートの目標構成比の合計が100%になっていなくても、
' 機番どうしの相対バランスとして機能する(Cバラ等AB以外への出荷分の影響も受けない)。
' respectZoneLimit=Trueならゾーン利用上限(maxPerZone)を満たすゾーンのみを対象にする。
' 該当候補が無ければ空文字を返す(呼び出し側でパス2以降にフォールバックする)
Function FindBestUnderTargetCandidate(zoneItems As Object, anchorZone As Integer, excludeItem1 As String, excludeItem2 As String, dictSwapped As Object, dictItemMach As Object, dictTargetRatio As Object, machHitLive() As Double, ByVal targetedHitTotal As Double, ByVal targetRatioSum As Double, dictZoneUsedCount As Object, ByVal respectZoneLimit As Boolean, ByVal maxPerZone As Integer, dictItemCat As Object, dictItemWt As Object, dictItemVol As Object, dictMachCatVol As Object, ByVal catWeight As Double, ByVal sizeWeight As Double, ByVal weightWeightCoef As Double) As String
    Dim bestDev As Double: bestDev = 2# ' 比率の差の理論上の最大値(-1～1)より大きい値で初期化
    Dim bestCand As String: bestCand = ""

    ' moverアイテム(excludeItem2)の属性は候補走査の前に1回だけ解決しておく(候補ごとに辞書引きし直すと、
    ' 候補数が多いときに無駄な処理が積み重なって動作が重くなるため)
    Dim moverHasCat As Boolean, moverCat As String
    Dim moverHasWt As Boolean, moverWt As Double
    Dim moverHasVol As Boolean, moverVol As Double
    Call ResolveMoverAttr(excludeItem2, dictItemCat, dictItemWt, dictItemVol, moverHasCat, moverCat, moverHasWt, moverWt, moverHasVol, moverVol)

    Dim zKey As Variant
    For Each zKey In zoneItems.Keys
        If CInt(zKey) <> anchorZone Then
            Dim zoneUsed As Integer
            If dictZoneUsedCount.Exists(zKey) Then zoneUsed = dictZoneUsedCount(zKey) Else zoneUsed = 0
            If Not respectZoneLimit Or zoneUsed < maxPerZone Then
                Dim candidate As Variant
                For Each candidate In zoneItems(zKey)
                    Dim candStr As String: candStr = CStr(candidate)
                    If candStr <> excludeItem1 And candStr <> excludeItem2 And Not dictSwapped.Exists(candStr) Then
                        Dim candMach As Long: candMach = dictItemMach(candStr)
                        Dim candMachKey As String: candMachKey = CStr(candMach)
                        If dictTargetRatio.Exists(candMachKey) Then
                            Dim dev As Double: dev = (machHitLive(candMach) / targetedHitTotal) - (dictTargetRatio(candMachKey) / targetRatioSum)
                            ' 在庫データが読み込まれていれば、比率の差に「同カテゴリー集中度・サイズ差・重量差」の
                            ' ソフトなペナルティを加味する(未読込なら常に0で従来と同じ結果になる)
                            Dim combinedScore As Double
                            combinedScore = dev + ComputeAttrPenalty(candStr, dictItemMach, dictItemVol, dictItemWt, dictMachCatVol, moverHasCat, moverCat, moverHasWt, moverWt, moverHasVol, moverVol, catWeight, sizeWeight, weightWeightCoef)
                            If combinedScore < bestDev Then
                                bestDev = combinedScore
                                bestCand = candStr
                            End If
                        End If
                    End If
                Next candidate
            End If
        End If
    Next zKey
    FindBestUnderTargetCandidate = bestCand
End Function

' 交換先候補の中から、条件を満たす最初の候補を返す(目標比率を考慮しない従来どおりのフォールバック探索)
Function FindFirstCandidate(zoneItems As Object, anchorZone As Integer, excludeItem1 As String, excludeItem2 As String, dictSwapped As Object, dictZoneUsedCount As Object, ByVal respectZoneLimit As Boolean, ByVal maxPerZone As Integer, dictItemMach As Object, dictItemCat As Object, dictItemWt As Object, dictItemVol As Object, dictMachCatVol As Object, ByVal catWeight As Double, ByVal sizeWeight As Double, ByVal weightWeightCoef As Double) As String
    ' 在庫データが読み込まれていなければ、従来どおり最初に見つかった候補をそのまま返す(挙動を変えない)
    Dim hasAttrData As Boolean: hasAttrData = (dictItemCat.Count > 0 Or dictItemVol.Count > 0 Or dictItemWt.Count > 0)
    Dim bestPenalty As Double: bestPenalty = -1
    Dim bestCand As String: bestCand = ""

    ' moverアイテム(excludeItem2)の属性は候補走査の前に1回だけ解決しておく(候補ごとの辞書引きを減らして高速化)
    Dim moverHasCat As Boolean, moverCat As String
    Dim moverHasWt As Boolean, moverWt As Double
    Dim moverHasVol As Boolean, moverVol As Double
    If hasAttrData Then
        Call ResolveMoverAttr(excludeItem2, dictItemCat, dictItemWt, dictItemVol, moverHasCat, moverCat, moverHasWt, moverWt, moverHasVol, moverVol)
    End If

    Dim zKey As Variant
    For Each zKey In zoneItems.Keys
        If CInt(zKey) <> anchorZone Then
            Dim zoneUsed As Integer
            If dictZoneUsedCount.Exists(zKey) Then zoneUsed = dictZoneUsedCount(zKey) Else zoneUsed = 0
            If Not respectZoneLimit Or zoneUsed < maxPerZone Then
                Dim candidate As Variant
                For Each candidate In zoneItems(zKey)
                    Dim candStr As String: candStr = CStr(candidate)
                    If candStr <> excludeItem1 And candStr <> excludeItem2 And Not dictSwapped.Exists(candStr) Then
                        If Not hasAttrData Then
                            FindFirstCandidate = candStr
                            Exit Function
                        End If
                        Dim candPenalty As Double
                        candPenalty = ComputeAttrPenalty(candStr, dictItemMach, dictItemVol, dictItemWt, dictMachCatVol, moverHasCat, moverCat, moverHasWt, moverWt, moverHasVol, moverVol, catWeight, sizeWeight, weightWeightCoef)
                        If bestPenalty < 0 Or candPenalty < bestPenalty Then
                            bestPenalty = candPenalty
                            bestCand = candStr
                        End If
                    End If
                Next candidate
            End If
        End If
    Next zKey
    FindFirstCandidate = bestCand
End Function

' 交換候補(candStr)にmoverアイテム(mItem)を入替配置した場合の「属人的判断」を数値化したペナルティ(小さいほど良い)。
' ①入替先候補が属する機番に、moverと同じ大分類コードの品が既にどれだけあるか(多いほど加点=同時ピッキング集中リスク)。
' ②候補とmoverのサイズ(体積)・重量の差(対数比。値が大きいほど物理的な入替えにくさが増す)。
' 在庫データが未読込(各dictが空)の品目は該当項目を単純にスキップする(0加点のまま)
Function ComputeAttrPenalty(candStr As String, dictItemMach As Object, dictItemVol As Object, dictItemWt As Object, dictMachCatVol As Object, ByVal moverHasCat As Boolean, ByVal moverCat As String, ByVal moverHasWt As Boolean, ByVal moverWt As Double, ByVal moverHasVol As Boolean, ByVal moverVol As Double, ByVal catWeight As Double, ByVal sizeWeight As Double, ByVal weightWeightCoef As Double) As Double
    Dim penalty As Double: penalty = 0
    If moverHasCat Then
        Dim tallyKey As String: tallyKey = CStr(dictItemMach(candStr)) & "|" & moverCat
        If dictMachCatVol.Exists(tallyKey) Then
            Dim simCount As Long: simCount = 0
            If moverHasVol And moverVol > 0 Then
                Dim nKey2 As String: nKey2 = tallyKey & "|N"
                If dictMachCatVol.Exists(nKey2) Then simCount = simCount + dictMachCatVol(nKey2)
                Dim mb As Long: mb = Int(Log(moverVol) / SIZE_SIMILAR_RATIO)
                Dim bi As Long
                For bi = mb - 1 To mb + 1
                    Dim bKey2 As String: bKey2 = tallyKey & "|B" & bi
                    If dictMachCatVol.Exists(bKey2) Then simCount = simCount + dictMachCatVol(bKey2)
                Next bi
            Else
                simCount = dictMachCatVol(tallyKey)
            End If
            penalty = penalty + catWeight * simCount
        End If
    End If
    If moverHasVol And dictItemVol.Exists(candStr) Then
        If dictItemVol(candStr) > 0 And moverVol > 0 Then
            penalty = penalty + sizeWeight * Abs(Log(dictItemVol(candStr) / moverVol))
        End If
    End If
    If moverHasWt And dictItemWt.Exists(candStr) Then
        If dictItemWt(candStr) > 0 And moverWt > 0 Then
            penalty = penalty + weightWeightCoef * Abs(Log(dictItemWt(candStr) / moverWt))
        End If
    End If
    ComputeAttrPenalty = penalty
End Function

' candStrアイテムの属性(カテゴリー・重量・体積)を1回だけ解決する(候補走査ループの前に1回だけ呼ぶ想定)。
' ComputeAttrPenaltyを候補ごとに呼ぶたびにmoverの辞書引きをやり直すと、候補数が多い号機間バランスなどで
' 無駄な処理が積み重なり動作が重くなるため、事前に解決した値を使い回す形にしている
Sub ResolveMoverAttr(mItem As String, dictItemCat As Object, dictItemWt As Object, dictItemVol As Object, ByRef moverHasCat As Boolean, ByRef moverCat As String, ByRef moverHasWt As Boolean, ByRef moverWt As Double, ByRef moverHasVol As Boolean, ByRef moverVol As Double)
    moverHasCat = dictItemCat.Exists(mItem)
    If moverHasCat Then moverCat = dictItemCat(mItem)
    moverHasWt = dictItemWt.Exists(mItem)
    If moverHasWt Then moverWt = dictItemWt(mItem)
    moverHasVol = dictItemVol.Exists(mItem)
    If moverHasVol Then moverVol = dictItemVol(mItem)
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
' 品名マスタ・ロケーションマスタ(任意)
' ----------------------------------------------------

' 品コード⇔品名の対応(品名マスタ)、機番・段・列⇔品コードの対応(ロケーションマスタ)を、
' CFシートとは別に外部ファイルから読み込めるようにする。ファイル選択ダイアログでキャンセルすれば、
' 何もせずCFシートの内容だけで従来通り動作する。2種類のファイルをまとめて選択でき、
' 先頭行が"B"で始まるかどうかでどちらのファイルかを自動判別する。
'   ロケーションマスタ:1行目"B"+日付、以降"E"+機番(2)+段(2)+列(2)+品コード(6)+…(固定長)
'   品名マスタ:1行目から品コード(7桁)+…+品名(半角カナ、55～72文字目)+…(固定長128バイト)
Sub LoadItemMasterFilesIfSelected(dictLocCode As Object, dictLocName As Object)
    Dim fd2 As Office.FileDialog
    Set fd2 = Application.FileDialog(msoFileDialogFilePicker)
    With fd2
        .Title = "品名マスタ(S01)・ロケーションマスタ(S74)を選択(任意・複数選択可。使わない場合はキャンセルでCFシートのみ使用)"
        .Filters.Clear
        .Filters.Add "すべてのファイル", "*.*"
        .AllowMultiSelect = True
        If .Show = False Then Exit Sub
    End With

    Dim dictItemNameByCode As Object: Set dictItemNameByCode = CreateObject("Scripting.Dictionary") ' 品コード(7桁文字列)→品名

    Dim fIdx2 As Long, filePath2 As String, fileNo2 As Integer, firstLine As String, textLine2 As String
    For fIdx2 = 1 To fd2.SelectedItems.Count
        filePath2 = fd2.SelectedItems(fIdx2)
        fileNo2 = FreeFile
        Open filePath2 For Input As #fileNo2
        If EOF(fileNo2) Then
            Close #fileNo2
        Else
            Line Input #fileNo2, firstLine

            If Left(firstLine, 1) = "B" Then
                ' ロケーションマスタ:E行の2～7文字目=機番段列(6桁)、8～15文字目=品コード(8桁固定域。
                ' 6桁品コードは末尾2文字が空白埋め、8桁品コード(28xxxxxx等)はそのまま埋まる)
                Do While Not EOF(fileNo2)
                    Line Input #fileNo2, textLine2
                    If Left(textLine2, 1) = "E" And Len(textLine2) >= 15 Then
                        Dim locStr As String: locStr = Mid(textLine2, 2, 6)
                        Dim itemCodeStr As String: itemCodeStr = Trim(Mid(textLine2, 8, 8))
                        If IsNumeric(locStr) And itemCodeStr <> "" And IsNumeric(itemCodeStr) Then
                            Dim mLocCode As String
                            mLocCode = CStr(CLng(Mid(locStr, 1, 2)) * 10000& + CLng(Mid(locStr, 3, 2)) * 100& + CLng(Mid(locStr, 5, 2)))
                            ' 先頭ゼロが意味を持つ品コードのため、数値変換せず文字列のまま保持する
                            dictLocCode(mLocCode) = itemCodeStr
                        End If
                    End If
                Loop
            Else
                ' 品名マスタ:1～7文字目=品コード(7桁)、55～72文字目=品名(半角カナ)
                Dim nameLine As String: nameLine = firstLine
                Do
                    If Len(nameLine) >= 72 And IsNumeric(Left(nameLine, 7)) Then
                        dictItemNameByCode(Left(nameLine, 7)) = Trim(Mid(nameLine, 55, 18))
                    End If
                    If EOF(fileNo2) Then Exit Do
                    Line Input #fileNo2, nameLine
                Loop
            End If
            Close #fileNo2
        End If
    Next fIdx2

    ' 品名マスタが読み込めた場合、ロケーション→品コードの対応(CF・ロケーションマスタ双方)を使って品名を上書きする。
    ' dictLocCodeの値はCF由来なら数値、ロケーションマスタ由来なら文字列と型が揃っていないため、
    ' 元の桁数のまま/先頭ゼロを1つ追加/先頭ゼロを1つ除去、の3通りで品名マスタと照合する
    If dictItemNameByCode.Count > 0 Then
        Dim locKeyIter As Variant
        For Each locKeyIter In dictLocCode.Keys
            Dim rawCode As String: rawCode = Trim(CStr(dictLocCode(locKeyIter)))
            If rawCode <> "" Then
                If dictItemNameByCode.Exists(rawCode) Then
                    dictLocName(locKeyIter) = dictItemNameByCode(rawCode)
                ElseIf dictItemNameByCode.Exists("0" & rawCode) Then
                    dictLocName(locKeyIter) = dictItemNameByCode("0" & rawCode)
                ElseIf Left(rawCode, 1) = "0" And dictItemNameByCode.Exists(Mid(rawCode, 2)) Then
                    dictLocName(locKeyIter) = dictItemNameByCode(Mid(rawCode, 2))
                End If
            End If
        Next locKeyIter
    End If
End Sub

' 在庫データ(在庫状況ダウンロード・WF021L1形式のCSV、任意)を読み込む。
' このファイルには品コードごとの大分類コード・実測(無ければ参考)の縦横高・重量が含まれており、
' これまで人が目視で判断していた「同カテゴリーを集中させない」「サイズ・重量が近いものを選ぶ」を
' 入替候補選定のソフトなスコアに反映するために使う。ファイル選択でキャンセルすれば何もせず、
' 従来どおり目標構成比・奇数偶数バランスだけで入替候補を選ぶ(スコアへの影響が無いだけで、機能は変わらない)。
' 1～3行目はタイトル・空行・見出し行、4行目以降が空行またはデータ行(先頭の拠点コード等が数値のみ)。
' D列(4列目)=商品コード、103列目=大分類コード、68～71列目=参考の縦/横/高/重量、
' 84～87列目=実測の縦/横/高/重量(実測が無ければ参考を使う)
' 在庫データ(在庫状況ダウンロード・WF021L1形式のCSV)を取り込み、「在庫データ」シートに保存する。
' 予測データ取込(Module8)と同様に一度取り込めば済み、以降はマクロ実行のたびにファイルを選び直す必要がない
' (「在庫データ」シートが残っている限り、AB対面分散・号機間バランスはそこから読み込む)。
' 既存の「在庫データ」シートは削除してから作り直すため、再取込すると内容が更新される
Sub ImportItemAttributeMaster()
    Call EnsureItemAttributeImportButton

    Dim fd5 As Office.FileDialog
    Set fd5 = Application.FileDialog(msoFileDialogFilePicker)
    With fd5
        .Title = "在庫データ(在庫状況ダウンロード・WF021L1形式のCSV)を選択"
        .Filters.Clear
        .Filters.Add "すべてのファイル", "*.*"
        .AllowMultiSelect = False
        If .Show = False Then Exit Sub
    End With

    Dim filePath5 As String: filePath5 = fd5.SelectedItems(1)
    Dim fileDate5 As Date: fileDate5 = FileDateTime(filePath5)

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    Dim fileNo5 As Integer: fileNo5 = FreeFile
    Dim textLine5 As String
    Dim outRows As Collection: Set outRows = New Collection

    Open filePath5 For Input As #fileNo5
    Do While Not EOF(fileNo5)
        Line Input #fileNo5, textLine5
        Dim cols5() As String: cols5 = Split(textLine5, ",")
        If UBound(cols5) >= 104 Then
            ' 一部のCSV書き出しでは、コード列などが ="00150" のように=と引用符で囲まれる
            ' (Excelが先頭ゼロを落とさないようにする書式)。そのままだとIsNumericが常にFalseに
            ' なってしまうため、各列を読むたびにこの記法を検出して中身だけを取り出す
            Dim rawCode5 As String: rawCode5 = StripCsvQuote5(Trim(cols5(3))) ' 4列目:商品コード
            If rawCode5 <> "" And IsNumeric(rawCode5) Then
                Dim codeKey5 As String: codeKey5 = CStr(CLng(rawCode5))
                Dim catL5 As String: catL5 = StripCsvQuote5(Trim(cols5(102))) ' 大分類コード
                Dim catM5 As String: catM5 = StripCsvQuote5(Trim(cols5(103))) ' 中分類コード
                Dim catS5 As String: catS5 = StripCsvQuote5(Trim(cols5(104))) ' 小分類コード

                ' 重量:実測梱重量(95列目)を優先、無ければ参考梱重量(83列目)を使う
                ' (ラックの1ロケーションには通常梱単位で格納されるため、商品単位ではなく梱単位の寸法・重量を使う)
                Dim wStr5 As String: wStr5 = StripCsvQuote5(Trim(cols5(94)))
                If Not (IsNumeric(wStr5) And CDbl(wStr5) > 0) Then wStr5 = StripCsvQuote5(Trim(cols5(82)))

                ' サイズ(縦横高):実測梱寸法(92～94列目)を優先、無ければ参考梱(80～82列目)を使う
                Dim dStr5 As String, wdStr5 As String, hStr5 As String
                dStr5 = StripCsvQuote5(Trim(cols5(91))): wdStr5 = StripCsvQuote5(Trim(cols5(92))): hStr5 = StripCsvQuote5(Trim(cols5(93)))
                If Not (IsNumeric(dStr5) And IsNumeric(wdStr5) And IsNumeric(hStr5) And CDbl(dStr5) > 0 And CDbl(wdStr5) > 0) Then
                    dStr5 = StripCsvQuote5(Trim(cols5(79))): wdStr5 = StripCsvQuote5(Trim(cols5(80))): hStr5 = StripCsvQuote5(Trim(cols5(81)))
                End If

                ' 在庫数・発売期間(スコア計算には使わないが、参考情報としてシートに保存しておく)
                Dim stockKon5 As String: stockKon5 = StripCsvQuote5(Trim(cols5(13)))  ' 14列目:通常在庫(梱)
                Dim stockBara5 As String: stockBara5 = StripCsvQuote5(Trim(cols5(14))) ' 15列目:通常在庫(バラ)
                Dim saleFrom5 As String: saleFrom5 = StripCsvQuote5(Trim(cols5(24)))  ' 25列目:発売開始年月日
                Dim saleTo5 As String: saleTo5 = StripCsvQuote5(Trim(cols5(25)))    ' 26列目:発売終了年月日

                Dim rowArr5(1 To 12) As Variant
                rowArr5(1) = codeKey5
                rowArr5(2) = catL5
                rowArr5(3) = catM5
                rowArr5(4) = catS5
                rowArr5(5) = IIf(IsNumeric(dStr5), CDbl(dStr5), 0)
                rowArr5(6) = IIf(IsNumeric(wdStr5), CDbl(wdStr5), 0)
                rowArr5(7) = IIf(IsNumeric(hStr5), CDbl(hStr5), 0)
                rowArr5(8) = IIf(IsNumeric(wStr5), CDbl(wStr5), 0)
                rowArr5(9) = IIf(IsNumeric(stockKon5), CDbl(stockKon5), 0)
                rowArr5(10) = IIf(IsNumeric(stockBara5), CDbl(stockBara5), 0)
                rowArr5(11) = saleFrom5
                rowArr5(12) = saleTo5
                outRows.Add rowArr5
            End If
        End If
    Loop
    Close #fileNo5

    If outRows.Count = 0 Then
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        MsgBox "商品コードを含むデータ行が見つかりませんでした。ファイルの内容を確認してください。", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Sheets("在庫データ").Delete
    On Error GoTo 0

    Dim wsAttr As Worksheet
    Dim wsPanel5 As Worksheet
    On Error Resume Next
    Set wsPanel5 = ThisWorkbook.Sheets("操作パネル")
    On Error GoTo 0
    If Not wsPanel5 Is Nothing Then
        Set wsAttr = ThisWorkbook.Sheets.Add(Before:=wsPanel5)
    Else
        Set wsAttr = Sheets.Add
    End If
    wsAttr.Name = "在庫データ"

    wsAttr.Columns("A:A").NumberFormat = "@" ' 品コードは先頭ゼロ落ち防止のため文字列扱いにする
    wsAttr.Columns("K:L").NumberFormat = "@" ' 発売開始・終了年月日(YYYYMMDD)は日付誤変換防止のため文字列扱いにする

    wsAttr.Range("A1:L1").Merge
    wsAttr.Range("A1").Value = "【在庫データ取込】ファイル: " & Dir(filePath5) & _
        "　／　ファイル更新日時: " & Format(fileDate5, "yyyy/mm/dd hh:mm") & _
        "　／　取込日時: " & Format(Now, "yyyy/mm/dd hh:mm")
    wsAttr.Range("A1").Font.Bold = True: wsAttr.Range("A1").Font.Size = 12
    wsAttr.Range("A1").HorizontalAlignment = xlLeft

    Const HEADER_ROW5 As Long = 3
    wsAttr.Range("A3:L3").Value = Array("品コード", "大分類コード", "中分類コード", "小分類コード", "梱-縦", "梱-横", "梱-高", "梱-重量(kg)", "通常在庫(梱)", "通常在庫(バラ)", "発売開始年月日", "発売終了年月日")
    wsAttr.Range("A3:L3").Interior.Color = RGB(220, 230, 255)
    wsAttr.Range("A3:L3").Font.Bold = True

    Dim outArr5() As Variant
    ReDim outArr5(1 To outRows.Count, 1 To 12)
    Dim ri5 As Long: ri5 = 0
    Dim rv5 As Variant
    For Each rv5 In outRows
        ri5 = ri5 + 1
        Dim c5 As Long
        For c5 = 1 To 12
            outArr5(ri5, c5) = rv5(c5)
        Next c5
    Next rv5
    wsAttr.Range(wsAttr.Cells(HEADER_ROW5 + 1, 1), wsAttr.Cells(HEADER_ROW5 + outRows.Count, 12)).Value = outArr5

    wsAttr.Range("A3:L3").AutoFilter
    wsAttr.Columns("A:L").AutoFit
    wsAttr.Rows(1).RowHeight = 20

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "「在庫データ」シートを更新しました。(" & outRows.Count & "件取込)" & vbCrLf & _
        "以降、AB対面分散・号機間バランスはこのシートのデータを使います(ファイル選択は不要です)。", vbInformation
End Sub

' ="00150" のようなExcel形式のCSVクォート(=と引用符で先頭ゼロなどを保護する書式)を検出し、
' 該当すれば中身だけを取り出す。該当しない(通常の値)場合はそのまま返す
Function StripCsvQuote5(ByVal s As String) As String
    If Len(s) >= 3 And Left(s, 2) = "=" & Chr(34) And Right(s, 1) = Chr(34) Then
        StripCsvQuote5 = Mid(s, 3, Len(s) - 3)
    Else
        StripCsvQuote5 = s
    End If
End Function

' 「在庫データ」シート(ImportItemAttributeMasterで取込済み)から、カテゴリー・重量・体積を読み込む。
' シートが無い場合、または「設定」シートのチェックボックス(L13)がオフの場合は何もしない(辞書は空のまま=
' 従来どおりの動作になる)。ファイル選択ダイアログは出さない(取込はImportItemAttributeMasterの役目)
Sub LoadItemAttributeMasterFromSheet(dictItemCategory As Object, dictItemWeightMaster As Object, dictItemVolumeMaster As Object)
    Dim wsSetChk As Worksheet
    On Error Resume Next
    Set wsSetChk = ThisWorkbook.Sheets("設定")
    On Error GoTo 0
    If Not wsSetChk Is Nothing Then
        If wsSetChk.Range("L13").Value = False Then Exit Sub
    End If

    Dim wsAttr As Worksheet
    On Error Resume Next
    Set wsAttr = ThisWorkbook.Sheets("在庫データ")
    On Error GoTo 0
    If Not wsAttr Is Nothing Then

    ' 「カテゴリー粒度」(L12)に応じて、大分類(B列)/中分類(C列)/小分類(D列)のどれを使うかを決める
    Dim catCol5 As Long: catCol5 = 2
    If Not wsSetChk Is Nothing Then
        Select Case Trim(CStr(wsSetChk.Range("L12").Value))
            Case "中分類": catCol5 = 3
            Case "小分類": catCol5 = 4
            Case Else: catCol5 = 2
        End Select
    End If

    Const HEADER_ROW6 As Long = 3
    Dim lastRow6 As Long: lastRow6 = wsAttr.Cells(wsAttr.Rows.Count, 1).End(xlUp).Row
    If lastRow6 > HEADER_ROW6 Then
        ' セルを1行ずつ読むと遅いため、範囲を配列に一括で読み込んでからループする
        Dim attrArr6 As Variant
        attrArr6 = wsAttr.Range(wsAttr.Cells(HEADER_ROW6 + 1, 1), wsAttr.Cells(lastRow6, 8)).Value
        Dim r6 As Long
        For r6 = 1 To UBound(attrArr6, 1)
            Dim codeKey6 As String: codeKey6 = Trim(CStr(attrArr6(r6, 1)))
            If codeKey6 <> "" And IsNumeric(codeKey6) Then
                codeKey6 = CStr(CLng(codeKey6))
                Dim catVal6 As String: catVal6 = Trim(CStr(attrArr6(r6, catCol5)))
                If catVal6 <> "" And Not dictItemCategory.Exists(codeKey6) Then dictItemCategory.Add codeKey6, catVal6

                Dim d6 As Double: d6 = Val(attrArr6(r6, 5))
                Dim w6 As Double: w6 = Val(attrArr6(r6, 6))
                Dim h6 As Double: h6 = Val(attrArr6(r6, 7))
                Dim wt6 As Double: wt6 = Val(attrArr6(r6, 8))

                If wt6 > 0 And Not dictItemWeightMaster.Exists(codeKey6) Then dictItemWeightMaster.Add codeKey6, wt6
                If d6 > 0 And w6 > 0 And h6 > 0 And Not dictItemVolumeMaster.Exists(codeKey6) Then
                    dictItemVolumeMaster.Add codeKey6, d6 * w6 * h6
                End If
            End If
        Next r6
    End If
    End If

    ' 予測データ自体の"カテゴリ"列(ブランド単位の分類など、大分類コードより判別的なもの)を優先的に上書き
    Call OverlayCategoryFromPredictionData(dictItemCategory)
End Sub

' 「予測データ」シート(Module8で取込済)に、品コードと同様の「カテゴリ(ブランド名)」列があれば上書き。
' WF021L1の大分類コードは01/07/空白の2～3種類しかなく判別に使えないため、より実感に近い
' ブランド名を優先的に同カテゴリーとして使う(存在する場合のみ置き換える)。
' 「予測データ」シートまたは「カテゴリ」列が無い場合は何もしない
Sub OverlayCategoryFromPredictionData(dictItemCategory As Object)
    Dim wsPred As Worksheet
    On Error Resume Next
    Set wsPred = ThisWorkbook.Sheets("予測データ")
    On Error GoTo 0
    If wsPred Is Nothing Then Exit Sub

    Const HEADER_ROW7 As Long = 3
    Dim lastRow7 As Long: lastRow7 = wsPred.Cells(wsPred.Rows.Count, 1).End(xlUp).Row
    Dim lastCol7 As Long: lastCol7 = wsPred.Cells(HEADER_ROW7, wsPred.Columns.Count).End(xlToLeft).Column
    If lastRow7 <= HEADER_ROW7 Then Exit Sub

    Dim headerArr7 As Variant
    headerArr7 = wsPred.Range(wsPred.Cells(HEADER_ROW7, 1), wsPred.Cells(HEADER_ROW7, lastCol7)).Value
    Dim codeColIdx7 As Long: codeColIdx7 = -1
    Dim catColIdx7 As Long: catColIdx7 = -1
    Dim hc7 As Long
    For hc7 = 1 To lastCol7
        Dim hName7 As String: hName7 = Trim(CStr(headerArr7(1, hc7)))
        If hName7 = "品名コード" Then codeColIdx7 = hc7
        If hName7 = "カテゴリ" Then catColIdx7 = hc7
    Next hc7
    If codeColIdx7 = -1 Or catColIdx7 = -1 Then Exit Sub

    ' セルを1行ずつ読むと遅いため、範囲を配列に一括で読み込んでからループする
    Dim predArr7 As Variant
    predArr7 = wsPred.Range(wsPred.Cells(HEADER_ROW7 + 1, 1), wsPred.Cells(lastRow7, lastCol7)).Value
    Dim r7 As Long
    For r7 = 1 To UBound(predArr7, 1)
        Dim codeKey7 As String: codeKey7 = Trim(CStr(predArr7(r7, codeColIdx7)))
        If codeKey7 <> "" And IsNumeric(codeKey7) Then
            codeKey7 = CStr(CLng(codeKey7))
            Dim catVal7 As String: catVal7 = Trim(CStr(predArr7(r7, catColIdx7)))
            If catVal7 <> "" Then
                If dictItemCategory.Exists(codeKey7) Then
                    dictItemCategory(codeKey7) = catVal7
                Else
                    dictItemCategory.Add codeKey7, catVal7
                End If
            End If
        End If
    Next r7
End Sub

' 「操作パネル」シートに在庫データ取込ボタンが無ければ追加する
Sub EnsureItemAttributeImportButton()
    Dim wsPanel As Worksheet
    Call MigrateRenamedSheets
    Call MigrateRenamedButtons

    On Error Resume Next
    Set wsPanel = ThisWorkbook.Sheets("操作パネル")
    On Error GoTo 0
    If wsPanel Is Nothing Then Exit Sub

    Dim existing As Shape
    On Error Resume Next
    Set existing = wsPanel.Shapes("在庫データ取込ボタン")
    On Error GoTo 0
    If existing Is Nothing Then
        Dim btn As Button
        Set btn = wsPanel.Buttons.Add(wsPanel.Range("H4").Left, wsPanel.Range("H4").Top, 220, 36)
        btn.Name = "在庫データ取込ボタン"
        btn.OnAction = "ImportItemAttributeMaster"
        btn.Characters.Text = "在庫データを取り込む"
        btn.Font.Size = 12
        btn.Font.Bold = True
    End If

    Call LayoutPanelButtons
End Sub

' locKey(機番+段+列)に対応する品コードをCFシート等の対応表(dictLocCode)から引き、
' 在庫データ(品コード→カテゴリー・重量・体積)を使ってdictItemCat/dictItemWt/dictItemVolに
' locKeyキーで登録する。品コードが引けない、または属性マスタに該当が無い場合は何もしない
Sub ResolveItemAttr(locKey As String, ByVal mach As Integer, ByVal dan As Integer, ByVal retsu As Integer, dictLocCode As Object, dictItemCategory As Object, dictItemWeightMaster As Object, dictItemVolumeMaster As Object, dictItemCat As Object, dictItemWt As Object, dictItemVol As Object)
    If dictItemCategory.Count = 0 And dictItemWeightMaster.Count = 0 And dictItemVolumeMaster.Count = 0 Then Exit Sub

    Dim locCodeKey As String: locCodeKey = CStr(CLng(mach) * 10000& + CLng(dan) * 100& + CLng(retsu))
    If Not dictLocCode.Exists(locCodeKey) Then Exit Sub

    Dim rawCode As String: rawCode = Trim(CStr(dictLocCode(locCodeKey)))
    If rawCode = "" Or Not IsNumeric(rawCode) Then Exit Sub
    Dim codeKey As String: codeKey = CStr(CLng(rawCode))

    If dictItemCategory.Exists(codeKey) Then dictItemCat(locKey) = dictItemCategory(codeKey)
    If dictItemWeightMaster.Exists(codeKey) Then dictItemWt(locKey) = dictItemWeightMaster(codeKey)
    If dictItemVolumeMaster.Exists(codeKey) Then dictItemVol(locKey) = dictItemVolumeMaster(codeKey)
End Sub

' ----------------------------------------------------
' 操作パネル(マクロの説明・実行ボタン)
' ----------------------------------------------------

' 「操作パネル」シートが無い場合、マクロの説明と実行ボタンを自動生成する。
' ブックの先頭シートとして配置し、以降このシートの左隣に各種出力シート(AB対面分散など)が追加されていく。
Sub EnsureOperationPanelSheet()
    Dim wsPanel As Worksheet
    Call MigrateRenamedSheets
    Call MigrateRenamedButtons

    On Error Resume Next
    Set wsPanel = ThisWorkbook.Sheets("操作パネル")
    On Error GoTo 0
    If wsPanel Is Nothing Then
        Set wsPanel = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
        wsPanel.Name = "操作パネル"

        wsPanel.Columns("A:A").ColumnWidth = 3
        wsPanel.Columns("B:H").ColumnWidth = 14

        wsPanel.Range("B2:H2").Merge
        wsPanel.Range("B2").Value = "【AB対面分散 操作パネル】"
        wsPanel.Range("B2").Font.Bold = True: wsPanel.Range("B2").Font.Size = 16
        wsPanel.Range("B2").HorizontalAlignment = xlLeft

        wsPanel.Range("B4:H26").Merge
        wsPanel.Range("B4").Value = _
            "このワークブックには、AB(自動倉庫ラック)編成の動線最適化に関するマクロが入っています。" & _
            "AB対面分散(Module3):ピッキング実績ログを解析して、AB内で同一号機・同一ゾーン(対面)で" & _
            "同時に出庫されやすい商品同士を検出し、それらを別ゾーンへ分散配置し直すための入替候補を提案します。" & _
            "同時ピッキングの集中を緩和し、機番間の作業負荷を均等化することを目的としています。" & vbCrLf & vbCrLf & _
            "【使い方(AB対面分散)】" & vbCrLf & _
            "①下の「AB対面分散を実行」ボタンを押す" & vbCrLf & _
            "②品名マスタ(S01)・ロケーションマスタ(S74)を使う場合はファイルを選ぶ(使わない場合はキャンセルでよい)" & vbCrLf & _
            "③ピッキング実績ファイル(S71で始まるファイル・複数選択可)を選ぶ" & vbCrLf & _
            "④「AB対面分散」「同号機分散」シートに入替候補・ヒートマップ・KPIが出力される" & vbCrLf & vbCrLf & _
            "【他にもあるマクロ】" & vbCrLf & _
            "「予測データ取込」(WMS等のCSVを取り込む)・「在庫データ取込」(在庫状況ダウンロード・WF021L1形式、任意)・" & _
            "「予測グラフ」「実績グラフ」(機番別構成比グラフ、実績側は「日別実績」も更新)・「号機間バランス」" & _
            "(予測データを機番別目標構成比に近づけるロケーション入替指示、同じ段の中でのみ入替え)・「ゾーンバランス」" & _
            "(予測データ・「品名実績」を元に、出荷回数順にAB(1～46号機)・Cバラ(51～68号機)・X拡張(70号機以上)の" & _
            "3ゾーン間で入替候補を作成、段は問わない)・「シート並び替え」(シートタブの並び順を整える)の各ボタンも" & _
            "用意されていますが、初回はまだボタンが表示されていない場合があります。その場合はAlt+F8のマクロ一覧から、" & _
            "対応するマクロ(ImportPredictionData・ImportItemAttributeMaster・CreateForecastRatioChart・" & _
            "CreateActualRatioChart・CreateRelocationPlan・CreateZoneRebalancePlan・SortKnownSheets)を" & _
            "一度実行すると、以降はボタンとして表示されます。" & vbCrLf & vbCrLf & _
            "【カスタマイズ】" & vbCrLf & _
            "除外機番・除外ロケーション・除外品コード・機番回数比シート名・入替候補件数・最大機番・AB間口数・機番別目標構成比などは「設定」シートで変更できます" & _
            "(シートが無ければ実行時に自動作成されます)。機番別目標構成比を入力すると、入替提案が奇数偶数バランスより目標比率への近さを優先します" & _
            "(ゾーンバランスは除外設定のみ共有し、機番別目標構成比は使いません)。" & _
            "「KPI」シートの各指標は、判定基準(合格/注意/不合格)に応じて緑・黄・赤に色分けされます" & _
            "(基準は「KPI」シート右側に一覧表示)。"
        wsPanel.Range("B4").Font.Size = 11
        wsPanel.Range("B4").WrapText = True
        wsPanel.Range("B4").VerticalAlignment = xlTop
        wsPanel.Rows("4:26").RowHeight = 18

        Dim btn As Button
        Set btn = wsPanel.Buttons.Add(wsPanel.Range("B28").Left, wsPanel.Range("B28").Top, 220, 36)
        btn.Name = "AB対面分散ボタン"
        btn.OnAction = "OptimizeABFormationFlow"
        btn.Characters.Text = "AB対面分散を実行"
        btn.Font.Size = 12
        btn.Font.Bold = True
    End If

    ' ボタンが下に伸び続けないよう、既存のボタンをすべて2列に並び替える
    ' (Module8・Module9・Module10のEnsure系Subからも毎回呼び出される)
    Call EnsureZoneRebalanceButton
    Call EnsureSortSheetsButton
    Call LayoutPanelButtons
End Sub

' 「操作パネル」シート上の各種ボタンを、決められた順序で2列に並べ直す
' (新しいボタンが追加されるたびに1列で下に伸び続けるのを防ぐため、名前が存在するものだけを詰めて配置する)
Sub LayoutPanelButtons()
    Dim wsPanel As Worksheet
    Call MigrateRenamedSheets
    Call MigrateRenamedButtons

    On Error Resume Next
    Set wsPanel = ThisWorkbook.Sheets("操作パネル")
    On Error GoTo 0
    If wsPanel Is Nothing Then Exit Sub

    Const BTN_WIDTH As Double = 220
    Const BTN_HEIGHT As Double = 36
    Const GAP_X As Double = 20
    Const GAP_Y As Double = 16
    Dim baseLeft As Double: baseLeft = wsPanel.Range("B28").Left
    Dim baseTop As Double: baseTop = wsPanel.Range("B28").Top

    Dim orderNames As Variant
    orderNames = Array("AB対面分散ボタン", "予測データ取込ボタン", "在庫データ取込ボタン", "予測グラフボタン", "実績グラフボタン", "号機間バランスボタン", "ゾーンバランスボタン", "シート並び替えボタン")

    Dim idx As Long, placedCount As Long: placedCount = 0
    For idx = LBound(orderNames) To UBound(orderNames)
        Dim shp As Shape
        On Error Resume Next
        Set shp = wsPanel.Shapes(CStr(orderNames(idx)))
        On Error GoTo 0
        If Not shp Is Nothing Then
            Dim colIdx As Long: colIdx = placedCount Mod 2
            Dim rowIdx As Long: rowIdx = placedCount \ 2
            shp.Left = baseLeft + colIdx * (BTN_WIDTH + GAP_X)
            shp.Top = baseTop + rowIdx * (BTN_HEIGHT + GAP_Y)
            placedCount = placedCount + 1
        End If
        Set shp = Nothing
    Next idx
End Sub

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
    If Not wsSet Is Nothing Then
        ' シート自体は既存でも、以下の2つは毎回最新化する(過去バージョンで作られたシートには
        ' 属性考慮の設定行(K9～K12)自体が無いことがあり、その場合はここで追加で補完する)
        Call EnsureAttrWeightSettings(wsSet)
        Call EnsureAttrCheckBox(wsSet)
        Call EnsureZoneWeekdaySetting(wsSet)
        Exit Sub
    End If

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
    wsSet.Columns("M:M").ColumnWidth = 3   ' 区切り
    wsSet.Columns("N:O").ColumnWidth = 12  ' 機番別目標構成比(機番/目標構成比%)

    wsSet.Range("A1:I1").Merge
    wsSet.Range("A1").Value = _
        "AB対面分散で除外する条件をここで設定します。①除外機番:スワップ対象・AB稼働率スコアから機番ごと除外。" & _
        "②除外ロケーション:常時使用スロットなど機番×段×列の範囲を、スワップ対象・稼働率・ヒートマップ集計のすべてから" & _
        "除外(段From/To・列From/Toはそれぞれ空欄にすると「全段」「全列」扱いになる)。" & _
        "③除外品コード:その品コードを格納場所を問わず全ての集計・スワップ対象から除外" & _
        "(CFシートの品コード列と同じ値で指定)。" & _
        "④機番別目標構成比:各機番の目標構成比(%)を入力すると、入替提案が奇数偶数バランスより目標比率への近さを" & _
        "優先するようになる(未入力ならこれまでどおり奇数偶数バランス優先)。合計が100%になっていなくても、" & _
        "入力した機番どうしの相対バランスとして扱われる(Cバラ等AB以外への出荷分があっても問題ない)。" & _
        "⑤属性考慮係数(L9～L11):在庫データ(在庫状況ダウンロード・WF021L1形式のCSV、任意)を読み込んだ場合のみ有効。" & _
        "入替候補の選定時、入替先号機の同カテゴリー品の集中度・サイズ差・重量差をスコアに軽く反映する" & _
        "(値が大きいほど強く反映)。各表の5行目以降に追加・削除して使ってください。" & _
        "⑥ゾーンバランス確認曜日(K14/L14):ゾーンバランス(Module10)の出荷回数ランキングに使う「品名実績」の" & _
        "曜日列(月・火・水・木・金・土・日)を指定する(既定は月)。" & _
        "⑦KPI判定基準:「KPI」シートに記録される各指標は、あらかじめ決められた基準に応じて緑(合格)・黄(注意)・" & _
        "赤(不合格)に自動で色分けされます(基準は「KPI」シート右側に一覧表示、この「設定」シートでは変更できません)。"
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
    wsSet.Range("L5").Value = 15 ' 「AB対面分散」に出力する入替候補の最大行数

    wsSet.Range("K6").Value = "最大機番"
    wsSet.Range("K6").Font.Bold = True
    wsSet.Range("L6").Value = 46 ' 拠点のラック総数(最大の機番)。この番号までを集計・スワップ対象にする

    wsSet.Range("K7").Value = "AB間口数"
    wsSet.Range("K7").Font.Bold = True
    wsSet.Range("L7").Value = 900 ' ABの総間口数。AB得意先スコアの理論値(回数上位◯件)算出に使う

    wsSet.Range("K8").Value = "ロケ変候補件数"
    wsSet.Range("K8").Font.Bold = True
    wsSet.Range("L8").Value = 20 ' 「号機間バランス」(予測データに基づく目標構成比への調整案)に出力する候補の最大件数。入替候補件数(L5)とは別の設定

    Call EnsureAttrWeightSettings(wsSet)

    wsSet.Range("K13").Value = "在庫データ"
    wsSet.Range("K13").Font.Bold = True

    wsSet.Range("N3").Value = "■機番別目標構成比"
    wsSet.Range("N3").Font.Bold = True
    wsSet.Range("N4").Value = "機番": wsSet.Range("O4").Value = "目標構成比(%)"
    wsSet.Range("N4:O4").Font.Bold = True
    ' 「AB01」のような機番ラベル、素の数値、C01・C02・Xのような機番以外のカテゴリラベル(構成比グラフ用)のいずれも入力できる。
    ' 既定値は実際の目標構成比(AB01～AB46・C01・C02・X)を初期値として入れておく(合計100%)
    Dim defaultTargetLabels As Variant
    Dim defaultTargetValues As Variant
    defaultTargetLabels = Array("AB01", "AB02", "AB03", "AB04", "AB05", "AB06", "AB07", "AB08", "AB09", "AB10", "AB11", "AB12", "AB13", "AB14", "AB15", "AB16", "AB17", "AB18", "AB19", "AB20", "AB21", "AB22", "AB23", "AB24", "AB25", "AB26", "AB27", "AB28", "AB29", "AB30", "AB31", "AB32", "AB33", "AB34", "AB35", "AB36", "AB37", "AB38", "AB39", "AB40", "AB41", "AB42", "AB43", "AB44", "AB45", "AB46", "C01", "C02", "X")
    defaultTargetValues = Array(1.8, 1.8, 1.8, 1.8, 1.9, 1.9, 1.9, 1.9, 2#, 2#, 2#, 2#, 2.1, 2.1, 2.1, 2.1, 2.1, 2.1, 2.1, 2.1, 2.1, 2.1, 2.1, 2.1, 2.1, 2.1, 2.1, 2.1, 2.1, 2.1, 2.1, 2.1, 2.1, 2.1, 2#, 2#, 2#, 2#, 2#, 2#, 1.9, 1.9, 1.8, 1.8, 1.8, 1.8, 3.5, 3.5, 1#)
    Dim dti As Long
    For dti = 0 To UBound(defaultTargetLabels)
        wsSet.Cells(5 + dti, 14).Value = defaultTargetLabels(dti)
        wsSet.Cells(5 + dti, 15).Value = defaultTargetValues(dti)
    Next dti

    Call EnsureAttrCheckBox(wsSet)
    Call EnsureZoneWeekdaySetting(wsSet)
End Sub

' 「設定」シートに「ゾーンバランス確認曜日」(K14/L14)が無ければ追加する
Sub EnsureZoneWeekdaySetting(wsSet As Worksheet)
    If Trim(CStr(wsSet.Range("K14").Value)) <> "" Then Exit Sub

    wsSet.Range("K14").Value = "ゾーンバランス確認曜日"
    wsSet.Range("K14").Font.Bold = True
    wsSet.Range("L14").Value = "月" ' ゾーンバランスの出荷回数ランキングに使う「品名実績」の曜日列(月・火・水・木・金・土・日)
    With wsSet.Range("L14").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="月,火,水,木,金,土,日"
    End With
End Sub

' カテゴリー重み・サイズ重み・重量重み・カテゴリー粒度(K9:L12)が無ければ追加する
' (既存の「設定」シートにはこれらの行自体が無いことがあるため、シートの有無に関わらず毎回呼び出して補完する。
' 既に入力済みならユーザーの設定値を尊重し、上書きしない)
Sub EnsureAttrWeightSettings(wsSet As Worksheet)
    If Trim(CStr(wsSet.Range("K9").Value)) <> "" Then Exit Sub

    wsSet.Range("K9").Value = "カテゴリー重み"
    wsSet.Range("K9").Font.Bold = True
    wsSet.Range("L9").Value = 0.005 ' 入替先号機の同カテゴリー品1件あたりの減点係数(在庫データ読込時のみ有効)

    wsSet.Range("K10").Value = "サイズ重み"
    wsSet.Range("K10").Font.Bold = True
    wsSet.Range("L10").Value = 0.01 ' サイズ(体積)差1桁(対数比)あたりの減点係数

    wsSet.Range("K11").Value = "重量重み"
    wsSet.Range("K11").Font.Bold = True
    wsSet.Range("L11").Value = 0.01 ' 重量差1桁(対数比)あたりの減点係数

    wsSet.Range("K12").Value = "カテゴリー粒度"
    wsSet.Range("K12").Font.Bold = True
    wsSet.Range("L12").Value = "大分類" ' 在庫データのカテゴリー一致判定に使う粒度(大分類/中分類/小分類)
    With wsSet.Range("L12").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="大分類,中分類,小分類"
    End With
End Sub

' 「在庫データを考慮する」チェックボックスが無ければ作成し、既存のものでもサイズ・位置・キャプションを
' 常に最新化する(古いバージョンで作られた、隣の表と重なるサイズのチェックボックスが残っていても直る)
Sub EnsureAttrCheckBox(wsSet As Worksheet)
    ' 過去バージョンで名前を付けずに作成した重複チェックボックスが残っていることがあるため、
    ' 正しい名前(在庫データ考慮チェック)以外で「考慮する」を含むチェックボックスは削除してから作り直す
    ' (Deleteしながら列挙すると取りこぼすことがあるため、対象名を先に集めてから別ループで削除する)
    Dim namesToDelete As Collection: Set namesToDelete = New Collection
    Dim cb As CheckBox
    For Each cb In wsSet.CheckBoxes
        If cb.Name <> "在庫データ考慮チェック" Then
            If InStr(cb.Caption, "考慮する") > 0 Then namesToDelete.Add cb.Name
        End If
    Next cb
    Dim delName As Variant
    For Each delName In namesToDelete
        On Error Resume Next
        wsSet.CheckBoxes(CStr(delName)).Delete
        On Error GoTo 0
    Next delName

    Dim chkAttr As CheckBox
    On Error Resume Next
    Set chkAttr = wsSet.CheckBoxes("在庫データ考慮チェック")
    On Error GoTo 0
    If chkAttr Is Nothing Then
        Set chkAttr = wsSet.CheckBoxes.Add(wsSet.Range("L13").Left, wsSet.Range("L13").Top - 2, 110, 18)
        chkAttr.Name = "在庫データ考慮チェック"
        chkAttr.LinkedCell = "$L$13"
        chkAttr.Value = xlOn ' オフにすると、「在庫データ」シートを取込済みでも入替候補選定への反映をスキップする
    Else
        chkAttr.Left = wsSet.Range("L13").Left
        chkAttr.Top = wsSet.Range("L13").Top - 2
        chkAttr.Width = 110
        chkAttr.Height = 18
    End If
    chkAttr.Caption = "考慮する"

    ' リンクセル(L13)はTRUE/FALSEの値そのものは保持しつつ、チェックボックスの陰から文字が
    ' はみ出て見えないよう、文字色を白にして見た目上は非表示にする(NumberFormatではTRUE/FALSEを隠せないため)
    wsSet.Range("L13").Font.Color = RGB(255, 255, 255)
End Sub

' 「設定」シートの内容を読み込み、除外機番・除外品コードの辞書と除外ロケーションの配列、シート名・件数・機番範囲設定を組み立てる
Sub LoadExclusionSettings(dictExcludedMach As Object, ByRef locMach() As Long, ByRef locDanFrom() As Long, ByRef locDanTo() As Long, ByRef locColFrom() As Long, ByRef locColTo() As Long, ByRef locCount As Long, dictExcludedItemCode As Object, ByRef ratioSheetName As String, ByRef maxSwapRows As Long, ByRef maxMachNum As Long, ByRef abSlotCount As Long, dictTargetRatio As Object, ByRef catWeight As Double, ByRef sizeWeight As Double, ByRef weightWeightCoef As Double)
    locCount = 0
    ReDim locMach(1 To 1)
    ReDim locDanFrom(1 To 1)
    ReDim locDanTo(1 To 1)
    ReDim locColFrom(1 To 1)
    ReDim locColTo(1 To 1)
    ratioSheetName = "機番回数比"
    maxSwapRows = 15
    maxMachNum = 46
    abSlotCount = 900

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

    ' 最大機番(L6)。1以上の数値が入っていればそれを使う(拠点のラック総数に合わせる)
    If IsNumeric(wsSet.Range("L6").Value) Then
        If CLng(wsSet.Range("L6").Value) >= 1 Then maxMachNum = CLng(wsSet.Range("L6").Value)
    End If

    ' AB間口数(L7)。1以上の数値が入っていればそれを使う
    If IsNumeric(wsSet.Range("L7").Value) Then
        If CLng(wsSet.Range("L7").Value) >= 1 Then abSlotCount = CLng(wsSet.Range("L7").Value)
    End If

    ' 属性考慮係数(L9～L11)。0以上の数値が入っていればそれを使う(在庫データ読込時のみ実際に効く)
    If IsNumeric(wsSet.Range("L9").Value) Then
        If CDbl(wsSet.Range("L9").Value) >= 0 Then catWeight = CDbl(wsSet.Range("L9").Value)
    End If
    If IsNumeric(wsSet.Range("L10").Value) Then
        If CDbl(wsSet.Range("L10").Value) >= 0 Then sizeWeight = CDbl(wsSet.Range("L10").Value)
    End If
    If IsNumeric(wsSet.Range("L11").Value) Then
        If CDbl(wsSet.Range("L11").Value) >= 0 Then weightWeightCoef = CDbl(wsSet.Range("L11").Value)
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

    ' 除外品コードリスト(I列、5行目以降)。ロケーションマスタ側の品コードは先頭ゼロ付きの文字列で
    ' 保持されているため、入力された値が数値の場合は先頭ゼロを除いた形でも登録し、
    ' 表記ゆれ(0326298 と 326298 など)があっても一致するようにする
    Dim lastI As Long: lastI = wsSet.Cells(wsSet.Rows.Count, "I").End(xlUp).Row
    Dim rI As Long
    For rI = 5 To lastI
        Dim codeStr As String: codeStr = Trim(CStr(wsSet.Cells(rI, 9).Value))
        If codeStr <> "" Then
            dictExcludedItemCode(codeStr) = True
            If IsNumeric(codeStr) Then dictExcludedItemCode(CStr(CLng(codeStr))) = True
        End If
    Next rI

    ' 機番別目標構成比(N:O列=機番/目標構成比%、5行目以降)。未入力ならdictTargetRatioは空のまま
    ' (呼び出し側で「未入力なら奇数偶数バランス優先」のフォールバックに使う)。
    ' N列は「AB01」のような機番ラベル、素の数値(1など)、またはC01・C02・Xのような
    ' 機番以外のカテゴリラベル(構成比グラフでのみ使う。スワップ判定では無視される)のいずれでもよい
    Dim lastN As Long: lastN = wsSet.Cells(wsSet.Rows.Count, "N").End(xlUp).Row
    Dim rN As Long
    For rN = 5 To lastN
        Dim trLabel As String: trLabel = Trim(CStr(wsSet.Cells(rN, 14).Value))
        If trLabel <> "" And IsNumeric(wsSet.Cells(rN, 15).Value) Then
            Dim trKeyStr As String
            If trLabel Like "AB##" Then
                trKeyStr = CStr(CInt(Mid(trLabel, 3, 2))) ' 「AB01」→「1」
            ElseIf IsNumeric(trLabel) Then
                trKeyStr = CStr(CLng(trLabel)) ' 素の数値がそのまま入っている場合(従来形式)
            Else
                trKeyStr = trLabel ' C01・C02・Xなど機番以外のカテゴリはラベルのままキーにする
            End If
            dictTargetRatio(trKeyStr) = CDbl(wsSet.Cells(rN, 15).Value) / 100
        End If
    Next rN
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
' 品コードは先頭ゼロの有無で表記ゆれが起きるため、元の文字列と先頭ゼロを除いた数値表記の両方で照合する
Function IsExcludedItemCode(dictLocCode As Object, dictExcludedItemCode As Object, ByVal mach As Integer, ByVal dan As Integer, ByVal retsu As Integer) As Boolean
    If dictExcludedItemCode.Count = 0 Then Exit Function
    Dim locCodeKey As String: locCodeKey = CStr(CLng(mach) * 10000& + CLng(dan) * 100& + CLng(retsu))
    If dictLocCode.Exists(locCodeKey) Then
        Dim rawCode As String: rawCode = Trim(CStr(dictLocCode(locCodeKey)))
        If dictExcludedItemCode.Exists(rawCode) Then
            IsExcludedItemCode = True
        ElseIf IsNumeric(rawCode) Then
            IsExcludedItemCode = dictExcludedItemCode.Exists(CStr(CLng(rawCode)))
        End If
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

' ----------------------------------------------------
' シートタブの並び順を、決められた順序(予測データ→在庫データ→予測グラフ→
' 実績グラフ→Cバラ交換→AB対面分散→同号機分散→号機間バランス→
' ゾーンバランス→操作パネル→設定→日別実績→品名実績→KPI)に揃える。
' このブックに存在しないシートは読み飛ばす(バンドルによって作成される
' シートが異なるため)。この一覧に無いシートの並び順は変更しない。
' ----------------------------------------------------
Sub SortKnownSheets()
    Call EnsureSortSheetsButton

    Dim orderNames As Variant
    orderNames = Array("予測データ", "在庫データ", "予測グラフ", "実績グラフ", "Cバラ交換", "AB対面分散", "同号機分散", "号機間バランス", "ゾーンバランス", "操作パネル", "設定", "日別実績", "品名実績", "KPI")

    Dim prevSheet As Worksheet: Set prevSheet = Nothing
    Dim idx As Long
    For idx = LBound(orderNames) To UBound(orderNames)
        Dim ws As Worksheet
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(CStr(orderNames(idx)))
        On Error GoTo 0
        If Not ws Is Nothing Then
            If prevSheet Is Nothing Then
                ws.Move Before:=ThisWorkbook.Sheets(1)
            Else
                ws.Move After:=prevSheet
            End If
            Set prevSheet = ws
        End If
        Set ws = Nothing
    Next idx

    MsgBox "シートの並び順を整えました。", vbInformation
End Sub

' 「操作パネル」シートにシート並び替えボタンが無ければ追加する
Sub EnsureSortSheetsButton()
    Dim wsPanel As Worksheet
    On Error Resume Next
    Set wsPanel = ThisWorkbook.Sheets("操作パネル")
    On Error GoTo 0
    If wsPanel Is Nothing Then Exit Sub

    Dim existing As Shape
    On Error Resume Next
    Set existing = wsPanel.Shapes("シート並び替えボタン")
    On Error GoTo 0
    If existing Is Nothing Then
        Dim btn As Button
        Set btn = wsPanel.Buttons.Add(wsPanel.Range("B28").Left, wsPanel.Range("B28").Top, 220, 36)
        btn.Name = "シート並び替えボタン"
        btn.OnAction = "SortKnownSheets"
        btn.Characters.Text = "シート並び替え"
        btn.Font.Size = 12
        btn.Font.Bold = True
    End If

    Call LayoutPanelButtons
End Sub

' ----------------------------------------------------
' 過去のバージョンで使っていたシート名が残っている場合、蓄積データ(日別実績・
' KPIの履歴など)を失わないよう、新しいシート名にリネームして引き継ぐ。
' 新しい名前のシートがまだ存在しない場合のみ実施する(両方あるときは触らない)。
' ----------------------------------------------------
Sub MigrateRenamedSheets()
    Dim pairs As Variant
    pairs = Array( _
        Array("在庫商品マスタ", "在庫データ"), _
        Array("予測構成比グラフ", "予測グラフ"), _
        Array("実績構成比グラフ", "実績グラフ"), _
        Array("同時ピッキング交換指示書", "Cバラ交換"), _
        Array("AB編成動線最適化", "AB対面分散"), _
        Array("同号機分散ロケーション変更指示", "同号機分散"), _
        Array("ロケ変指示", "号機間バランス"), _
        Array("ゾーン間入替候補", "ゾーンバランス"), _
        Array("日別ロケーション実績", "日別実績"), _
        Array("AB編成KPI", "KPI") _
    )

    Dim i As Long
    For i = LBound(pairs) To UBound(pairs)
        Dim oldName As String: oldName = pairs(i)(0)
        Dim newName As String: newName = pairs(i)(1)
        Dim wsOld As Worksheet, wsNew As Worksheet
        On Error Resume Next
        Set wsOld = ThisWorkbook.Sheets(oldName)
        Set wsNew = ThisWorkbook.Sheets(newName)
        On Error GoTo 0
        If Not wsOld Is Nothing And wsNew Is Nothing Then
            wsOld.Name = newName
        End If
        Set wsOld = Nothing
        Set wsNew = Nothing
    Next i
End Sub

' 過去のバージョンで使っていたボタン名が「操作パネル」に残っている場合、二重に
' ボタンが作られないよう、新しいボタン名・表示文言にリネームして引き継ぐ。
' 新しい名前のボタンがまだ無い場合のみ実施する(両方あるときは触らない)。
Sub MigrateRenamedButtons()
    Dim wsPanel As Worksheet
    On Error Resume Next
    Set wsPanel = ThisWorkbook.Sheets("操作パネル")
    On Error GoTo 0
    If wsPanel Is Nothing Then Exit Sub

    Dim pairs As Variant
    pairs = Array( _
        Array("ロケ変指示ボタン", "号機間バランスボタン", "号機間バランスを作成"), _
        Array("ゾーン間入替候補ボタン", "ゾーンバランスボタン", "ゾーンバランス作成"), _
        Array("在庫商品マスタ取込ボタン", "在庫データ取込ボタン", "在庫データを取り込む"), _
        Array("AB編成動線最適化ボタン", "AB対面分散ボタン", "AB対面分散を実行"), _
        Array("予測構成比グラフボタン", "予測グラフボタン", "予測グラフを作成"), _
        Array("実績構成比グラフボタン", "実績グラフボタン", "実績グラフを作成") _
    )

    Dim i As Long
    For i = LBound(pairs) To UBound(pairs)
        Dim oldName As String: oldName = pairs(i)(0)
        Dim newName As String: newName = pairs(i)(1)
        Dim newCaption As String: newCaption = pairs(i)(2)
        Dim shpOld As Shape, shpNew As Shape
        On Error Resume Next
        Set shpOld = wsPanel.Shapes(oldName)
        Set shpNew = wsPanel.Shapes(newName)
        On Error GoTo 0
        If Not shpOld Is Nothing And shpNew Is Nothing Then
            shpOld.Name = newName
            shpOld.Characters.Text = newCaption
        End If
        Set shpOld = Nothing
        Set shpNew = Nothing
    Next i
End Sub
