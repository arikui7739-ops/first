Attribute VB_Name = "Module3_金沢"
Option Explicit

' 金沢は片面・編成の概念が無く、オリコンは1個ずつ移動するため、対面ペアで入替提案を作る
' 「AB対面分散」機能は搭載していない(沼南・石狩にはある)。ここには号機間バランス(Module10)・
' 構成比グラフ(Module9)と共通の設定管理・除外判定・在庫データ取込のみを置く。
' カテゴリー集中ペナルティで「サイズが近い」と判定する閾値(体積比lnの絶対値。0.4は1.5倍以内)
' (号機間バランスの入替候補選定で使う。Module10のCreateRelocationPlanから呼ばれる)
Public Const SIZE_SIMILAR_RATIO As Double = 0.4

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

' 1つのゾーン内で、奇数号機・偶数号機の組み方まで含めて最適配置した場合の
' 「理論上最小の対面ヒット数」を局所探索(Kernighan-Linに近い2分割法)で求める。
' itemsArr: そのゾーンに属するアイテムキーの配列／dictMach: アイテム→号機／weightDict: "item1,item2"(ソート済)→編成内共起回数
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
' シートが無い場合、または「設定」シートのチェックボックス(L12)がオフの場合は何もしない(辞書は空のまま=
' 従来どおりの動作になる)。ファイル選択ダイアログは出さない(取込はImportItemAttributeMasterの役目)
Sub LoadItemAttributeMasterFromSheet(dictItemCategory As Object, dictItemWeightMaster As Object, dictItemVolumeMaster As Object)
    Dim wsSetChk As Worksheet
    On Error Resume Next
    Set wsSetChk = ThisWorkbook.Sheets("設定")
    On Error GoTo 0
    If Not wsSetChk Is Nothing Then
        If wsSetChk.Range("L12").Value = False Then Exit Sub
    End If

    Dim wsAttr As Worksheet
    On Error Resume Next
    Set wsAttr = ThisWorkbook.Sheets("在庫データ")
    On Error GoTo 0
    If Not wsAttr Is Nothing Then

    ' 「カテゴリー粒度」(L11)に応じて、大分類(B列)/中分類(C列)/小分類(D列)のどれを使うかを決める
    Dim catCol5 As Long: catCol5 = 2
    If Not wsSetChk Is Nothing Then
        Select Case Trim(CStr(wsSetChk.Range("L11").Value))
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
    End If

    ' 列幅・タイトル・説明文・ボタン配置は、シートが既にあってもレイアウト変更を反映できるよう毎回更新する。
    ' 以前のバージョンで結合されたセル(B4:H18やB4:H40など)が残っていると、形の違う範囲を
    ' 結合しようとしたときにうまく反映されないことがあるため、先に結合を解除してからやり直す
    wsPanel.Range("B2:K60").UnMerge

    wsPanel.Columns("A:A").ColumnWidth = 3
    wsPanel.Columns("B:F").ColumnWidth = 14 ' 説明文エリア
    wsPanel.Columns("G:G").ColumnWidth = 3  ' 区切り
    wsPanel.Columns("H:K").ColumnWidth = 14 ' ボタン配置エリア(説明文の右側)

    wsPanel.Range("B2:K2").Merge
    wsPanel.Range("B2").Value = "【金沢 操作パネル】"
    wsPanel.Range("B2").Font.Bold = True: wsPanel.Range("B2").Font.Size = 16
    wsPanel.Range("B2").HorizontalAlignment = xlLeft

    wsPanel.Range("B4:F55").Merge
    Dim panelDesc As String
    panelDesc = _
        "このワークブックには、AB(自動倉庫ラック)の動線最適化に関する4つのマクロが入っています。" & _
        "金沢は片面・編成の概念が無く、オリコンは1個ずつ移動するため、沼南・石狩にある「AB対面分散」" & _
        "(対面ペアで入替提案を作る機能)は搭載していません。" & _
        "①予測データ取込(Module8):WMS等から出力した予測データCSVを取り込みます。②③④の元データになります。" & _
        "②構成比グラフ(Module9):予測データ・実績(S71)それぞれの号機別構成比を、目標構成比と比較できるグラフを" & _
        "作成します(号機ごとに独立した棒で表示。ペア表示はありません)。実績側は「日別実績」「品名実績」の履歴も" & _
        "自動更新します。" & _
        "③号機間バランス(Module10):予測データを元に、号機別構成比を目標構成比に近づけるロケーション入替指示を" & _
        "作成します(同じ段の中でのみ入替えます)。" & _
        "④ゾーンバランス(Module10):予測データ・「品名実績」を元に、出荷回数の順位に応じてAB(1～24号機)と" & _
        "Cバラ(61～66号機・81～88号機)の間で入替候補を作成します(段は問いません。それ以外の号機は拡張エリアとして" & _
        "対象外)。" & vbCrLf & vbCrLf
    panelDesc = panelDesc & _
        "【空のワークブックで初めて使うとき】" & vbCrLf & _
        "①VBEでこの4つのモジュール(Module3_金沢・Module8_金沢・Module9_金沢・Module10_金沢、またはModule3～10)を" & _
        "「ファイルのインポート」で追加する" & vbCrLf & _
        "②いずれかのマクロを一度実行する(右の「予測データを取り込む」ボタンでよい。ファイル選択はキャンセルして" & _
        "かまわない)。これで本シートと「設定」シートが自動作成され、以降すべてのボタンが使えるようになる" & vbCrLf & _
        "③「設定」シートの■ABブロック(金沢の号機範囲)と■号機別目標構成比を、実際のラック配置・目標値に合わせて" & _
        "入力する" & vbCrLf & vbCrLf & _
        "【使う順番の目安】" & vbCrLf & _
        "①「予測データを取り込む」でCSVを取り込む(②③④の前提)" & vbCrLf & _
        "②「予測グラフを作成」「実績グラフを作成」でグラフを作成(実績側はS71ファイルが必要)" & vbCrLf
    panelDesc = panelDesc & _
        "③「号機間バランスを作成」で号機間バランスを作成(予測データの取込と、「設定」シートの■号機別目標構成比の入力が必要)" & vbCrLf & _
        "④「ゾーンバランス作成」でゾーンバランスを作成(予測データの取込が必要。「設定」の■ゾーンバランス確認基準に" & _
        "応じて「品名実績」の指定曜日実績、または「予測」なら予測データの投入回数_予測を使う)" & vbCrLf & _
        "⑤「在庫データを取り込む」で在庫状況ダウンロード(WF021L1形式)を取り込むと(任意)、③の入替候補選定に" & _
        "サイズ・重量・カテゴリーの近さが反映されます(一度取り込めば以降のファイル選択は不要です)。" & vbCrLf & vbCrLf & _
        "【カスタマイズ】" & vbCrLf & _
        "除外号機・除外ロケーション・除外品コード・号機回数比シート名・ロケ変候補件数・ABブロック・" & _
        "号機別目標構成比などは「設定」シートで変更できます(シートが無ければ実行時に自動作成されます)。" & _
        "号機別目標構成比を入力すると、号機間バランス(Module10)が作成できるようになります" & _
        "(ゾーンバランスは除外設定のみ共有し、号機別目標構成比は使いません)。シートタブの並び順がバラバラに" & _
        "なったときは「シート並び替え」ボタンで整えられます。"
    wsPanel.Range("B4").Value = panelDesc
    wsPanel.Range("B4").Font.Size = 11
    wsPanel.Range("B4").WrapText = True
    wsPanel.Range("B4").VerticalAlignment = xlTop
    wsPanel.Rows("4:55").RowHeight = 18

    ' 「KPI」シートも他のシート同様、いずれかのマクロを実行した時点で自動生成する
    Call EnsureKPISheet

    ' このマクロだけを実行しても各ボタンがすべて揃うよう、他モジュールのボタンも一緒に用意する
    ' (Module8・Module9・Module10は同じVBAプロジェクトに揃っている前提。揃っていないとここでコンパイルエラーになる)
    Call EnsurePredictionImportButton
    Call EnsureItemAttributeImportButton
    Call EnsureRatioChartButtons
    Call EnsureRelocationPlanButton
    Call EnsureZoneRebalanceButton
    Call EnsureSortSheetsButton

    ' ボタンが下に伸び続けないよう、既存のボタンをすべて2列に並び替える
    ' (Module8・Module9・Module10のEnsure系Subからも毎回呼び出される)
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
    Dim baseLeft As Double: baseLeft = wsPanel.Range("H4").Left
    Dim baseTop As Double: baseTop = wsPanel.Range("H4").Top

    Dim orderNames As Variant
    orderNames = Array("予測データ取込ボタン", "在庫データ取込ボタン", "予測グラフボタン", "実績グラフボタン", "号機間バランスボタン", "ゾーンバランスボタン", "シート並び替えボタン")

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
' KPI(実施日・AB上限回数比率・AB実績回数比率・AB同時ピッキング回避スコア・均衡化スコアの履歴)
' ----------------------------------------------------

' 「KPI」シートが無ければ見出し行だけを用意して自動生成する
Sub EnsureKPISheet()
    Dim wsKPI As Worksheet
    Call MigrateRenamedSheets

    On Error Resume Next
    Set wsKPI = ThisWorkbook.Sheets("KPI")
    On Error GoTo 0
    If wsKPI Is Nothing Then
        Set wsKPI = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsKPI.Name = "KPI"

        wsKPI.Range("A1:H1").Merge
        wsKPI.Range("A1").Value = "【KPI推移】実施日ごとに1行で記録されます(同じ日に複数回実行した場合は上書き)"
        wsKPI.Range("A1").Font.Bold = True: wsKPI.Range("A1").Font.Size = 14

        ' 金沢はAB対面分散を行わないため、沼南・石狩の指標(AB実績回数比率・AB同時ピッキング回避スコア・
        ' 均衡化スコア等、いずれも対面分散の前後比較が前提)は使えない。記録する指標が決まるまでは
        ' 「実施日」列のみを用意しておく(号機間バランス・ゾーンバランスの実施記録用に追記予定)
        wsKPI.Range("A3").Value = "実施日"
        wsKPI.Range("A3").Interior.Color = RGB(220, 230, 255)
        wsKPI.Range("A3").Font.Bold = True

        wsKPI.Columns("A:A").ColumnWidth = 12
        wsKPI.Columns("A:A").NumberFormat = "yyyy/mm/dd"
    End If

    ' 判定基準の凡例(緑=合格/黄=注意/赤=不合格)は、指標が決まり次第ApplyKPIJudgeCriteriaに追加する。
    ' シートが既にあった場合も常に最新の内容に更新する
    Call ApplyKPIJudgeCriteria(wsKPI)
End Sub

' 「KPI」シートの各指標に判定基準(合格/注意/不合格)を条件付き書式で色分けし(緑=合格/黄=注意/赤=不合格)、
' 基準の一覧をシート右側(L:N列、データ列とは重ならない位置)に書き出す。
' EnsureKPISheetから毎回呼び出し、シートが既にあった場合も含めて常に最新の基準に更新する。
' 金沢はAB対面分散を行わないため、現時点で色分け対象の指標は無い(指標が決まり次第、
' AddKPITrafficLightで列ごとに追加する。沼南・石狩のApplyKPIJudgeCriteriaを参考にすること)
Sub ApplyKPIJudgeCriteria(wsKPI As Worksheet)
    wsKPI.Columns("L:L").ColumnWidth = 28
    wsKPI.Columns("M:M").ColumnWidth = 3
    wsKPI.Columns("N:N").ColumnWidth = 46

    wsKPI.Range("L1:N1").Merge
    wsKPI.Range("L1").Value = "■KPI判定基準(緑=合格・黄=注意・赤=不合格)"
    wsKPI.Range("L1").Font.Bold = True
    wsKPI.Range("L1").HorizontalAlignment = xlLeft

    wsKPI.Cells(2, 12).Value = "(記録する指標は今後追加予定です)"
End Sub

' ----------------------------------------------------
' 拠点カスタマイズ設定(除外号機・除外ロケーション)
' ----------------------------------------------------

' 「設定」シートが無い場合、金沢の実際のラック配置(ABブロック:1～24の1ブロック。61～66はCバラ01、
' 81～88はCバラ02、それ以外は拡張Xとして扱いAB編成の対象外)を初期値として自動生成する。
' 除外号機・除外ロケーションは拠点固有の情報が無いため空欄で初期化し、必要に応じて追記する。
' 列幅は用途ごとに固定値で設定する(説明文の長さに引っ張られて横に広がらないようにするため、AutoFitは使わない)。
Sub EnsureExclusionSettingsSheet()
    Dim wsSet As Worksheet
    On Error Resume Next
    Set wsSet = ThisWorkbook.Sheets("設定")
    On Error GoTo 0
    If Not wsSet Is Nothing Then
        ' シート自体は既存でも、以下の2つは毎回最新化する(過去バージョンで作られたシートには
        ' 属性考慮の設定行(K8～K11)自体が無いことがあり、その場合はここで追加で補完する)
        Call EnsureAttrWeightSettings(wsSet)
        Call EnsureAttrCheckBox(wsSet)
        Call EnsureZoneWeekdaySetting(wsSet)
        Exit Sub
    End If

    Set wsSet = ThisWorkbook.Sheets.Add
    wsSet.Name = "設定"

    wsSet.Columns("A:A").ColumnWidth = 10  ' 除外号機
    wsSet.Columns("B:B").ColumnWidth = 3   ' 区切り
    wsSet.Columns("C:G").ColumnWidth = 8   ' 除外ロケーション(号機/段From/段To/列From/列To)
    wsSet.Columns("H:H").ColumnWidth = 3   ' 区切り
    wsSet.Columns("I:I").ColumnWidth = 14  ' 除外品コード
    wsSet.Columns("I:I").NumberFormat = "@" ' 品コードは先頭0落ち・数値化を防ぐため文字列扱いにする
    wsSet.Columns("J:J").ColumnWidth = 3   ' 区切り
    wsSet.Columns("K:K").ColumnWidth = 20  ' シート名設定ラベル
    wsSet.Columns("L:L").ColumnWidth = 16  ' シート名設定値
    wsSet.Columns("M:M").ColumnWidth = 3   ' 区切り
    wsSet.Columns("N:O").ColumnWidth = 10  ' ABブロック(開始号機/終了号機)
    wsSet.Columns("P:P").ColumnWidth = 3   ' 区切り
    wsSet.Columns("Q:R").ColumnWidth = 12  ' 号機別目標構成比(号機/目標構成比%)

    wsSet.Range("A1:R1").Merge
    wsSet.Range("A1").Value = _
        "この「設定」シートは、構成比グラフ(Module9)・号機間バランス(Module10)・ゾーンバランス(Module10)で" & _
        "共通して使う設定です。金沢は片面・編成の概念が無いため、AB対面分散(奇数偶数ペアでの入替)は行いません。" & _
        "①除外号機:AB稼働率スコア・号機間バランスの入替対象から号機ごと除外。" & _
        "②除外ロケーション:常時使用スロットなど号機×段×列の範囲を、入替対象・稼働率集計のすべてから除外" & _
        "(段From/To・列From/Toはそれぞれ空欄にすると「全段」「全列」扱いになる)。" & _
        "③除外品コード:その品コードを格納場所を問わず全ての集計・入替対象から除外(CFシートの品コード列と同じ値で指定)。" & _
        "④シート名設定:号機回数比シート名(L4)のほか、ロケ変候補件数(L7、号機間バランスの出力件数)を数値で指定する。" & _
        "⑤ABブロック:構成比グラフ・号機間バランスがAB範囲とみなす号機範囲(複数ブロック可)。ブロック外の号機" & _
        "(金沢ではCバラ01=61～66、Cバラ02=81～88、拡張X=それ以外)は号機間バランスの対象外" & _
        "(ゾーンバランスが対象とするAB・Cバラ・X拡張の範囲とは別の区分です)。" & _
        "⑥号機別目標構成比:各号機の目標構成比(%)を入力すると、号機間バランス(Module10)が作成できるようになります" & _
        "(未入力の場合、号機間バランスは作成できません)。" & _
        "合計が100%になっていなくても、入力した号機どうしの相対バランスとして扱われる(Cバラ等AB以外への出荷分があっても問題ない)。" & _
        "ゾーンバランスは①②③の除外設定のみ共有し、この目標構成比は使いません" & _
        "(AB=1～24号機とCバラ=61～66号機・81～88号機の間で、出荷回数順に構成比92%・7%を目指します。" & _
        "それ以外の号機は拡張エリアとして入替の対象外です。出荷回数のランキングに使う「品名実績」の曜日は" & _
        "■ゾーンバランス確認基準で変更できます(既定は月曜。「予測」を選ぶと実績ではなく" & _
        "予測データの投入回数_予測をそのまま順位付けに使います)。"
    wsSet.Range("A1").Value = wsSet.Range("A1").Value & _
        "⑦属性考慮係数(L8～L10):在庫データ(在庫状況ダウンロード・WF021L1形式のCSV、任意)を読み込んだ場合のみ有効。" & _
        "号機間バランスの入替候補選定時、入替先号機の同カテゴリー品の集中度・サイズ差・重量差をスコアに軽く反映する" & _
        "(値が大きいほど強く反映)。各表の5行目以降に追加・削除して使ってください。" & _
        "⑧KPI判定基準:「KPI」シートに記録する指標は今後追加予定です(現時点では判定基準はありません)。"
    wsSet.Range("A1").Font.Bold = True
    wsSet.Range("A1").WrapText = True
    wsSet.Range("A1").VerticalAlignment = xlTop
    wsSet.Rows(1).RowHeight = 130

    wsSet.Range("A3").Value = "■除外号機"
    wsSet.Range("A3").Font.Bold = True
    wsSet.Range("A4").Value = "号機"
    wsSet.Range("A4").Font.Bold = True

    wsSet.Range("C3").Value = "■除外ロケーション"
    wsSet.Range("C3").Font.Bold = True
    wsSet.Range("C4").Value = "号機": wsSet.Range("D4").Value = "段From": wsSet.Range("E4").Value = "段To": wsSet.Range("F4").Value = "列From": wsSet.Range("G4").Value = "列To"
    wsSet.Range("C4:G4").Font.Bold = True

    wsSet.Range("I3").Value = "■除外品コード"
    wsSet.Range("I3").Font.Bold = True
    wsSet.Range("I4").Value = "品コード"
    wsSet.Range("I4").Font.Bold = True

    wsSet.Range("K3").Value = "■シート名設定"
    wsSet.Range("K3").Font.Bold = True
    wsSet.Range("K4").Value = "号機回数比シート名"
    wsSet.Range("K4").Font.Bold = True
    wsSet.Range("L4").Value = "号機回数比" ' AB稼働率スコアの目標比率を読むシート名。拠点によって名前が違う場合はここを書き換える

    wsSet.Range("K7").Value = "ロケ変候補件数"
    wsSet.Range("K7").Font.Bold = True
    wsSet.Range("L7").Value = 20 ' 「号機間バランス」(予測データに基づく目標構成比への調整案)に出力する候補の最大件数。入替候補件数(L5)とは別の設定

    Call EnsureAttrWeightSettings(wsSet)

    wsSet.Range("K12").Value = "在庫データ"
    wsSet.Range("K12").Font.Bold = True

    wsSet.Range("N3").Value = "■ABブロック"
    wsSet.Range("N3").Font.Bold = True
    wsSet.Range("N4").Value = "開始号機": wsSet.Range("O4").Value = "終了号機"
    wsSet.Range("N4:O4").Font.Bold = True
    ' 金沢の実際のラック配置:1～24(ゾーン1～12、1ブロックのみ)。61～66(Cバラ01)、81～88(Cバラ02)、
    ' それ以外(拡張X)はここに含めない=AB編成のゾーン・スワップ対象外になる
    wsSet.Range("N5").Value = 1: wsSet.Range("O5").Value = 24

    wsSet.Range("Q3").Value = "■号機別目標構成比"
    wsSet.Range("Q3").Font.Bold = True
    wsSet.Range("Q4").Value = "号機": wsSet.Range("R4").Value = "目標構成比(%)"
    wsSet.Range("Q4:R4").Font.Bold = True
    ' 「AB01」のような号機ラベル、素の数値、C01・C02・Xのような号機以外のカテゴリラベル(構成比グラフ用)のいずれも入力できる。
    ' 例:1号機を1.8%にしたい場合はQ5=AB01(またはQ5=1)・R5=1.8のように行を追加する(未入力なら奇数偶数バランス優先のまま)

    Call EnsureAttrCheckBox(wsSet)
    Call EnsureZoneWeekdaySetting(wsSet)
End Sub

' 「設定」シートに「ゾーンバランス確認基準」(K13/L13)が無ければ追加する
Sub EnsureZoneWeekdaySetting(wsSet As Worksheet)
    If Trim(CStr(wsSet.Range("K13").Value)) <> "" Then Exit Sub

    wsSet.Range("K13").Value = "ゾーンバランス確認基準"
    wsSet.Range("K13").Font.Bold = True
    wsSet.Range("L13").Value = "月" ' ゾーンバランスの出荷回数ランキングに使う「品名実績」の曜日列(月・火・水・木・金・土・日)、
                                    ' または「予測」(=予測データの投入回数_予測をそのまま使う)
    With wsSet.Range("L13").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="月,火,水,木,金,土,日,予測"
    End With
End Sub

' カテゴリー重み・サイズ重み・重量重み・カテゴリー粒度(K8:L11)が無ければ追加する
' (既存の「設定」シートにはこれらの行自体が無いことがあるため、シートの有無に関わらず毎回呼び出して補完する。
' 既に入力済みならユーザーの設定値を尊重し、上書きしない)
Sub EnsureAttrWeightSettings(wsSet As Worksheet)
    If Trim(CStr(wsSet.Range("K8").Value)) <> "" Then Exit Sub

    wsSet.Range("K8").Value = "カテゴリー重み"
    wsSet.Range("K8").Font.Bold = True
    wsSet.Range("L8").Value = 0.005 ' 入替先号機の同カテゴリー品1件あたりの減点係数(在庫データ読込時のみ有効)

    wsSet.Range("K9").Value = "サイズ重み"
    wsSet.Range("K9").Font.Bold = True
    wsSet.Range("L9").Value = 0.01 ' サイズ(体積)差1桁(対数比)あたりの減点係数

    wsSet.Range("K10").Value = "重量重み"
    wsSet.Range("K10").Font.Bold = True
    wsSet.Range("L10").Value = 0.01 ' 重量差1桁(対数比)あたりの減点係数

    wsSet.Range("K11").Value = "カテゴリー粒度"
    wsSet.Range("K11").Font.Bold = True
    wsSet.Range("L11").Value = "大分類" ' 在庫データのカテゴリー一致判定に使う粒度(大分類/中分類/小分類)
    With wsSet.Range("L11").Validation
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
        Set chkAttr = wsSet.CheckBoxes.Add(wsSet.Range("L12").Left, wsSet.Range("L12").Top - 2, 110, 18)
        chkAttr.Name = "在庫データ考慮チェック"
        chkAttr.LinkedCell = "$L$12"
        chkAttr.Value = xlOn ' オフにすると、「在庫データ」シートを取込済みでも入替候補選定への反映をスキップする
    Else
        chkAttr.Left = wsSet.Range("L12").Left
        chkAttr.Top = wsSet.Range("L12").Top - 2
        chkAttr.Width = 110
        chkAttr.Height = 18
    End If
    chkAttr.Caption = "考慮する"

    ' リンクセル(L12)はTRUE/FALSEの値そのものは保持しつつ、チェックボックスの陰から文字が
    ' はみ出て見えないよう、文字色を白にして見た目上は非表示にする(NumberFormatではTRUE/FALSEを隠せないため)
    wsSet.Range("L12").Font.Color = RGB(255, 255, 255)
End Sub

' 「設定」シートの内容を読み込み、除外号機・除外品コードの辞書と除外ロケーションの配列、シート名・件数・号機範囲設定を組み立てる
Sub LoadExclusionSettings(dictExcludedMach As Object, ByRef locMach() As Long, ByRef locDanFrom() As Long, ByRef locDanTo() As Long, ByRef locColFrom() As Long, ByRef locColTo() As Long, ByRef locCount As Long, dictExcludedItemCode As Object, ByRef ratioSheetName As String, ByRef maxSwapRows As Long, ByRef abSlotCount As Long, ByRef abBlockFrom() As Long, ByRef abBlockTo() As Long, ByRef abBlockCount As Long, dictTargetRatio As Object, ByRef catWeight As Double, ByRef sizeWeight As Double, ByRef weightWeightCoef As Double)
    locCount = 0
    ReDim locMach(1 To 1)
    ReDim locDanFrom(1 To 1)
    ReDim locDanTo(1 To 1)
    ReDim locColFrom(1 To 1)
    ReDim locColTo(1 To 1)
    ratioSheetName = "号機回数比"
    maxSwapRows = 15
    abSlotCount = 850
    ' ABブロックの既定値:金沢の実際のラック配置(1～24の1ブロック)
    ReDim abBlockFrom(1 To 1): ReDim abBlockTo(1 To 1)
    abBlockFrom(1) = 1: abBlockTo(1) = 24
    abBlockCount = 1

    Dim wsSet As Worksheet
    On Error Resume Next
    Set wsSet = ThisWorkbook.Sheets("設定")
    On Error GoTo 0
    If wsSet Is Nothing Then Exit Sub

    ' 号機回数比シート名(L4)。空欄ならデフォルト名のまま
    If Trim(CStr(wsSet.Range("L4").Value)) <> "" Then ratioSheetName = Trim(CStr(wsSet.Range("L4").Value))

    ' 入替候補件数(L5)。1以上の数値が入っていればそれを使う
    If IsNumeric(wsSet.Range("L5").Value) Then
        If CLng(wsSet.Range("L5").Value) >= 1 Then maxSwapRows = CLng(wsSet.Range("L5").Value)
    End If

    ' AB間口数(L6)。1以上の数値が入っていればそれを使う
    If IsNumeric(wsSet.Range("L6").Value) Then
        If CLng(wsSet.Range("L6").Value) >= 1 Then abSlotCount = CLng(wsSet.Range("L6").Value)
    End If

    ' 属性考慮係数(L8～L10)。0以上の数値が入っていればそれを使う(在庫データ読込時のみ実際に効く)
    If IsNumeric(wsSet.Range("L8").Value) Then
        If CDbl(wsSet.Range("L8").Value) >= 0 Then catWeight = CDbl(wsSet.Range("L8").Value)
    End If
    If IsNumeric(wsSet.Range("L9").Value) Then
        If CDbl(wsSet.Range("L9").Value) >= 0 Then sizeWeight = CDbl(wsSet.Range("L9").Value)
    End If
    If IsNumeric(wsSet.Range("L10").Value) Then
        If CDbl(wsSet.Range("L10").Value) >= 0 Then weightWeightCoef = CDbl(wsSet.Range("L10").Value)
    End If

    ' ABブロック(N:O列、5行目以降)。データがあれば既定値を上書きする
    Dim lastN As Long: lastN = wsSet.Cells(wsSet.Rows.Count, "N").End(xlUp).Row
    If lastN >= 5 Then
        Dim tmpCount As Long: tmpCount = 0
        Dim tmpFrom() As Long, tmpTo() As Long
        ReDim tmpFrom(1 To lastN - 4)
        ReDim tmpTo(1 To lastN - 4)
        Dim rN As Long
        For rN = 5 To lastN
            If IsNumeric(wsSet.Cells(rN, 14).Value) And IsNumeric(wsSet.Cells(rN, 15).Value) Then
                tmpCount = tmpCount + 1
                tmpFrom(tmpCount) = CLng(wsSet.Cells(rN, 14).Value)
                tmpTo(tmpCount) = CLng(wsSet.Cells(rN, 15).Value)
            End If
        Next rN
        If tmpCount > 0 Then
            ReDim abBlockFrom(1 To tmpCount)
            ReDim abBlockTo(1 To tmpCount)
            Dim ci As Long
            For ci = 1 To tmpCount
                abBlockFrom(ci) = tmpFrom(ci)
                abBlockTo(ci) = tmpTo(ci)
            Next ci
            abBlockCount = tmpCount
        End If
    End If

    ' 除外号機リスト(A列、5行目以降)
    Dim lastA As Long: lastA = wsSet.Cells(wsSet.Rows.Count, "A").End(xlUp).Row
    Dim rA As Long
    For rA = 5 To lastA
        If IsNumeric(wsSet.Cells(rA, 1).Value) And Trim(CStr(wsSet.Cells(rA, 1).Value)) <> "" Then
            dictExcludedMach(CStr(CLng(wsSet.Cells(rA, 1).Value))) = True
        End If
    Next rA

    ' 除外ロケーションリスト(C:G列=号機/段From/段To/列From/列To、5行目以降)
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

    ' 号機別目標構成比(Q:R列=号機/目標構成比%、5行目以降)。未入力ならdictTargetRatioは空のまま
    ' (呼び出し側で「未入力なら奇数偶数バランス優先」のフォールバックに使う)。
    ' Q列は「AB01」のような号機ラベル、素の数値(1など)、またはC01・C02・Xのような
    ' 号機以外のカテゴリラベル(構成比グラフでのみ使う。スワップ判定では無視される)のいずれでもよい
    Dim lastQ As Long: lastQ = wsSet.Cells(wsSet.Rows.Count, "Q").End(xlUp).Row
    Dim rQ As Long
    For rQ = 5 To lastQ
        Dim trLabel As String: trLabel = Trim(CStr(wsSet.Cells(rQ, 17).Value))
        If trLabel <> "" And IsNumeric(wsSet.Cells(rQ, 18).Value) Then
            Dim trKeyStr As String
            If trLabel Like "AB##" Then
                trKeyStr = CStr(CInt(Mid(trLabel, 3, 2))) ' 「AB01」→「1」
            ElseIf IsNumeric(trLabel) Then
                trKeyStr = CStr(CLng(trLabel)) ' 素の数値がそのまま入っている場合(従来形式)
            Else
                trKeyStr = trLabel ' C01・C02・Xなど号機以外のカテゴリはラベルのままキーにする
            End If
            dictTargetRatio(trKeyStr) = CDbl(wsSet.Cells(rQ, 18).Value) / 100
        End If
    Next rQ
End Sub

' 指定の号機が、いずれかのABブロックに含まれるかを判定する(ブロック外はAB編成のゾーン・スワップ対象外)
Function IsInABBlock(mach As Integer, abBlockFrom() As Long, abBlockTo() As Long, abBlockCount As Long) As Boolean
    Dim bi As Long
    For bi = 1 To abBlockCount
        If mach >= abBlockFrom(bi) And mach <= abBlockTo(bi) Then
            IsInABBlock = True
            Exit Function
        End If
    Next bi
    IsInABBlock = False
End Function

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

' 指定の号機が「除外号機」設定に該当するか判定する(1～4番機のような、サイズが異なる品を格納する号機など)
Function IsExcludedSlot3(dictExcludedMach As Object, mach As Integer) As Boolean
    IsExcludedSlot3 = dictExcludedMach.Exists(CStr(mach))
End Function

' 指定の号機・段・列にある品が「除外品コード」設定に該当するか判定する(CFシートのロケーション⇔品コード対応表を使って引く)
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
