Attribute VB_Name = "Module9_石狩"
Option Explicit

' ----------------------------------------------------
' 構成比グラフの作成
' 「予測グラフ」:「予測データ」シート(Module8で取込済み)の投入回数_予測をゾーン別に集計してグラフ化する
' 「実績グラフ」:ピッキング実績ファイル(S71)をダイアログで選択し、ゾーン別ヒット数を集計してグラフ化する
' レイアウトは「設定」シートの「■機番別目標構成比」と同じ考え方で、AB01～AB46は機番ペア(ゾーン)ごとに
' 奇数機番を上向き・偶数機番を下向きに表示し、Cバラ(C01・C02)・拡張(X)は上向きの単独項目として追加する。
' 「設定」シートに目標構成比が入力されていれば、実績/予測の比率と並べて目標比率も折れ線で比較できるようにする。
' データ表にはCバラ・拡張も含めるが、グラフにはAB01～AB46のみを表示する(Cバラ・拡張はグラフの対象外)。
' グラフの縦軸目盛りは予測・実績のグラフ間で見比べやすいよう固定スケール(既定は±5%・1%刻み)にし、
' 実データがそれを超える場合のみ切り上げて広げる(拠点間でも同じ基準を使う)。
' 「実績グラフ」の作成と同時に、「日別実績」シート(品名コード・ロケ分類・品名・
' ロケーション・予測回数ごとの日別実績を並べた履歴表)も更新する。日別列はG～Pの最大10列で、
' 11日目以降は最も古い日(G列)を消して1列ずつ左に詰め、新しい日をP列に追加する。
' 既存の同名シートは削除してから作り直すため、再実行すると内容が更新される
' ----------------------------------------------------

Const MAX_AB_MACH As Long = 46 ' 石狩のAB機番範囲(1～46)。この範囲は奇数/偶数ペアのゾーン表示にする
Const AXIS_DEFAULT_MAX As Double = 0.05 ' グラフ縦軸の既定スケール(±5%)
Const AXIS_UNIT As Double = 0.01 ' グラフ縦軸の目盛り間隔(1%)

Sub CreateForecastRatioChart()
    Call EnsureRatioChartButtons

    Dim wsData As Worksheet
    On Error Resume Next
    Set wsData = ThisWorkbook.Sheets("予測データ")
    On Error GoTo 0
    If wsData Is Nothing Then
        MsgBox "「予測データ」シートが見つかりません。先に「予測データを取り込む」を実行してください。", vbExclamation
        Exit Sub
    End If

    Dim dictTargetByLabel As Object
    Call LoadTargetRatioByLabel(dictTargetByLabel)

    ' 「予測データ」シートは1行目=取込情報、3行目=見出し(Module8の出力形式)。
    ' ゾーン列(AB01～AB46・C01・C02・Xなどのラベル)・投入回数_予測列を見出し名から探す
    Const HEADER_ROW As Long = 3
    Dim lastRow As Long: lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    Dim lastCol As Long: lastCol = wsData.Cells(HEADER_ROW, wsData.Columns.Count).End(xlToLeft).Column

    Dim zoneColIdx As Long: zoneColIdx = -1
    Dim cntColIdx As Long: cntColIdx = -1
    Dim hc As Long
    For hc = 1 To lastCol
        Dim hName As String: hName = Trim(CStr(wsData.Cells(HEADER_ROW, hc).Value))
        If hName = "ゾーン" Then zoneColIdx = hc
        If hName = "投入回数_予測" Then cntColIdx = hc
    Next hc
    If zoneColIdx = -1 Or cntColIdx = -1 Then
        MsgBox "「予測データ」シートに「ゾーン」または「投入回数_予測」の列が見つかりません。", vbExclamation
        Exit Sub
    End If

    Dim dictActualByLabel As Object: Set dictActualByLabel = CreateObject("Scripting.Dictionary")
    Dim r As Long
    For r = HEADER_ROW + 1 To lastRow
        Dim lbl As String: lbl = Trim(CStr(wsData.Cells(r, zoneColIdx).Value))
        If lbl <> "" Then
            dictActualByLabel(lbl) = dictActualByLabel(lbl) + Val(wsData.Cells(r, cntColIdx).Value)
        End If
    Next r

    Call BuildRatioChartSheet("予測グラフ", "号機別構成比(予測データ)", dictActualByLabel, dictTargetByLabel)

    MsgBox "「予測グラフ」を作成しました。", vbInformation
End Sub

Sub CreateActualRatioChart()
    Call EnsureRatioChartButtons

    Dim dictTargetByLabel As Object
    Call LoadTargetRatioByLabel(dictTargetByLabel)

    ' 機番→ゾーンラベルの対応づけは「予測データ」シート(ゾーン列・機番列を持つ)があればそこから作る。
    ' 無ければAB01～AB46(機番=ゾーン番号)のみ集計し、Cバラ・拡張(機番だけでは判別できない)は対象外にする
    Dim dictMachToZone As Object: Set dictMachToZone = CreateObject("Scripting.Dictionary")
    Dim hasZoneMap As Boolean: hasZoneMap = False
    Dim wsData As Worksheet
    On Error Resume Next
    Set wsData = ThisWorkbook.Sheets("予測データ")
    On Error GoTo 0
    If Not wsData Is Nothing Then
        Const HEADER_ROW2 As Long = 3
        Dim lastRow2 As Long: lastRow2 = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
        Dim lastCol2 As Long: lastCol2 = wsData.Cells(HEADER_ROW2, wsData.Columns.Count).End(xlToLeft).Column
        Dim machColIdx2 As Long: machColIdx2 = -1
        Dim zoneColIdx2 As Long: zoneColIdx2 = -1
        Dim hc2 As Long
        For hc2 = 1 To lastCol2
            Dim hName2 As String: hName2 = Trim(CStr(wsData.Cells(HEADER_ROW2, hc2).Value))
            If hName2 = "号機" Then machColIdx2 = hc2
            If hName2 = "ゾーン" Then zoneColIdx2 = hc2
        Next hc2
        If machColIdx2 > 0 And zoneColIdx2 > 0 Then
            Dim r2 As Long
            For r2 = HEADER_ROW2 + 1 To lastRow2
                Dim machKey As String: machKey = Trim(CStr(wsData.Cells(r2, machColIdx2).Value))
                If machKey <> "" And Not dictMachToZone.Exists(machKey) Then
                    dictMachToZone(machKey) = Trim(CStr(wsData.Cells(r2, zoneColIdx2).Value))
                End If
            Next r2
            hasZoneMap = (dictMachToZone.Count > 0)
        End If
    End If

    ' ファイル選択(複数選択・全ファイル形式。Module3と同じくS71実績ファイルを想定)
    Dim fd As Office.FileDialog
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

    ' E行(機番2桁+段2桁+列2桁の9文字区切り)からゾーンラベルを判定して集計する(除外設定・品コードは考慮しない生の実績)。
    ' 併せて、機番+段+列単位の実績も集計する(「日別実績」の更新に使う)
    Dim dictActualByLabel As Object: Set dictActualByLabel = CreateObject("Scripting.Dictionary")
    Dim dictLocationHits As Object: Set dictLocationHits = CreateObject("Scripting.Dictionary")
    Dim businessDate As Date: businessDate = DateSerial(1900, 1, 1)
    Dim skipMode As Boolean: skipMode = False
    Dim fIdx As Long, filePath As String, fileNo As Integer, textLine As String
    For fIdx = 1 To fd.SelectedItems.Count
        filePath = fd.SelectedItems(fIdx)
        fileNo = FreeFile
        Open filePath For Input As #fileNo
        Do While Not EOF(fileNo)
            Line Input #fileNo, textLine
            If Left(textLine, 1) = "B" And Len(textLine) >= 9 Then
                ' B行の2～9文字目(8桁)が集計日(YYYYMMDD)。「日別実績」の日付見出しに使う
                Dim bDateStr As String: bDateStr = Mid(textLine, 2, 8)
                If IsNumeric(bDateStr) Then
                    Dim bDate As Date
                    On Error Resume Next
                    bDate = DateSerial(CInt(Left(bDateStr, 4)), CInt(Mid(bDateStr, 5, 2)), CInt(Mid(bDateStr, 7, 2)))
                    On Error GoTo 0
                    If bDate > businessDate Then businessDate = bDate
                End If
            ElseIf Left(textLine, 1) = "H" Then
                ' H99999は在庫サマリー行。以降のE行(在庫全体の棚卸)は出荷実績としてカウントしない
                skipMode = (Mid(textLine, 2, 5) = "99999")
            ElseIf Left(textLine, 1) = "E" And Len(textLine) >= 10 And Not skipMode Then
                ' E行は13文字おきに最大3件のレコード(先頭9文字=機番2桁+段2桁+列2桁+3桁)が
                ' 詰められていることがある。末尾の余白がスペース埋め(新形式)またはゼロ埋め
                ' (旧形式)のいずれかで、ゼロ埋めの場合は余白がたまたま数字のみになり、
                ' 「機番0・段0・列0」という実在しないレコードが混じることがあるため、
                ' 機番0のレコードは明示的に除外する(下のmach>=1判定)
                Dim slotStart As Long, rec As String
                For slotStart = 2 To Len(textLine) - 8 Step 13
                    rec = Mid(textLine, slotStart, 9)
                    If Trim(rec) <> "" And Len(Trim(rec)) = 9 And IsNumeric(rec) Then
                        Dim mach As Long: mach = Val(Mid(rec, 1, 2))
                        Dim dan As Long: dan = Val(Mid(rec, 3, 2))
                        Dim retsu As Long: retsu = Val(Mid(rec, 5, 2))
                        Dim machKeyStr As String: machKeyStr = Format(mach, "00")
                        Dim zoneLbl As String: zoneLbl = ""
                        If hasZoneMap And dictMachToZone.Exists(machKeyStr) Then
                            zoneLbl = dictMachToZone(machKeyStr)
                        ElseIf mach >= 1 And mach <= MAX_AB_MACH Then
                            zoneLbl = "AB" & Format(mach, "00")
                        End If
                        If zoneLbl <> "" Then
                            dictActualByLabel(zoneLbl) = dictActualByLabel(zoneLbl) + 1
                        End If

                        If mach >= 1 Then
                            Dim locHitKey As String: locHitKey = machKeyStr & Format(dan, "00") & Format(retsu, "00")
                            dictLocationHits(locHitKey) = dictLocationHits(locHitKey) + 1
                        End If
                    End If
                Next slotStart
            End If
        Loop
        Close #fileNo
    Next fIdx

    Call BuildRatioChartSheet("実績グラフ", "号機別構成比(S71実績)", dictActualByLabel, dictTargetByLabel)

    If businessDate = DateSerial(1900, 1, 1) Then businessDate = Date ' B行から日付が読み取れなければ実行日を使う
    Call UpdateDailyLocationHistory(dictLocationHits, businessDate)
    Call UpdateDailyItemHistory(dictLocationHits, businessDate)

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    Dim noteMsg As String
    If Not hasZoneMap Then noteMsg = vbCrLf & "※「予測データ」シートが無いため、Cバラ(C01/C02)・拡張(X)は集計されていません(AB01～AB46のみ)。"
    MsgBox "「実績グラフ」「日別実績」を作成・更新しました。(" & fd.SelectedItems.Count & "ファイル読込)" & noteMsg, vbInformation
End Sub

' 「設定」シートの「■機番別目標構成比」(N:O列)を、ラベル文字列をキーにしたまま読み込む
' (Module3のdictTargetRatioは機番のみ数値キーに正規化されるため、C01・C02・Xを含むグラフ用にはこちらを使う)
Private Sub LoadTargetRatioByLabel(ByRef dictTargetByLabel As Object)
    Set dictTargetByLabel = CreateObject("Scripting.Dictionary")
    Call EnsureExclusionSettingsSheet

    Dim wsSet As Worksheet
    On Error Resume Next
    Set wsSet = ThisWorkbook.Sheets("設定")
    On Error GoTo 0
    If wsSet Is Nothing Then Exit Sub

    Dim lastN As Long: lastN = wsSet.Cells(wsSet.Rows.Count, "N").End(xlUp).Row
    Dim rN As Long
    For rN = 5 To lastN
        Dim lbl As String: lbl = Trim(CStr(wsSet.Cells(rN, 14).Value))
        If lbl <> "" And IsNumeric(wsSet.Cells(rN, 15).Value) Then
            ' 素の数値(1など)で入力されている場合は「AB01」形式に揃えてキーにする(このグラフはラベル単位で集計するため)
            If IsNumeric(lbl) Then lbl = "AB" & Format(CLng(lbl), "00")
            dictTargetByLabel(lbl) = CDbl(wsSet.Cells(rN, 15).Value) / 100
        End If
    Next rN
End Sub

' ゾーンラベル別の実績・目標比率から、データ表とグラフを持つシートを作成する(既存の同名シートは削除して作り直す)。
' AB01～AB46は機番ペア(ゾーン)ごとに奇数機番を正の値・偶数機番を負の値で持たせ、グラフ上で上下に分かれるようにする
' (負の値はセルの表示形式でマイナス符号を隠し、絶対値の比率として見せる)。
' Cバラ(C01・C02)・拡張(X)などAB以外のラベルは、データ表には奇数側の列に単独の項目として追加するが、
' グラフにはAB01～AB46のみを表示する(グラフの対象範囲をAB行までに限定する)
Private Sub BuildRatioChartSheet(ByVal sheetName As String, ByVal chartTitle As String, dictActualByLabel As Object, dictTargetByLabel As Object)
    On Error Resume Next
    ThisWorkbook.Sheets(sheetName).Delete
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
    wsOut.Name = sheetName

    ' 全体合計(AB01～AB46 + Cバラ・拡張等その他カテゴリすべて)を比率の分母にする
    Dim grandTotal As Double: grandTotal = 0
    Dim k As Variant
    For Each k In dictActualByLabel.Keys
        grandTotal = grandTotal + dictActualByLabel(k)
    Next k

    Dim hasTarget As Boolean: hasTarget = (dictTargetByLabel.Count > 0)

    wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(1, 5)).Merge
    wsOut.Cells(1, 1).Value = "【" & chartTitle & "】作成日時: " & Format(Now, "yyyy/mm/dd hh:mm")
    wsOut.Cells(1, 1).Font.Bold = True: wsOut.Cells(1, 1).Font.Size = 12
    wsOut.Cells(1, 1).HorizontalAlignment = xlLeft

    Dim headerArr As Variant
    headerArr = Array("ゾーン", "奇数比率", "偶数比率", "奇数目標比率", "偶数目標比率")
    wsOut.Range(wsOut.Cells(3, 1), wsOut.Cells(3, 5)).Value = headerArr
    wsOut.Range(wsOut.Cells(3, 1), wsOut.Cells(3, 5)).Interior.Color = RGB(220, 230, 255)
    wsOut.Range(wsOut.Cells(3, 1), wsOut.Cells(3, 5)).Font.Bold = True

    Dim r As Long: r = 3
    Dim maxAbsVal As Double: maxAbsVal = 0
    Dim zoneCount As Long: zoneCount = MAX_AB_MACH \ 2
    Dim z As Long
    For z = 1 To zoneCount
        r = r + 1
        Dim oddMach As Long: oddMach = z * 2 - 1
        Dim evenMach As Long: evenMach = z * 2
        Dim oddLabel As String: oddLabel = "AB" & Format(oddMach, "00")
        Dim evenLabel As String: evenLabel = "AB" & Format(evenMach, "00")
        wsOut.Cells(r, 1).Value = oddMach & "," & evenMach

        Dim oddVal As Double: oddVal = 0
        If dictActualByLabel.Exists(oddLabel) Then oddVal = dictActualByLabel(oddLabel)
        Dim evenVal As Double: evenVal = 0
        If dictActualByLabel.Exists(evenLabel) Then evenVal = dictActualByLabel(evenLabel)

        Dim oddRatio As Double: oddRatio = IIf(grandTotal > 0, oddVal / grandTotal, 0)
        Dim evenRatio As Double: evenRatio = IIf(grandTotal > 0, evenVal / grandTotal, 0)
        wsOut.Cells(r, 2).Value = oddRatio
        wsOut.Cells(r, 2).NumberFormat = "0.0%"
        wsOut.Cells(r, 3).Value = -evenRatio
        wsOut.Cells(r, 3).NumberFormat = "0.0%;0.0%"
        If oddRatio > maxAbsVal Then maxAbsVal = oddRatio
        If evenRatio > maxAbsVal Then maxAbsVal = evenRatio

        If hasTarget Then
            If dictTargetByLabel.Exists(oddLabel) Then
                Dim oddTarget As Double: oddTarget = dictTargetByLabel(oddLabel)
                wsOut.Cells(r, 4).Value = oddTarget
                wsOut.Cells(r, 4).NumberFormat = "0.0%"
                If oddTarget > maxAbsVal Then maxAbsVal = oddTarget
            End If
            If dictTargetByLabel.Exists(evenLabel) Then
                Dim evenTarget As Double: evenTarget = dictTargetByLabel(evenLabel)
                wsOut.Cells(r, 5).Value = -evenTarget
                wsOut.Cells(r, 5).NumberFormat = "0.0%;0.0%"
                If evenTarget > maxAbsVal Then maxAbsVal = evenTarget
            End If
        End If
    Next z

    ' グラフに含めるのはここまで(AB01～AB46)。この後ろに追加するCバラ・拡張の行はデータ表のみに含め、グラフの対象範囲には含めない
    Dim lastZoneRow As Long: lastZoneRow = r

    ' AB以外のラベル(Cバラ・拡張等)を、実績データまたは目標構成比のどちらかに存在するものすべて集めて、
    ' 奇数側の列に単独の上向き項目として追加する(奇数/偶数のペア概念が無いため偶数側は使わない)
    Dim dictExtraLabels As Object: Set dictExtraLabels = CreateObject("Scripting.Dictionary")
    For Each k In dictTargetByLabel.Keys
        If Not (CStr(k) Like "AB##") Then
            If Not dictExtraLabels.Exists(CStr(k)) Then dictExtraLabels.Add CStr(k), True
        End If
    Next k
    For Each k In dictActualByLabel.Keys
        If Not (CStr(k) Like "AB##") Then
            If Not dictExtraLabels.Exists(CStr(k)) Then dictExtraLabels.Add CStr(k), True
        End If
    Next k

    Dim ek As Variant
    For Each ek In dictExtraLabels.Keys
        r = r + 1
        wsOut.Cells(r, 1).Value = CStr(ek)
        Dim extraVal As Double: extraVal = 0
        If dictActualByLabel.Exists(ek) Then extraVal = dictActualByLabel(ek)
        wsOut.Cells(r, 2).Value = IIf(grandTotal > 0, extraVal / grandTotal, 0)
        wsOut.Cells(r, 2).NumberFormat = "0.0%"
        If hasTarget And dictTargetByLabel.Exists(ek) Then
            wsOut.Cells(r, 4).Value = dictTargetByLabel(ek)
            wsOut.Cells(r, 4).NumberFormat = "0.0%"
        End If
    Next ek

    Dim lastDataRow As Long: lastDataRow = r
    wsOut.Range(wsOut.Cells(3, 1), wsOut.Cells(lastDataRow, 5)).Columns.AutoFit

    ' グラフ(奇数機番を上向き・偶数機番を下向きの面グラフで表示。AB01～AB46のみが対象で、Cバラ・拡張は含めない。
    ' 目標構成比が入力されていれば、目標比率を折れ線で重ねて比較できるようにする)
    Dim srcCols As Long: srcCols = IIf(hasTarget, 5, 3)
    Dim srcRange As Range: Set srcRange = wsOut.Range(wsOut.Cells(3, 1), wsOut.Cells(lastZoneRow, srcCols))

    Dim chtObj As ChartObject
    Set chtObj = wsOut.ChartObjects.Add(wsOut.Cells(3, 7).Left, wsOut.Cells(3, 7).Top, 900, 380)
    With chtObj.Chart
        .SetSourceData Source:=srcRange
        .PlotBy = xlColumns
        .ChartType = xlArea
        .HasTitle = True
        .ChartTitle.Text = chartTitle
        .Axes(xlCategory).TickLabels.Font.Size = 7
        .HasLegend = True
        If hasTarget Then
            .SeriesCollection(3).ChartType = xlLineMarkers
            .SeriesCollection(4).ChartType = xlLineMarkers
        End If
    End With

    ' 縦軸の目盛りを固定スケールに統一する(既定は±5%・1%刻み。実データがこれを超える場合のみ1%単位で切り上げて広げる。
    ' 予測・実績のグラフ間、拠点間でも同じ基準にすることで見比べやすくする)
    Dim axisMax As Double: axisMax = AXIS_DEFAULT_MAX
    If maxAbsVal > axisMax Then axisMax = Application.WorksheetFunction.RoundUp(maxAbsVal / AXIS_UNIT, 0) * AXIS_UNIT
    With chtObj.Chart.Axes(xlValue)
        .MinimumScale = -axisMax
        .MaximumScale = axisMax
        .MajorUnit = AXIS_UNIT
        .TickLabels.NumberFormat = "0%;0%"
    End With
End Sub

' 「操作パネル」シートに構成比グラフのボタン(予測・実績)が無ければ追加する
' (既存のボタン・図形と重ならないよう、一番下にあるものの少し下に順番に配置する)
Sub EnsureRatioChartButtons()
    Dim wsPanel As Worksheet
    Call MigrateRenamedSheets
    Call MigrateRenamedButtons

    On Error Resume Next
    Set wsPanel = ThisWorkbook.Sheets("操作パネル")
    On Error GoTo 0
    If wsPanel Is Nothing Then Exit Sub

    Dim existing As Shape

    On Error Resume Next
    Set existing = wsPanel.Shapes("予測グラフボタン")
    On Error GoTo 0
    If existing Is Nothing Then
        Dim btn1 As Button
        Set btn1 = wsPanel.Buttons.Add(wsPanel.Range("B20").Left, wsPanel.Range("B20").Top, 220, 36)
        btn1.Name = "予測グラフボタン"
        btn1.OnAction = "CreateForecastRatioChart"
        btn1.Characters.Text = "予測グラフを作成"
        btn1.Font.Size = 12
        btn1.Font.Bold = True
    End If

    Set existing = Nothing
    On Error Resume Next
    Set existing = wsPanel.Shapes("実績グラフボタン")
    On Error GoTo 0
    If existing Is Nothing Then
        Dim btn2 As Button
        Set btn2 = wsPanel.Buttons.Add(wsPanel.Range("B20").Left, wsPanel.Range("B20").Top, 220, 36)
        btn2.Name = "実績グラフボタン"
        btn2.OnAction = "CreateActualRatioChart"
        btn2.Characters.Text = "実績グラフを作成"
        btn2.Font.Size = 12
        btn2.Font.Bold = True
    End If

    ' ボタンが下に伸び続けないよう、2列に並び替える(Module3の共通処理)
    Call LayoutPanelButtons
End Sub

' 「日別実績」シートを更新する(「予測データ」の行(品名コード・ロケ分類・品名・ロケーション・予測回数)を
' 土台にして、日別のS71実績ヒット数(dictLocationHits、機番+段+列キー)を1日分の列として追加する。
' 「予測データ」に無いロケーションでS71実績があった場合は、品名・ロケ分類を不明のまま行を追加する。
' 日別列はG～Pの最大10列。既に同じ日付の列があればそこを上書きし、10列すべて埋まっていれば
' 最も古い日(G列)を消して1列ずつ左に詰めてからP列に新しい日を書く。日別列の右端(Q列)には
' その行の日別実績の総計を書く。
' シートは「予測データ」の現在の内容で毎回作り直すが、既存の日別実績はロケーション単位で退避して引き継ぐため、
' 「予測データ」を再取込みして行が増減しても、過去の日別実績が失われることはない
Private Sub UpdateDailyLocationHistory(dictLocationHits As Object, ByVal businessDate As Date)
    Dim wsData As Worksheet
    Call MigrateRenamedSheets

    On Error Resume Next
    Set wsData = ThisWorkbook.Sheets("予測データ")
    On Error GoTo 0
    If wsData Is Nothing Then Exit Sub ' ロケ分類・品名等の元データが無ければ更新しない

    Const HEADER_ROW As Long = 3
    Dim lastRow As Long: lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    Dim lastCol As Long: lastCol = wsData.Cells(HEADER_ROW, wsData.Columns.Count).End(xlToLeft).Column

    Dim machColIdx As Long: machColIdx = -1
    Dim danColIdx As Long: danColIdx = -1
    Dim colColIdx As Long: colColIdx = -1
    Dim itemCodeColIdx As Long: itemCodeColIdx = -1
    Dim itemNameColIdx As Long: itemNameColIdx = -1
    Dim locClassColIdx As Long: locClassColIdx = -1
    Dim forecastColIdx As Long: forecastColIdx = -1
    Dim hc As Long
    For hc = 1 To lastCol
        Dim hName As String: hName = Trim(CStr(wsData.Cells(HEADER_ROW, hc).Value))
        If hName = "号機" Then machColIdx = hc
        If hName = "段" Then danColIdx = hc
        If hName = "列" Then colColIdx = hc
        If hName = "品名コード" Then itemCodeColIdx = hc
        If hName = "品名" Then itemNameColIdx = hc
        If hName = "ロケ分類" Then locClassColIdx = hc
        If hName = "投入回数_予測" Then forecastColIdx = hc
    Next hc
    If machColIdx = -1 Or danColIdx = -1 Or colColIdx = -1 Or itemCodeColIdx = -1 Then Exit Sub ' 必要な列が無ければ更新しない

    Const DATE_COL_FIRST As Long = 7  ' G列
    Const DATE_COL_LAST As Long = 16  ' P列(最大10列)
    Const TOTAL_COL As Long = 17      ' Q列(日別実績の総計)
    Dim histSheetName As String: histSheetName = "日別実績"

    ' 既存シートがあれば、日付見出しと日別実績(ロケーションキー→日付ごとの値)を退避しておく
    Dim wsOld As Worksheet
    On Error Resume Next
    Set wsOld = ThisWorkbook.Sheets(histSheetName)
    On Error GoTo 0

    ' 列の並び順は読み込んだ順ではなく、実際の日付の新旧で決める(古い日付ほど左、
    ' 最大10件を超える分は最も古い日付から削られる。同じ日付が既にあれば上書きする)
    Dim dictOldHistory As Object: Set dictOldHistory = CreateObject("Scripting.Dictionary") ' ロケーションキー→Dictionary(日付シリアル値→値)
    Dim finalDates() As Long
    Call BuildSortedHistoryDates(wsOld, 4, DATE_COL_FIRST, DATE_COL_LAST, businessDate, dictOldHistory, finalDates)
    Dim businessDateSerial As Long: businessDateSerial = CLng(businessDate)

    ' シートを「予測データ」の現在の内容で作り直す
    On Error Resume Next
    ThisWorkbook.Sheets(histSheetName).Delete
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
    wsOut.Name = histSheetName

    wsOut.Columns("A:A").NumberFormat = "@" ' コード(先頭ゼロ落ち防止)
    wsOut.Columns("D:D").NumberFormat = "00\-00\-00" ' ロケーション(機番-段-列表示)
    wsOut.Columns("F:F").NumberFormat = "@"

    wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(1, 6)).Value = Array("コード", "ロケ分類", "品名", "ロケーション", "予測回数", "行ラベル")
    Dim dNo As Long
    For dNo = 1 To UBound(finalDates)
        wsOut.Cells(1, DATE_COL_FIRST + dNo - 1).Value = FormatHistoryDateHeader(CDate(finalDates(dNo)))
    Next dNo
    wsOut.Cells(1, TOTAL_COL).Value = "総計"
    wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(1, TOTAL_COL)).Font.Bold = True
    wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(1, TOTAL_COL)).Interior.Color = RGB(220, 230, 255)

    Dim outRow As Long: outRow = 1
    Dim dictCoveredLocKeys As Object: Set dictCoveredLocKeys = CreateObject("Scripting.Dictionary") ' 「予測データ」でカバー済みのロケーション(機番+段+列キー)
    Dim r As Long
    For r = HEADER_ROW + 1 To lastRow
        If IsNumeric(wsData.Cells(r, machColIdx).Value) And IsNumeric(wsData.Cells(r, danColIdx).Value) And IsNumeric(wsData.Cells(r, colColIdx).Value) Then
            Dim mach As Long: mach = CLng(wsData.Cells(r, machColIdx).Value)
            Dim dan As Long: dan = CLng(wsData.Cells(r, danColIdx).Value)
            Dim colv As Long: colv = CLng(wsData.Cells(r, colColIdx).Value)
            Dim itemCode As String: itemCode = Trim(CStr(wsData.Cells(r, itemCodeColIdx).Value))

            outRow = outRow + 1
            wsOut.Cells(outRow, 1).Value = itemCode
            wsOut.Cells(outRow, 2).Value = IIf(locClassColIdx > 0, Trim(CStr(wsData.Cells(r, locClassColIdx).Value)), "")
            wsOut.Cells(outRow, 3).Value = IIf(itemNameColIdx > 0, Trim(CStr(wsData.Cells(r, itemNameColIdx).Value)), "")
            wsOut.Cells(outRow, 4).Value = mach * 10000 + dan * 100 + colv
            wsOut.Cells(outRow, 5).Value = IIf(forecastColIdx > 0, Val(wsData.Cells(r, forecastColIdx).Value), 0)
            wsOut.Cells(outRow, 6).Value = itemCode

            Dim rLocKey As String: rLocKey = CStr(mach * 10000 + dan * 100 + colv)
            Dim eLocKey As String: eLocKey = Format(mach, "00") & Format(dan, "00") & Format(colv, "00")
            If Not dictCoveredLocKeys.Exists(eLocKey) Then dictCoveredLocKeys.Add eLocKey, True

            Call WriteHistoryValues(wsOut, outRow, eLocKey, rLocKey, dictLocationHits, dictOldHistory, finalDates, DATE_COL_FIRST, TOTAL_COL, businessDateSerial)
        End If
    Next r

    ' 「予測データ」に無いロケーションでS71実績があった場合は、品名・ロケ分類を不明のまま行を追加する
    Dim extraKey As Variant
    For Each extraKey In dictLocationHits.Keys
        Dim exKeyStr As String: exKeyStr = CStr(extraKey)
        If Not dictCoveredLocKeys.Exists(exKeyStr) And Len(exKeyStr) = 6 Then
            Dim exMach As Long: exMach = CLng(Mid(exKeyStr, 1, 2))
            Dim exDan As Long: exDan = CLng(Mid(exKeyStr, 3, 2))
            Dim exCol As Long: exCol = CLng(Mid(exKeyStr, 5, 2))
            Dim exRLocKey As String: exRLocKey = CStr(exMach * 10000 + exDan * 100 + exCol)

            outRow = outRow + 1
            wsOut.Cells(outRow, 1).Value = ""
            wsOut.Cells(outRow, 2).Value = ""
            wsOut.Cells(outRow, 3).Value = "(品名不明)"
            wsOut.Cells(outRow, 4).Value = exMach * 10000 + exDan * 100 + exCol
            wsOut.Cells(outRow, 5).Value = 0
            wsOut.Cells(outRow, 6).Value = ""

            Call WriteHistoryValues(wsOut, outRow, exKeyStr, exRLocKey, dictLocationHits, dictOldHistory, finalDates, DATE_COL_FIRST, TOTAL_COL, businessDateSerial)
        End If
    Next extraKey

    wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(outRow, TOTAL_COL)).Columns.AutoFit
    wsOut.Range("A1").AutoFilter
End Sub

' 「7/31(木)」のような表記の日付見出しを実際の日付に変換する。パースできなければEmptyを返す。
' ローリング10営業日分の範囲でしか使わないため、今日から半年以上離れる場合のみ前後1年ずらして
' 年またぎを補正する
Function ParseHistoryDateHeader(ByVal rawVal As Variant) As Variant
    ParseHistoryDateHeader = Empty
    If IsEmpty(rawVal) Then Exit Function
    Dim s As String: s = Trim(CStr(rawVal))
    If s = "" Then Exit Function
    Dim pPos As Long: pPos = InStr(s, "(")
    If pPos > 1 Then s = Left(s, pPos - 1)
    If Not IsDate(s) Then Exit Function
    Dim d As Date: d = CDate(s)
    Do While d > Date + 200
        d = DateAdd("yyyy", -1, d)
    Loop
    Do While d < Date - 200
        d = DateAdd("yyyy", 1, d)
    Loop
    ParseHistoryDateHeader = d
End Function

' arr(lo..hi)を昇順(古い日付=小さい値が先)に並べ替える(QuickSort)
Sub SortLongArrayAsc(arr() As Long, ByVal lo As Long, ByVal hi As Long)
    If lo >= hi Then Exit Sub
    Dim pivot As Long: pivot = arr((lo + hi) \ 2)
    Dim i As Long: i = lo
    Dim j As Long: j = hi
    Do While i <= j
        Do While arr(i) < pivot
            i = i + 1
        Loop
        Do While arr(j) > pivot
            j = j - 1
        Loop
        If i <= j Then
            Dim tmp As Long: tmp = arr(i)
            arr(i) = arr(j)
            arr(j) = tmp
            i = i + 1
            j = j - 1
        End If
    Loop
    If lo < j Then Call SortLongArrayAsc(arr, lo, j)
    If i < hi Then Call SortLongArrayAsc(arr, i, hi)
End Sub

' 既存シート(wsOldOut、無ければNothing)のkeyColIdx列をキーとして、日付見出し・行データを
' dictOldHistory(キー→Dictionary(日付シリアル値→値))へ退避し、今回の実施日を加えたうえで、
' 列に並べる日付一覧(finalDates、古い日付順・最大dateColLast-dateColFirst+1件)を決定する。
' 読み込み順に関わらず実際の日付の新旧だけで並び順が決まり、同じ日付が既にあれば上書き、
' 上限を超える分は最も古い日付から削られる
Sub BuildSortedHistoryDates(ByVal wsOldOut As Worksheet, ByVal keyColIdx As Long, ByVal dateColFirst As Long, ByVal dateColLast As Long, ByVal businessDate As Date, dictOldHistory As Object, finalDates() As Long)
    Dim existingDates As Object: Set existingDates = CreateObject("Scripting.Dictionary")

    If Not wsOldOut Is Nothing Then
        Dim oldColDate() As Variant
        ReDim oldColDate(dateColFirst To dateColLast)
        Dim dc As Long
        For dc = dateColFirst To dateColLast
            Dim parsedDate As Variant: parsedDate = ParseHistoryDateHeader(wsOldOut.Cells(1, dc).Value)
            oldColDate(dc) = parsedDate
            If Not IsEmpty(parsedDate) Then existingDates(CLng(CDate(parsedDate))) = True
        Next dc

        Dim lastOutRow As Long: lastOutRow = wsOldOut.Cells(wsOldOut.Rows.Count, 1).End(xlUp).Row
        If lastOutRow >= 2 Then
            Dim orow As Long
            For orow = 2 To lastOutRow
                Dim oKey As String: oKey = Trim(CStr(wsOldOut.Cells(orow, keyColIdx).Value))
                If oKey <> "" And Not dictOldHistory.Exists(oKey) Then
                    Dim dateValMap As Object: Set dateValMap = CreateObject("Scripting.Dictionary")
                    For dc = dateColFirst To dateColLast
                        If Not IsEmpty(oldColDate(dc)) Then
                            Dim cv As Variant: cv = wsOldOut.Cells(orow, dc).Value
                            If Not IsEmpty(cv) And cv <> "" Then
                                dateValMap(CLng(CDate(oldColDate(dc)))) = cv
                            End If
                        End If
                    Next dc
                    dictOldHistory.Add oKey, dateValMap
                End If
            Next orow
        End If
    End If

    existingDates(CLng(businessDate)) = True

    Dim nd As Long: nd = existingDates.Count
    Dim allDatesArr() As Long
    ReDim allDatesArr(1 To nd)
    Dim di As Long: di = 0
    Dim dk As Variant
    For Each dk In existingDates.Keys
        di = di + 1
        allDatesArr(di) = CLng(dk)
    Next dk
    Call SortLongArrayAsc(allDatesArr, 1, nd)

    Dim maxCols As Long: maxCols = dateColLast - dateColFirst + 1
    Dim keepFrom As Long: keepFrom = 1
    If nd > maxCols Then keepFrom = nd - maxCols + 1

    Dim finalCount As Long: finalCount = nd - keepFrom + 1
    ReDim finalDates(1 To finalCount)
    Dim fi As Long
    For fi = 1 To finalCount
        finalDates(fi) = allDatesArr(keepFrom + fi - 1)
    Next fi
End Sub

' finalDates(1..N)の各日付について、hitKeyの今日の実績(dictTodayHits)、または
' persistKeyの過去実績(dictOldHistory)から値を求めて列に書き込む
Sub WriteHistoryValues(ByVal wsOut As Worksheet, ByVal outRow As Long, ByVal hitKey As String, ByVal persistKey As String, dictTodayHits As Object, dictOldHistory As Object, finalDates() As Long, ByVal dateColFirst As Long, ByVal totalCol As Long, ByVal businessDateSerial As Long)
    Dim hasOldMap As Boolean: hasOldMap = dictOldHistory.Exists(persistKey)
    Dim oldMap As Object
    If hasOldMap Then Set oldMap = dictOldHistory(persistKey)

    Dim totalVal As Double: totalVal = 0
    Dim n As Long: n = UBound(finalDates)
    Dim i As Long
    For i = 1 To n
        Dim colIdx As Long: colIdx = dateColFirst + i - 1
        Dim cellVal As Variant: cellVal = Empty
        If finalDates(i) = businessDateSerial Then
            Dim hitVal As Double: hitVal = 0
            If dictTodayHits.Exists(hitKey) Then hitVal = dictTodayHits(hitKey)
            cellVal = hitVal
        ElseIf hasOldMap Then
            If oldMap.Exists(finalDates(i)) Then cellVal = oldMap(finalDates(i))
        End If
        If Not IsEmpty(cellVal) Then
            wsOut.Cells(outRow, colIdx).Value = cellVal
            If IsNumeric(cellVal) Then totalVal = totalVal + CDbl(cellVal)
        End If
    Next i
    wsOut.Cells(outRow, totalCol).Value = totalVal
End Sub

' 「品名実績」シートを更新する(機番-段-列→品名コード対応を使い、
' その日のロケーション別S71実績を品名コード単位で合算する。同じ品名コードが
' 複数ロケーションにまたがっていれば合計する)。
' S71実績はロケーション単位でしか記録されないため、「日別実績」はロケーションに
' 実績が紐づき、商品が移動すると過去実績はそのロケーションに残ってしまう。
' このシートは品名コードを単位に実績を蓄積するため、商品が別のロケーションへ
' 移動しても実績はその商品についてくる(移動先の実績にすり替わらない)。
' ただし「予測データ」に無いロケーションのS71実績(品名不明分)は品名コードに
' 割り当てられないため、このシートの合計には含まれない(「日別実績」の合計とは
' 一致しないことがある)。
' 列の並び順は読み込んだ順ではなく、実際の日付の新旧で決める(古い日付ほど左)。
' シートは「予測データ」の現在の内容で毎回作り直すが、既存の実績は品名コード単位で
' 退避して引き継ぐため、商品が入れ替わっても過去実績が失われることはない
Function FormatHistoryDateHeader(ByVal d As Date) As String
    Dim wdNames As Variant: wdNames = Array("日", "月", "火", "水", "木", "金", "土")
    FormatHistoryDateHeader = Format(d, "m/d") & "(" & wdNames(Weekday(d) - 1) & ")"
End Function

Private Sub UpdateDailyItemHistory(dictLocationHits As Object, ByVal businessDate As Date)
    Dim wsData As Worksheet
    On Error Resume Next
    Set wsData = ThisWorkbook.Sheets("予測データ")
    On Error GoTo 0
    If wsData Is Nothing Then Exit Sub

    Const HEADER_ROW As Long = 3
    Dim lastRow As Long: lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    Dim lastCol As Long: lastCol = wsData.Cells(HEADER_ROW, wsData.Columns.Count).End(xlToLeft).Column

    Dim machColIdx As Long: machColIdx = -1
    Dim danColIdx As Long: danColIdx = -1
    Dim colColIdx As Long: colColIdx = -1
    Dim itemCodeColIdx As Long: itemCodeColIdx = -1
    Dim itemNameColIdx As Long: itemNameColIdx = -1
    Dim locClassColIdx As Long: locClassColIdx = -1
    Dim forecastColIdx As Long: forecastColIdx = -1
    Dim hc As Long
    For hc = 1 To lastCol
        Dim hName As String: hName = Trim(CStr(wsData.Cells(HEADER_ROW, hc).Value))
        If hName = "号機" Then machColIdx = hc
        If hName = "段" Then danColIdx = hc
        If hName = "列" Then colColIdx = hc
        If hName = "品名コード" Then itemCodeColIdx = hc
        If hName = "品名" Then itemNameColIdx = hc
        If hName = "ロケ分類" Then locClassColIdx = hc
        If hName = "投入回数_予測" Then forecastColIdx = hc
    Next hc
    If machColIdx = -1 Or danColIdx = -1 Or colColIdx = -1 Or itemCodeColIdx = -1 Then Exit Sub

    Const DATE_COL_FIRST As Long = 7  ' G列
    Const DATE_COL_LAST As Long = 16  ' P列(最大10列)
    Const TOTAL_COL As Long = 17      ' Q列(実績の総計)
    Dim histSheetName As String: histSheetName = "品名実績"

    ' 既存シートがあれば、日付見出しと実績(品名コード→日付ごとの値)を退避しておく
    Dim wsOld As Worksheet
    On Error Resume Next
    Set wsOld = ThisWorkbook.Sheets(histSheetName)
    On Error GoTo 0

    Dim dictOldHistory As Object: Set dictOldHistory = CreateObject("Scripting.Dictionary") ' 品名コード→Dictionary(日付シリアル値→値)
    Dim finalDates() As Long
    Call BuildSortedHistoryDates(wsOld, 1, DATE_COL_FIRST, DATE_COL_LAST, businessDate, dictOldHistory, finalDates)
    Dim businessDateSerial As Long: businessDateSerial = CLng(businessDate)

    ' 「予測データ」の現在の機番-段-列→品名コード対応で、今日のロケーション別実績を
    ' 品名コード単位に合算する(同じ品名コードが複数ロケーションにあれば合計する)
    Dim dictItemHitsToday As Object: Set dictItemHitsToday = CreateObject("Scripting.Dictionary")
    Dim dictItemLocCount As Object: Set dictItemLocCount = CreateObject("Scripting.Dictionary")
    Dim dictItemForecastSum As Object: Set dictItemForecastSum = CreateObject("Scripting.Dictionary")
    Dim dictItemName As Object: Set dictItemName = CreateObject("Scripting.Dictionary")
    Dim dictItemLocClass As Object: Set dictItemLocClass = CreateObject("Scripting.Dictionary")

    Dim r As Long
    For r = HEADER_ROW + 1 To lastRow
        If IsNumeric(wsData.Cells(r, machColIdx).Value) And IsNumeric(wsData.Cells(r, danColIdx).Value) And IsNumeric(wsData.Cells(r, colColIdx).Value) Then
            Dim mach As Long: mach = CLng(wsData.Cells(r, machColIdx).Value)
            Dim dan As Long: dan = CLng(wsData.Cells(r, danColIdx).Value)
            Dim colv As Long: colv = CLng(wsData.Cells(r, colColIdx).Value)
            Dim itemCode As String: itemCode = Trim(CStr(wsData.Cells(r, itemCodeColIdx).Value))
            If itemCode <> "" Then
                Dim eLocKey As String: eLocKey = Format(mach, "00") & Format(dan, "00") & Format(colv, "00")
                Dim hitVal As Double: hitVal = 0
                If dictLocationHits.Exists(eLocKey) Then hitVal = dictLocationHits(eLocKey)

                If Not dictItemHitsToday.Exists(itemCode) Then
                    dictItemHitsToday.Add itemCode, 0
                    dictItemLocCount.Add itemCode, 0
                    dictItemForecastSum.Add itemCode, 0
                    dictItemName.Add itemCode, IIf(itemNameColIdx > 0, Trim(CStr(wsData.Cells(r, itemNameColIdx).Value)), "")
                    dictItemLocClass.Add itemCode, IIf(locClassColIdx > 0, Trim(CStr(wsData.Cells(r, locClassColIdx).Value)), "")
                End If
                dictItemHitsToday(itemCode) = dictItemHitsToday(itemCode) + hitVal
                dictItemLocCount(itemCode) = dictItemLocCount(itemCode) + 1
                dictItemForecastSum(itemCode) = dictItemForecastSum(itemCode) + IIf(forecastColIdx > 0, Val(wsData.Cells(r, forecastColIdx).Value), 0)
            End If
        End If
    Next r

    ' シートを作り直す
    On Error Resume Next
    ThisWorkbook.Sheets(histSheetName).Delete
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
    wsOut.Name = histSheetName

    wsOut.Columns("A:A").NumberFormat = "@" ' コード(先頭ゼロ落ち防止)

    wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(1, 6)).Value = Array("コード", "品名", "ロケ分類", "号機数", "予測回数", "")
    Dim dNo As Long
    For dNo = 1 To UBound(finalDates)
        wsOut.Cells(1, DATE_COL_FIRST + dNo - 1).Value = FormatHistoryDateHeader(CDate(finalDates(dNo)))
    Next dNo
    wsOut.Cells(1, TOTAL_COL).Value = "総計"
    wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(1, TOTAL_COL)).Font.Bold = True
    wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(1, TOTAL_COL)).Interior.Color = RGB(220, 230, 255)

    Dim outRow As Long: outRow = 1
    Dim itemKeyV As Variant
    For Each itemKeyV In dictItemHitsToday.Keys
        Dim ik As String: ik = CStr(itemKeyV)
        outRow = outRow + 1
        wsOut.Cells(outRow, 1).Value = ik
        wsOut.Cells(outRow, 2).Value = dictItemName(ik)
        wsOut.Cells(outRow, 3).Value = dictItemLocClass(ik)
        wsOut.Cells(outRow, 4).Value = dictItemLocCount(ik)
        wsOut.Cells(outRow, 5).Value = dictItemForecastSum(ik)

        Call WriteHistoryValues(wsOut, outRow, ik, ik, dictItemHitsToday, dictOldHistory, finalDates, DATE_COL_FIRST, TOTAL_COL, businessDateSerial)
    Next itemKeyV

    wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(outRow, TOTAL_COL)).Columns.AutoFit
    wsOut.Range("A1").AutoFilter
End Sub
