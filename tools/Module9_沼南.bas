Attribute VB_Name = "Module9_沼南"
Option Explicit

' ----------------------------------------------------
' 構成比グラフの作成
' 「予測構成比グラフ」:「予測データ」シート(Module8で取込済み)の投入回数_予測をゾーン別に集計してグラフ化する
' 「実績構成比グラフ」:ピッキング実績ファイル(S71)をダイアログで選択し、ゾーン別ヒット数を集計してグラフ化する
' レイアウトは「設定」シートの「■号機別目標構成比」と同じ考え方で、ABブロック内の号機は機番ペア(ゾーン)ごとに
' 奇数号機を上向き・偶数号機を下向きに表示し(ABブロックの境界はGetMachPairLabelで判定するため、
' 沼南の非連続なブロック配置(1～30、37～50)にも正しく対応する)、Cバラ(C01・C02)・拡張(X)は
' 上向きの単独項目として、必ずC01・C02・Xの順で追加する(実績側は号機の範囲で判定:ABブロック外かつ
' 61～73はC01、81～93はC02、それ以外はXという沼南の実際のラック配置に基づく固定ルール)。
' 「設定」シートに目標構成比が入力されていれば、実績/予測の比率と並べて目標比率も折れ線で比較できるようにする。
' データ表にはCバラ・拡張も含めるが、グラフにはABブロック内の号機のみを表示する(Cバラ・拡張はグラフの対象外)。
' グラフの縦軸目盛りは予測・実績のグラフ間、また石狩のグラフとも見比べやすいよう固定スケール(既定は±5%・1%刻み)にし、
' 実データがそれを超える場合のみ切り上げて広げる。
' 既存の同名シートは削除してから作り直すため、再実行すると内容が更新される
' ----------------------------------------------------

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
    Dim abBlockFrom() As Long, abBlockTo() As Long, abBlockCount As Long
    Call LoadSettingsForChart(dictTargetByLabel, abBlockFrom, abBlockTo, abBlockCount)

    ' 「予測データ」シートは1行目=取込情報、3行目=見出し(Module8の出力形式)。
    ' ゾーン列(AB01～AB50・C01・C02・Xなどのラベル)・投入回数_予測列を見出し名から探す
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

    Call BuildRatioChartSheet("予測構成比グラフ", "号機別構成比(予測データ)", dictActualByLabel, dictTargetByLabel, abBlockFrom, abBlockTo, abBlockCount)

    MsgBox "「予測構成比グラフ」を作成しました。", vbInformation
End Sub

Sub CreateActualRatioChart()
    Call EnsureRatioChartButtons

    Dim dictTargetByLabel As Object
    Dim abBlockFrom() As Long, abBlockTo() As Long, abBlockCount As Long
    Call LoadSettingsForChart(dictTargetByLabel, abBlockFrom, abBlockTo, abBlockCount)

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

    ' E行(号機2桁+段2桁+列2桁の9文字区切り)からゾーンラベルを判定して集計する(除外設定・品コードは考慮しない生の実績)。
    ' ABブロック内はGetMachPairLabel等と同じくIsInABBlockで判定し、ブロック外は沼南の実際のラック配置
    ' (Cバラ01=61～73、Cバラ02=81～93、それ以外は拡張X)を号機の範囲で直接判定する
    Dim dictActualByLabel As Object: Set dictActualByLabel = CreateObject("Scripting.Dictionary")
    Dim fIdx As Long, filePath As String, fileNo As Integer, textLine As String
    For fIdx = 1 To fd.SelectedItems.Count
        filePath = fd.SelectedItems(fIdx)
        fileNo = FreeFile
        Open filePath For Input As #fileNo
        Do While Not EOF(fileNo)
            Line Input #fileNo, textLine
            If Left(textLine, 1) = "E" And Len(textLine) >= 10 Then
                Dim slotStart As Long
                For slotStart = 2 To Len(textLine) - 8 Step 13
                    Dim rec As String: rec = Mid(textLine, slotStart, 9)
                    If Trim(rec) <> "" And Len(Trim(rec)) = 9 And IsNumeric(rec) Then
                        Dim mach As Long: mach = Val(Mid(rec, 1, 2))
                        If mach >= 1 Then
                            Dim zoneLbl As String
                            If IsInABBlock(CInt(mach), abBlockFrom, abBlockTo, abBlockCount) Then
                                zoneLbl = "AB" & Format(mach, "00")
                            ElseIf mach >= 61 And mach <= 73 Then
                                zoneLbl = "C01"
                            ElseIf mach >= 81 And mach <= 93 Then
                                zoneLbl = "C02"
                            Else
                                zoneLbl = "X"
                            End If
                            dictActualByLabel(zoneLbl) = dictActualByLabel(zoneLbl) + 1
                        End If
                    End If
                Next slotStart
            End If
        Loop
        Close #fileNo
    Next fIdx

    Call BuildRatioChartSheet("実績構成比グラフ", "号機別構成比(S71実績)", dictActualByLabel, dictTargetByLabel, abBlockFrom, abBlockTo, abBlockCount)

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "「実績構成比グラフ」を作成しました。(" & fd.SelectedItems.Count & "ファイル読込)", vbInformation
End Sub

' 「設定」シートからABブロック構成と「■号機別目標構成比」(Q:R列)をラベル文字列をキーにしたまま読み込む
' (Module3のdictTargetRatioは号機のみ数値キーに正規化されるため、C01・C02・Xを含むグラフ用にはこちらを使う)
Private Sub LoadSettingsForChart(ByRef dictTargetByLabel As Object, ByRef abBlockFrom() As Long, ByRef abBlockTo() As Long, ByRef abBlockCount As Long)
    Set dictTargetByLabel = CreateObject("Scripting.Dictionary")
    Call EnsureExclusionSettingsSheet

    ' ABブロック構成(N:O列)はModule3のLoadExclusionSettingsをそのまま流用して取得する
    Dim dictExcludedMach As Object: Set dictExcludedMach = CreateObject("Scripting.Dictionary")
    Dim locMach() As Long, locDanFrom() As Long, locDanTo() As Long, locColFrom() As Long, locColTo() As Long
    Dim locCount As Long
    Dim dictExcludedItemCode As Object: Set dictExcludedItemCode = CreateObject("Scripting.Dictionary")
    Dim ratioSheetName As String, maxSwapRows As Long, abSlotCount As Long
    Dim dictTargetRatioDummy As Object: Set dictTargetRatioDummy = CreateObject("Scripting.Dictionary")
    Call LoadExclusionSettings(dictExcludedMach, locMach, locDanFrom, locDanTo, locColFrom, locColTo, locCount, dictExcludedItemCode, ratioSheetName, maxSwapRows, abSlotCount, abBlockFrom, abBlockTo, abBlockCount, dictTargetRatioDummy)

    ' 「■号機別目標構成比」(Q:R列)はラベル文字列のまま読み込み直す(C01・C02・Xを保持するため)
    Dim wsSet As Worksheet
    On Error Resume Next
    Set wsSet = ThisWorkbook.Sheets("設定")
    On Error GoTo 0
    If wsSet Is Nothing Then Exit Sub

    Dim lastQ As Long: lastQ = wsSet.Cells(wsSet.Rows.Count, "Q").End(xlUp).Row
    Dim rQ As Long
    For rQ = 5 To lastQ
        Dim lbl As String: lbl = Trim(CStr(wsSet.Cells(rQ, 17).Value))
        If lbl <> "" And IsNumeric(wsSet.Cells(rQ, 18).Value) Then
            ' 素の数値(1など)で入力されている場合は「AB01」形式に揃えてキーにする(このグラフはラベル単位で集計するため)
            If IsNumeric(lbl) Then lbl = "AB" & Format(CLng(lbl), "00")
            dictTargetByLabel(lbl) = CDbl(wsSet.Cells(rQ, 18).Value) / 100
        End If
    Next rQ
End Sub

' ゾーンラベル別の実績・目標比率から、データ表とグラフを持つシートを作成する(既存の同名シートは削除して作り直す)。
' ABブロック内の号機は機番ペア(ゾーン)ごとに奇数号機を正の値・偶数号機を負の値で持たせ、グラフ上で上下に分かれるようにする
' (負の値はセルの表示形式でマイナス符号を隠し、絶対値の比率として見せる)。ゾーンの境界はGetMachPairLabelで
' 判定するため、沼南の非連続なABブロック配置(1～30、37～50)にも正しく対応する。
' Cバラ(C01・C02)・拡張(X)などAB以外のラベルは、データ表には奇数側の列に単独の項目として追加するが、
' グラフにはABブロック内の号機のみを表示する(グラフの対象範囲をABゾーン行までに限定する)
Private Sub BuildRatioChartSheet(ByVal sheetName As String, ByVal chartTitle As String, dictActualByLabel As Object, dictTargetByLabel As Object, abBlockFrom() As Long, abBlockTo() As Long, ByVal abBlockCount As Long)
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

    ' 全体合計(ABブロック内の号機 + Cバラ・拡張等その他カテゴリすべて)を比率の分母にする
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
    Dim zoneCount As Long: zoneCount = GetTotalZoneCount(abBlockFrom, abBlockTo, abBlockCount)
    Dim z As Long
    For z = 1 To zoneCount
        r = r + 1
        Dim pairLbl As String: pairLbl = GetMachPairLabel(z, abBlockFrom, abBlockTo, abBlockCount)
        Dim pairParts() As String: pairParts = Split(pairLbl, "&")
        Dim oddMach As Long: oddMach = CLng(pairParts(0))
        Dim evenMach As Long: evenMach = CLng(pairParts(1))
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

    ' グラフに含めるのはここまで(ABブロック内の号機)。この後ろに追加するCバラ・拡張の行はデータ表のみに含め、グラフの対象範囲には含めない
    Dim lastZoneRow As Long: lastZoneRow = r

    ' AB以外のラベル(Cバラ01・Cバラ02・拡張X)を、必ずこの順番で奇数側の列に単独の上向き項目として追加する
    ' (奇数/偶数のペア概念が無いため偶数側は使わない。沼南の実際のラック配置に合わせた固定順)
    Dim extraLabels As Variant: extraLabels = Array("C01", "C02", "X")
    Dim ei As Long
    For ei = LBound(extraLabels) To UBound(extraLabels)
        Dim ek As String: ek = CStr(extraLabels(ei))
        r = r + 1
        wsOut.Cells(r, 1).Value = ek
        Dim extraVal As Double: extraVal = 0
        If dictActualByLabel.Exists(ek) Then extraVal = dictActualByLabel(ek)
        wsOut.Cells(r, 2).Value = IIf(grandTotal > 0, extraVal / grandTotal, 0)
        wsOut.Cells(r, 2).NumberFormat = "0.0%"
        If hasTarget And dictTargetByLabel.Exists(ek) Then
            wsOut.Cells(r, 4).Value = dictTargetByLabel(ek)
            wsOut.Cells(r, 4).NumberFormat = "0.0%"
        End If
    Next ei

    Dim lastDataRow As Long: lastDataRow = r
    wsOut.Range(wsOut.Cells(3, 1), wsOut.Cells(lastDataRow, 5)).Columns.AutoFit

    ' グラフ(奇数号機を上向き・偶数号機を下向きの面グラフで表示。ABブロック内の号機のみが対象で、Cバラ・拡張は含めない。
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
    ' 予測・実績のグラフ間、石狩のグラフとも同じ基準にすることで見比べやすくする)
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
    On Error Resume Next
    Set wsPanel = ThisWorkbook.Sheets("操作パネル")
    On Error GoTo 0
    If wsPanel Is Nothing Then Exit Sub

    Dim maxBottom As Double: maxBottom = 0
    Dim shp As Shape
    For Each shp In wsPanel.Shapes
        If shp.Top + shp.Height > maxBottom Then maxBottom = shp.Top + shp.Height
    Next shp
    If maxBottom = 0 Then maxBottom = wsPanel.Range("B20").Top

    Dim existing As Shape

    On Error Resume Next
    Set existing = wsPanel.Shapes("予測構成比グラフボタン")
    On Error GoTo 0
    If existing Is Nothing Then
        Dim btn1 As Button
        Set btn1 = wsPanel.Buttons.Add(wsPanel.Range("B2").Left, maxBottom + 16, 220, 36)
        btn1.Name = "予測構成比グラフボタン"
        btn1.OnAction = "CreateForecastRatioChart"
        btn1.Characters.Text = "予測構成比グラフを作成"
        btn1.Font.Size = 12
        btn1.Font.Bold = True
        maxBottom = maxBottom + 16 + 36
    End If

    Set existing = Nothing
    On Error Resume Next
    Set existing = wsPanel.Shapes("実績構成比グラフボタン")
    On Error GoTo 0
    If existing Is Nothing Then
        Dim btn2 As Button
        Set btn2 = wsPanel.Buttons.Add(wsPanel.Range("B2").Left, maxBottom + 16, 220, 36)
        btn2.Name = "実績構成比グラフボタン"
        btn2.OnAction = "CreateActualRatioChart"
        btn2.Characters.Text = "実績構成比グラフを作成"
        btn2.Font.Size = 12
        btn2.Font.Bold = True
    End If
End Sub
