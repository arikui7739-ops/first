Attribute VB_Name = "Module9_石狩"
Option Explicit

' ----------------------------------------------------
' 構成比グラフの作成
' 「予測構成比グラフ」:「予測データ」シート(Module8で取込済み)の投入回数_予測を号機別に集計してグラフ化する
' 「実績構成比グラフ」:ピッキング実績ファイル(S71)をダイアログで選択し、号機別ヒット数を集計してグラフ化する
' どちらも「設定」シートの「■機番別目標構成比」が入力されていれば、目標比率の系列を並べて比較できるようにする
' 既存の同名シートは削除してから作り直すため、再実行すると内容が更新される
' ----------------------------------------------------

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

    Dim maxMachNum As Long
    Dim dictTargetRatio As Object
    Call LoadSettingsForChart(maxMachNum, dictTargetRatio)

    ' 「予測データ」シートは1行目=取込情報、3行目=見出し(Module8の出力形式)。
    ' 号機列・投入回数_予測列を見出し名から探す(列の並びが変わっても追随できるようにする)
    Const HEADER_ROW As Long = 3
    Dim lastRow As Long: lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    Dim lastCol As Long: lastCol = wsData.Cells(HEADER_ROW, wsData.Columns.Count).End(xlToLeft).Column

    Dim machColIdx As Long: machColIdx = -1
    Dim cntColIdx As Long: cntColIdx = -1
    Dim hc As Long
    For hc = 1 To lastCol
        Dim hName As String: hName = Trim(CStr(wsData.Cells(HEADER_ROW, hc).Value))
        If hName = "号機" Then machColIdx = hc
        If hName = "投入回数_予測" Then cntColIdx = hc
    Next hc
    If machColIdx = -1 Or cntColIdx = -1 Then
        MsgBox "「予測データ」シートに「号機」または「投入回数_予測」の列が見つかりません。", vbExclamation
        Exit Sub
    End If

    Dim machTotal() As Double
    ReDim machTotal(1 To maxMachNum)
    Dim r As Long
    For r = HEADER_ROW + 1 To lastRow
        If IsNumeric(wsData.Cells(r, machColIdx).Value) Then
            Dim m As Long: m = CLng(wsData.Cells(r, machColIdx).Value)
            If m >= 1 And m <= maxMachNum Then
                machTotal(m) = machTotal(m) + Val(wsData.Cells(r, cntColIdx).Value)
            End If
        End If
    Next r

    Call BuildRatioChartSheet("予測構成比グラフ", "号機別構成比(予測データ)", machTotal, maxMachNum, dictTargetRatio)

    MsgBox "「予測構成比グラフ」を作成しました。", vbInformation
End Sub

Sub CreateActualRatioChart()
    Call EnsureRatioChartButtons

    Dim maxMachNum As Long
    Dim dictTargetRatio As Object
    Call LoadSettingsForChart(maxMachNum, dictTargetRatio)

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

    ' E行(機番2桁+段2桁+列2桁の9文字区切り)から号機だけを集計する(除外設定・品コードは考慮しない生の実績)
    Dim machTotal() As Double
    ReDim machTotal(1 To maxMachNum)
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
                        If mach >= 1 And mach <= maxMachNum Then
                            machTotal(mach) = machTotal(mach) + 1
                        End If
                    End If
                Next slotStart
            End If
        Loop
        Close #fileNo
    Next fIdx

    Call BuildRatioChartSheet("実績構成比グラフ", "号機別構成比(S71実績)", machTotal, maxMachNum, dictTargetRatio)

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "「実績構成比グラフ」を作成しました。(" & fd.SelectedItems.Count & "ファイル読込)", vbInformation
End Sub

' 「設定」シートから最大機番・機番別目標構成比を読み込む(Module3のプロシージャをそのまま流用する。
' 除外機番・除外ロケーション等はこのグラフでは使わないので受け皿の変数に読み捨てる)
Private Sub LoadSettingsForChart(ByRef maxMachNum As Long, ByRef dictTargetRatio As Object)
    Dim dictExcludedMach As Object: Set dictExcludedMach = CreateObject("Scripting.Dictionary")
    Dim locMach() As Long, locDanFrom() As Long, locDanTo() As Long, locColFrom() As Long, locColTo() As Long
    Dim locCount As Long
    Dim dictExcludedItemCode As Object: Set dictExcludedItemCode = CreateObject("Scripting.Dictionary")
    Dim ratioSheetName As String, maxSwapRows As Long, abSlotCount As Long
    Set dictTargetRatio = CreateObject("Scripting.Dictionary")

    Call EnsureExclusionSettingsSheet
    Call LoadExclusionSettings(dictExcludedMach, locMach, locDanFrom, locDanTo, locColFrom, locColTo, locCount, dictExcludedItemCode, ratioSheetName, maxSwapRows, maxMachNum, abSlotCount, dictTargetRatio)
End Sub

' 号機別の合計値(machTotal)から、構成比データ表とグラフを持つシートを作成する(既存の同名シートは削除して作り直す)
Private Sub BuildRatioChartSheet(ByVal sheetName As String, ByVal chartTitle As String, machTotal() As Double, ByVal maxMachNum As Long, dictTargetRatio As Object)
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

    Dim grandTotal As Double: grandTotal = 0
    Dim m As Long
    For m = 1 To maxMachNum
        grandTotal = grandTotal + machTotal(m)
    Next m

    wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(1, 5)).Merge
    wsOut.Cells(1, 1).Value = "【" & chartTitle & "】作成日時: " & Format(Now, "yyyy/mm/dd hh:mm")
    wsOut.Cells(1, 1).Font.Bold = True: wsOut.Cells(1, 1).Font.Size = 12
    wsOut.Cells(1, 1).HorizontalAlignment = xlLeft

    Dim hasTarget As Boolean: hasTarget = (dictTargetRatio.Count > 0)

    wsOut.Range(wsOut.Cells(3, 1), wsOut.Cells(3, 5)).Value = Array("号機", "合計", "比率", "目標比率", "乖離")
    wsOut.Range(wsOut.Cells(3, 1), wsOut.Cells(3, 5)).Interior.Color = RGB(220, 230, 255)
    wsOut.Range(wsOut.Cells(3, 1), wsOut.Cells(3, 5)).Font.Bold = True

    Dim r As Long: r = 3
    For m = 1 To maxMachNum
        r = r + 1
        wsOut.Cells(r, 1).Value = "AB" & Format(m, "00")
        wsOut.Cells(r, 2).Value = machTotal(m)
        Dim ratioVal As Double: ratioVal = 0
        If grandTotal > 0 Then ratioVal = machTotal(m) / grandTotal
        wsOut.Cells(r, 3).Value = ratioVal
        wsOut.Cells(r, 3).NumberFormat = "0.0%"
        If hasTarget And dictTargetRatio.Exists(CStr(m)) Then
            Dim tRatio As Double: tRatio = dictTargetRatio(CStr(m))
            wsOut.Cells(r, 4).Value = tRatio
            wsOut.Cells(r, 4).NumberFormat = "0.0%"
            wsOut.Cells(r, 5).Value = Abs(ratioVal - tRatio)
            wsOut.Cells(r, 5).NumberFormat = "0.0%"
        End If
    Next m
    Dim lastDataRow As Long: lastDataRow = r

    wsOut.Range(wsOut.Cells(3, 1), wsOut.Cells(lastDataRow, 5)).Columns.AutoFit

    ' グラフ(号機ごとの比率。目標構成比が入力されていれば目標比率も並べて比較できるようにする)
    Dim valueRange As Range
    If hasTarget Then
        Set valueRange = wsOut.Range(wsOut.Cells(3, 3), wsOut.Cells(lastDataRow, 4))
    Else
        Set valueRange = wsOut.Range(wsOut.Cells(3, 3), wsOut.Cells(lastDataRow, 3))
    End If
    Dim catRange As Range: Set catRange = wsOut.Range(wsOut.Cells(3, 1), wsOut.Cells(lastDataRow, 1))

    Dim chtObj As ChartObject
    Set chtObj = wsOut.ChartObjects.Add(wsOut.Cells(3, 7).Left, wsOut.Cells(3, 7).Top, 760, 380)
    With chtObj.Chart
        .SetSourceData Source:=Union(catRange, valueRange)
        .PlotBy = xlColumns
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = chartTitle
        .Axes(xlValue).TickLabels.NumberFormat = "0%"
        .Axes(xlCategory).TickLabels.Font.Size = 7
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
