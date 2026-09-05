Attribute VB_Name = "Module8_金沢"
Option Explicit

' ----------------------------------------------------
' 予測データの取込(CSVファイルをダイアログで選択し「予測データ」シートへ読み込む)
' 既存の「予測データ」シートは削除してから作り直すため、再取込すると内容が上書き更新される
' ----------------------------------------------------

Sub ImportPredictionData()
    ' 「操作パネル」シートにボタンが無ければ追加する(初回はここでボタンが作られ、以降はワンクリックで実行できる)
    Call EnsurePredictionImportButton

    ' 1. ファイル選択(単一選択・全ファイル形式)
    Dim fd As Office.FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "予測データファイル(CSV)を選択"
        .Filters.Clear
        .Filters.Add "すべてのファイル", "*.*"
        .AllowMultiSelect = False
        If .Show = False Then Exit Sub
    End With

    Dim filePath As String: filePath = fd.SelectedItems(1)
    Dim fileDate As Date: fileDate = FileDateTime(filePath) ' 予測データが出力された日時の目安として使う

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    ' 2. ファイルの読込(1行目=見出し、2行目以降=データ。カンマ区切り・引用符なしの単純なCSV)
    Dim fileNo As Integer: fileNo = FreeFile
    Dim textLine As String
    Dim headerArr() As String
    Dim dataRows As Collection: Set dataRows = New Collection
    Dim lineIdx As Long: lineIdx = 0

    Open filePath For Input As #fileNo
    Do While Not EOF(fileNo)
        Line Input #fileNo, textLine
        If Trim(textLine) <> "" Then
            lineIdx = lineIdx + 1
            If lineIdx = 1 Then
                headerArr = Split(textLine, ",")
            Else
                dataRows.Add Split(textLine, ",")
            End If
        End If
    Loop
    Close #fileNo

    If lineIdx <= 1 Then
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        MsgBox "データ行が見つかりませんでした。ファイルの内容を確認してください。", vbExclamation
        Exit Sub
    End If

    Dim colCount As Long: colCount = UBound(headerArr) - LBound(headerArr) + 1
    Dim rowCount As Long: rowCount = dataRows.Count

    ' 3. 先頭ゼロを保つべき列(号機・段・列・品名コードなど)を見出し名から特定する
    '    (列の並びが変わっても追随できるよう、列番号を決め打ちしない)
    Dim textColNames As Variant: textColNames = Array("号機", "段", "列", "品名コード")
    Dim textColIdx() As Long: ReDim textColIdx(LBound(textColNames) To UBound(textColNames))
    Dim tc As Long
    For tc = LBound(textColNames) To UBound(textColNames)
        textColIdx(tc) = -1
    Next tc
    Dim hc As Long
    For hc = LBound(headerArr) To UBound(headerArr)
        For tc = LBound(textColNames) To UBound(textColNames)
            If Trim(headerArr(hc)) = textColNames(tc) Then textColIdx(tc) = hc - LBound(headerArr) + 1 ' 1始まりの列番号
        Next tc
    Next hc

    ' 4. データを配列にまとめてから一括で書き込む(1行ずつCells書き込みするより高速)
    Dim dataArr() As Variant
    ReDim dataArr(1 To rowCount, 1 To colCount)
    Dim r As Long: r = 0
    Dim rowVariant As Variant
    For Each rowVariant In dataRows
        r = r + 1
        Dim fields() As String: fields = rowVariant
        Dim c As Long
        For c = 1 To colCount
            If c - 1 >= LBound(fields) And c - 1 <= UBound(fields) Then
                dataArr(r, c) = fields(c - 1)
            Else
                dataArr(r, c) = ""
            End If
        Next c
    Next rowVariant

    ' 5. 「予測データ」シートを作り直す(既存があれば削除。これにより再取込のたびに内容が更新される)
    On Error Resume Next
    Sheets("予測データ").Delete
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
    wsOut.Name = "予測データ"

    ' 先頭ゼロを保つ列は、書き込む前に文字列書式にしておく(数値変換されて先頭ゼロが落ちるのを防ぐ)
    For tc = LBound(textColIdx) To UBound(textColIdx)
        If textColIdx(tc) > 0 Then wsOut.Columns(textColIdx(tc)).NumberFormat = "@"
    Next tc

    ' ファイル名・ファイル更新日時(=予測データの基準日時の目安)・取込日時を1行目に表示する
    wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(1, colCount)).Merge
    wsOut.Cells(1, 1).Value = "【予測データ取込】ファイル: " & Dir(filePath) & _
        "　／　ファイル更新日時: " & Format(fileDate, "yyyy/mm/dd hh:mm") & _
        "　／　取込日時: " & Format(Now, "yyyy/mm/dd hh:mm")
    wsOut.Cells(1, 1).Font.Bold = True: wsOut.Cells(1, 1).Font.Size = 12
    wsOut.Cells(1, 1).HorizontalAlignment = xlLeft

    Const HEADER_ROW As Long = 3
    Dim hi As Long
    For hi = 1 To colCount
        wsOut.Cells(HEADER_ROW, hi).Value = headerArr(hi - 1)
    Next hi
    wsOut.Range(wsOut.Cells(HEADER_ROW, 1), wsOut.Cells(HEADER_ROW, colCount)).Interior.Color = RGB(220, 230, 255)
    wsOut.Range(wsOut.Cells(HEADER_ROW, 1), wsOut.Cells(HEADER_ROW, colCount)).Font.Bold = True

    wsOut.Range(wsOut.Cells(HEADER_ROW + 1, 1), wsOut.Cells(HEADER_ROW + rowCount, colCount)).Value = dataArr

    wsOut.Range(wsOut.Cells(HEADER_ROW, 1), wsOut.Cells(HEADER_ROW, colCount)).AutoFilter
    wsOut.Range(wsOut.Cells(HEADER_ROW, 1), wsOut.Cells(HEADER_ROW, colCount)).Columns.AutoFit
    wsOut.Rows(1).RowHeight = 20

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "「予測データ」シートを更新しました。(" & rowCount & "行取込)" & vbCrLf & _
        "ファイル更新日時: " & Format(fileDate, "yyyy/mm/dd hh:mm"), vbInformation
End Sub

' 「操作パネル」シートに予測データ取込ボタンが無ければ追加する
' (既存のボタン・図形と重ならないよう、一番下にあるものの少し下に配置する)
Sub EnsurePredictionImportButton()
    Dim wsPanel As Worksheet
    On Error Resume Next
    Set wsPanel = ThisWorkbook.Sheets("操作パネル")
    On Error GoTo 0
    If wsPanel Is Nothing Then Exit Sub

    Dim existing As Shape
    On Error Resume Next
    Set existing = wsPanel.Shapes("予測データ取込ボタン")
    On Error GoTo 0
    If existing Is Nothing Then
        Dim btn As Button
        Set btn = wsPanel.Buttons.Add(wsPanel.Range("B20").Left, wsPanel.Range("B20").Top, 220, 36)
        btn.Name = "予測データ取込ボタン"
        btn.OnAction = "ImportPredictionData"
        btn.Characters.Text = "予測データを取り込む"
        btn.Font.Size = 12
        btn.Font.Bold = True
    End If

    ' ボタンが下に伸び続けないよう、2列に並び替える(Module3の共通処理)
    Call LayoutPanelButtons
End Sub
