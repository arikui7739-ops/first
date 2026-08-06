Attribute VB_Name = "Module9_石狩"
Option Explicit

' ----------------------------------------------------
' 構成比グラフの作成
' 「予測構成比グラフ」:「予測データ」シート(Module8で取込済み)の投入回数_予測をゾーン別に集計してグラフ化する
' 「実績構成比グラフ」:ピッキング実績ファイル(S71)をダイアログで選択し、ゾーン別ヒット数を集計してグラフ化する
' レイアウトは「設定」シートの「■機番別目標構成比」と同じ考え方で、AB01～AB46は機番ペア(ゾーン)ごとに
' 奇数機番を上向き・偶数機番を下向きに表示し、Cバラ(C01・C02)・拡張(X)は上向きの単独項目として追加する。
' 「設定」シートに目標構成比が入力されていれば、実績/予測の比率と並べて目標比率も折れ線で比較できるようにする。
' 既存の同名シートは削除してから作り直すため、再実行すると内容が更新される
' ----------------------------------------------------

Const MAX_AB_MACH As Long = 46 ' 石狩のAB機番範囲(1～46)。この範囲は奇数/偶数ペアのゾーン表示にする

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

    Call BuildRatioChartSheet("予測構成比グラフ", "号機別構成比(予測データ)", dictActualByLabel, dictTargetByLabel)

    MsgBox "「予測構成比グラフ」を作成しました。", vbInformation
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

    ' E行(機番2桁+段2桁+列2桁の9文字区切り)からゾーンラベルを判定して集計する(除外設定・品コードは考慮しない生の実績)
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
                    End If
                Next slotStart
            End If
        Loop
        Close #fileNo
    Next fIdx

    Call BuildRatioChartSheet("実績構成比グラフ", "号機別構成比(S71実績)", dictActualByLabel, dictTargetByLabel)

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    Dim noteMsg As String
    If Not hasZoneMap Then noteMsg = vbCrLf & "※「予測データ」シートが無いため、Cバラ(C01/C02)・拡張(X)は集計されていません(AB01～AB46のみ)。"
    MsgBox "「実績構成比グラフ」を作成しました。(" & fd.SelectedItems.Count & "ファイル読込)" & noteMsg, vbInformation
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
' Cバラ(C01・C02)・拡張(X)などAB以外のラベルは、奇数側の列に単独の上向き項目として追加する
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

        wsOut.Cells(r, 2).Value = IIf(grandTotal > 0, oddVal / grandTotal, 0)
        wsOut.Cells(r, 2).NumberFormat = "0.0%"
        wsOut.Cells(r, 3).Value = -IIf(grandTotal > 0, evenVal / grandTotal, 0)
        wsOut.Cells(r, 3).NumberFormat = "0.0%;0.0%"

        If hasTarget Then
            If dictTargetByLabel.Exists(oddLabel) Then
                wsOut.Cells(r, 4).Value = dictTargetByLabel(oddLabel)
                wsOut.Cells(r, 4).NumberFormat = "0.0%"
            End If
            If dictTargetByLabel.Exists(evenLabel) Then
                wsOut.Cells(r, 5).Value = -dictTargetByLabel(evenLabel)
                wsOut.Cells(r, 5).NumberFormat = "0.0%;0.0%"
            End If
        End If
    Next z

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

    ' グラフ(奇数機番を上向き・偶数機番を下向きの面グラフで表示。Cバラ・拡張は奇数側に単独の項目として並ぶ。
    ' 目標構成比が入力されていれば、目標比率を折れ線で重ねて比較できるようにする)
    Dim srcCols As Long: srcCols = IIf(hasTarget, 5, 3)
    Dim srcRange As Range: Set srcRange = wsOut.Range(wsOut.Cells(3, 1), wsOut.Cells(lastDataRow, srcCols))

    Dim chtObj As ChartObject
    Set chtObj = wsOut.ChartObjects.Add(wsOut.Cells(3, 7).Left, wsOut.Cells(3, 7).Top, 900, 380)
    With chtObj.Chart
        .SetSourceData Source:=srcRange
        .PlotBy = xlColumns
        .ChartType = xlArea
        .HasTitle = True
        .ChartTitle.Text = chartTitle
        .Axes(xlValue).TickLabels.NumberFormat = "0%;0%"
        .Axes(xlCategory).TickLabels.Font.Size = 7
        .HasLegend = True
        If hasTarget Then
            .SeriesCollection(3).ChartType = xlLineMarkers
            .SeriesCollection(4).ChartType = xlLineMarkers
        End If
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
