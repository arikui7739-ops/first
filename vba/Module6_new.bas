Attribute VB_Name = "Module6"
Sub CreateControlPanel()
    Dim wsPanel As Worksheet

    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("操作パネル").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set wsPanel = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
    wsPanel.name = "操作パネル"

    wsPanel.Columns("A").ColumnWidth = 2
    wsPanel.Columns("B").ColumnWidth = 55
    wsPanel.Columns("C").ColumnWidth = 3
    wsPanel.Columns("D").ColumnWidth = 3

    wsPanel.Cells(1, 2).Value = "ロケーション管理マクロ 操作パネル"
    wsPanel.Cells(1, 2).Font.Bold = True
    wsPanel.Cells(1, 2).Font.Size = 16
    wsPanel.Rows(1).RowHeight = 28

    wsPanel.Cells(2, 2).Value = "各マクロの内容を確認のうえ、右側のボタンから実行してください。"
    wsPanel.Cells(2, 2).Font.Italic = True

    Dim titles(1 To 4) As String
    Dim descs(1 To 4) As String
    Dim actions(1 To 4) As String
    Dim colors(1 To 4) As Long

    titles(1) = "① 同時ピッキング交換指示書（Cエリア・実績ベース）"
    descs(1) = "ピッキング実績ファイル(H/E形式、複数選択可)を読み込み、Cエリア(51～68号機)で同じ受注内に同時出現しやすい商品ペアを検出し、近くの低頻度品と入れ替える指示を作成します。実行するとファイル選択ダイアログが開きます。"
    actions(1) = "Module1.SwapLocationsByCorrelationFast_Fix"
    colors(1) = RGB(255, 235, 235)

    titles(2) = "② AB編成流れ最適化（対面ピッキング・実績ベース）"
    descs(2) = "ピッキング実績ファイル(複数選択可)を読み込み、ABエリア(5～46号機)で6オーダーを1編成として、対面同時ピッキングや同一編成内の複数回ピッキングを減らす入替候補を相関上位10件抽出します。奇数/偶数号機のバランスも考慮します。"
    actions(2) = "Module3.OptimizeABFormationFlow"
    colors(2) = RGB(220, 230, 255)

    titles(3) = "③ ゾーンバランス最適（ABエリア・予測ベース）"
    descs(3) = "CFシートの投入回数予測と号機回数比シートの目標比率を基に、ABエリアで目標比率を超えている号機から売れ筋を分散させ、Cエリアとの間で昇格・降格の移動指示を作成します。CFシートをアクティブにして実行してください。"
    actions(3) = "Module4.CreatePredictionLocationInstruction"
    colors(3) = RGB(255, 245, 220)

    titles(4) = "④ Cバラ動線最適化（Cエリア・予測ベース）"
    descs(4) = "Cエリア(51～68号機)の号機・段・列ごとのペナルティスコアをもとに、投入回数予測の高い商品を好立地スロットへ再配置する最適ロケーション変更指示書を作成します。CFシート等、号機・段・列・品名コード・投入回数_予測の列があるシートをアクティブにして実行してください。"
    actions(4) = "Module5.GenerateLocationInstructions"
    colors(4) = RGB(220, 245, 230)

    Dim i As Integer
    Dim curRow As Long: curRow = 4

    For i = 1 To 4
        Dim titleRow As Long: titleRow = curRow
        Dim descRow As Long: descRow = curRow + 1

        wsPanel.Cells(titleRow, 2).Value = titles(i)
        wsPanel.Cells(titleRow, 2).Font.Bold = True
        wsPanel.Cells(titleRow, 2).Font.Size = 12
        wsPanel.Rows(titleRow).RowHeight = 20

        wsPanel.Cells(descRow, 2).Value = descs(i)
        wsPanel.Cells(descRow, 2).WrapText = True
        wsPanel.Cells(descRow, 2).VerticalAlignment = xlTop
        wsPanel.Rows(descRow).RowHeight = 60

        wsPanel.Range(wsPanel.Cells(titleRow, 2), wsPanel.Cells(descRow, 2)).Interior.Color = colors(i)

        ' ボタンをブロックの右側（E列付近）に配置し、対応するマクロを割り当てる
        Dim btnTop As Double: btnTop = wsPanel.Cells(titleRow, 5).Top
        Dim btnLeft As Double: btnLeft = wsPanel.Cells(titleRow, 5).Left
        Dim btnHeight As Double: btnHeight = wsPanel.Rows(titleRow).RowHeight + wsPanel.Rows(descRow).RowHeight

        Dim btn As Button
        Set btn = wsPanel.Buttons.Add(btnLeft, btnTop, 100, btnHeight)
        btn.OnAction = actions(i)
        btn.Characters.Text = "実行する"
        btn.Font.Size = 11
        btn.Font.Bold = True

        curRow = descRow + 2 ' 次のブロックとの間隔
    Next i

    wsPanel.Columns("A:E").AutoFit
    wsPanel.Columns("B").ColumnWidth = 55
    wsPanel.Activate
    wsPanel.Range("A1").Select
    MsgBox "「操作パネル」シートを作成しました。各ボタンから対応するマクロを実行できます。" & vbCrLf & _
        "ボタンの位置・大きさは後から自由にドラッグして調整できます（機能には影響しません）。", vbInformation
End Sub
