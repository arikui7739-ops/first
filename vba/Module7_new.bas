Attribute VB_Name = "Module7"
' ==========================================================================
' KPI記録シートの管理（実績データのみを対象。予測ベースのModule4・Module5は対象外）
' 複合スコアの重み（相対比。実行済みの指標だけで比率を再正規化するため合計100でなくてもよい）：
'   AB号機使用比率:40、AB占有率(理論値実績差):15、奇数偶数比率:25、対面同時ヒット:20、Cバラ無駄歩行:15
' 行のキーは「実績日」（ピッキング実績ファイルの日付）。同じ実績日なら行を更新し、新しい行は増やしません。
' 評価の目安: 70点以上=◎良好(緑)、40～69点=○普通(黄)、40点未満=×要対策(赤)
' ==========================================================================
Private Const W_ABRATIO As Double = 0.4    ' AB号機使用比率スコアの重み
Private Const W_ABOCCUPANCY As Double = 0.15 ' AB占有率スコア(理論値実績差)の重み
Private Const W_ODDEVEN As Double = 0.25   ' 奇数偶数比率スコアの重み
Private Const W_CROSSFACE As Double = 0.2  ' 対面同時ヒットスコアの重み
Private Const W_OTHER As Double = 0.15     ' Cバラ無駄歩行スコアの重み

Private Function GetKPISheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("KPI記録")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.name = "KPI記録"
        ' 見出しは2行に折り返して横幅をコンパクトにする（正式名称はO列以降の説明を参照）
        ws.Range("A1:M1").Value = Array( _
            "実績日", "曜日", _
            "AB号機" & vbLf & "使用比率", _
            "AB占有率" & vbLf & "スコア", _
            "理論比率" & vbLf & "(占有率)", _
            "実績比率" & vbLf & "(占有率)", _
            "奇偶バランス" & vbLf & "(交換前)", _
            "奇偶バランス" & vbLf & "(交換後)", _
            "対面同時" & vbLf & "ヒット", _
            "Cバラ" & vbLf & "無駄歩行", _
            "総合" & vbLf & "スコア", _
            "評価", _
            "更新時刻")
        ws.Range("A1:M1").Font.Bold = True
        ws.Range("A1:M1").Interior.Color = RGB(210, 225, 245)
        ws.Range("A1:M1").WrapText = True
        ws.Range("A1:M1").HorizontalAlignment = xlCenter
        ws.Range("A1:M1").VerticalAlignment = xlCenter

        ws.Columns("A:A").ColumnWidth = 11
        ws.Columns("B:B").ColumnWidth = 6
        ws.Columns("C:D").ColumnWidth = 9
        ws.Columns("E:F").ColumnWidth = 10
        ws.Columns("G:K").ColumnWidth = 9
        ws.Columns("L:L").ColumnWidth = 8
        ws.Columns("M:M").ColumnWidth = 16

        ws.Columns("A:A").NumberFormat = "yyyy/mm/dd"
        ws.Columns("C:D").NumberFormat = "0" ' スコアは整数表示
        ws.Columns("E:F").NumberFormat = "0.00%" ' 理論比率・実績比率は％小数点2位まで
        ws.Columns("G:K").NumberFormat = "0" ' スコアは整数表示
        ws.Columns("M:M").NumberFormat = "yyyy/mm/dd hh:mm"
        ws.Rows(1).AutoFit ' 見出し行の高さだけ折り返しに合わせて調整（列幅は固定のまま）
    End If

    ' O列以降にKPIの説明を転記（データ行が増えても重ならないよう右側に配置）
    ' シートが既に存在していてもO1が未記入なら追記する（既存データは消さない）
    ' ★行の高さは同じ行のA:M列（データ行）にも影響するため、説明は1行1文の通常の高さで書く（大きくしない）
    If ws.Cells(1, 15).Value = "" Then
        ws.Columns("N").ColumnWidth = 3
        ws.Columns("O").ColumnWidth = 90

        Dim expLines() As String
        Dim expPart1 As String, expPart2 As String, expPart3 As String
        expPart1 = _
            "【KPI記録の説明】|" & _
            "|" & _
            "■ 各列の意味|" & _
            "実績日：Module1・Module3で読み込んだピッキング実績ファイルのB行（先頭Bと8桁日付）から読み取った日付。この日付をキーに、同じ実績日なら行を更新します。|" & _
            "曜日：実績日の曜日。|" & _
            "AB号機使用比率スコア：号機回数比シートの目標比率(理論値)と、実績ファイルから集計した号機別実績比率の近さを評価。|" & _
            "AB占有率スコア：AB全体としての理論的な占有率と実績占有率の差を評価（下記参照）。|" & _
            "AB占有率_理論比率／実績比率：上記スコアの元になった実際の比率（％）。差の大きさを直接確認したいときに使う。|" & _
            "奇数偶数比率スコア(交換前)：Module3のスワップ提案を行う前、実績データそのものの奇数・偶数号機バランス。日々の実態のブレが出る指標。|" & _
            "奇数偶数比率スコア(交換後)：Module3のスワップ提案を反映した場合の見込みバランス。総合スコアはこちらを使用（提案アルゴリズムの性質上、高得点に収束しやすい）。|" & _
            "対面同時ヒットスコア：対面競合の少なさ。理論上どこまで減らせるかを基準に評価（下記参照）。|" & _
            "Cバラ無駄歩行：Module1実行時に記録する参考指標（暫定正規化）。|" & _
            "総合スコア：下記の計算式による複合スコア。評価：総合スコアに応じた◎/○/×判定。|" & _
            "※予測データを使うModule4・Module5は、実績データのみを評価するというKPIの方針上、この記録の対象外です。|"
        expPart2 = _
            "|" & _
            "■ 総合スコアの計算式|" & _
            "総合スコア = AB号機使用比率×40 + AB占有率×15 + 奇数偶数比率×25 + 対面同時ヒット×20 + Cバラ無駄歩行×15 の加重平均|" & _
            "（その日にまだ全部のマクロを実行していない場合は、揃っている指標だけで比率を再正規化して計算します）|" & _
            "|" & _
            "■ 各スコアの算出方法|" & _
            "AB号機使用比率スコア = 100×(1－Σ｜実績比率－目標比率｜÷0.6)。0.6以上は0点、0で100点、比例配分|" & _
            "　目標比率=号機回数比シートのE列。実績比率=実績ファイルの号機別ヒット数を号機5～46号機の合計で正規化した値。|" & _
            "AB占有率スコア = 理論値・実績値の差（後述）が1%以内なら100点、6%以上なら0点、その間は比例配分|" & _
            "　理論値：実績ファイル全体を回数の多い順に並べ、上位900アイテムの回数合計 ÷ 全アイテムの回数合計|"
        expPart3 = _
            "　実績値：AB間口(1～46号機、6～14列は除く)の実績回数合計 ÷ 全アイテムの回数合計|" & _
            "奇数偶数比率スコア(交換前/交換後共通の式) = 100×(1－｜奇数合計－偶数合計｜÷(奇数合計+偶数合計))|" & _
            "　交換前：Module3がスワップ提案を作る前の実績ヒット数で計算。交換後：提案どおりに交換した場合の見込みヒット数で計算。|" & _
            "対面同時ヒットスコア = 100×理論最小対面ヒット数÷実績対面ヒット数（100が上限）|" & _
            "　理論最小対面ヒット数：ゾーンごとに今の奇数側/偶数側の個数を保ったまま最適配置し直した場合に達成できる対面ヒット数の最小値(局所探索で算出)|" & _
            "Cバラ無駄歩行 = 100－min(100, 平均無駄歩行スコア÷20×100)　【暫定式】|" & _
            "|" & _
            "■ 評価の目安|" & _
            "70点以上：◎良好（緑）　　40～69点：○普通（黄）　　40点未満：×要対策（赤）|" & _
            "※Cバラ無駄歩行の暫定式（÷20の部分）は実データがまだ無い状態での仮の基準です。数週間分の実績が溜まったら調整してください。"
        expLines = Split(expPart1 & expPart2 & expPart3, "|")

        Dim li As Long
        For li = 0 To UBound(expLines)
            Dim lineText As String: lineText = expLines(li)
            ws.Cells(li + 1, 15).Value = lineText
            If Left(lineText, 1) = "【" Then
                ws.Cells(li + 1, 15).Font.Bold = True: ws.Cells(li + 1, 15).Font.Size = 14
            ElseIf Left(lineText, 1) = "■" Then
                ws.Cells(li + 1, 15).Font.Bold = True
            End If
        Next li
    End If
    Set GetKPISheet = ws
End Function

' 「実績日」をキーに、既存行があればその行番号、無ければ新規行を作って返す
Private Function GetRowForDate(ByVal actualDate As Date) As Long
    Dim ws As Worksheet: Set ws = GetKPISheet()
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    Dim r As Long

    For r = 2 To lastRow
        If ws.Cells(r, 1).Value <> "" Then
            If CDate(ws.Cells(r, 1).Value) = actualDate Then
                GetRowForDate = r
                Exit Function
            End If
        End If
    Next r

    GetRowForDate = lastRow + 1
    ws.Cells(GetRowForDate, 1).Value = actualDate
    ws.Cells(GetRowForDate, 2).Value = Format(actualDate, "aaa")
End Function

' スコアに応じてセルを色分けする（70以上=緑／40～69=黄／40未満=赤）
Private Sub ColorByScore(ByVal cell As Range, ByVal score As Double)
    If score >= 70 Then
        cell.Interior.Color = RGB(198, 239, 206)
        cell.Font.Color = RGB(0, 97, 0)
    ElseIf score >= 40 Then
        cell.Interior.Color = RGB(255, 235, 156)
        cell.Font.Color = RGB(156, 101, 0)
    Else
        cell.Interior.Color = RGB(255, 199, 206)
        cell.Font.Color = RGB(156, 0, 6)
    End If
End Sub

Private Function GradeText(ByVal score As Double) As String
    If score >= 70 Then
        GradeText = "◎良好"
    ElseIf score >= 40 Then
        GradeText = "○普通"
    Else
        GradeText = "×要対策"
    End If
End Function

Private Sub RecalcRow(ByVal r As Long)
    Dim ws As Worksheet: Set ws = GetKPISheet()

    ' 各スコア列(C,D,G,H,I,J)を色分け（値がある場合のみ）。E,F(理論比率/実績比率)はスコアではないため対象外
    Dim colIdx As Variant
    Dim scoreCols As Variant: scoreCols = Array(3, 4, 7, 8, 9, 10)
    For Each colIdx In scoreCols
        If ws.Cells(r, colIdx).Value <> "" Then
            ws.Cells(r, colIdx).Value = Application.WorksheetFunction.Round(ws.Cells(r, colIdx).Value, 0)
            ColorByScore ws.Cells(r, colIdx), ws.Cells(r, colIdx).Value
        End If
    Next colIdx

    ' 総合スコア(その日そろっている指標だけで比率を再正規化)
    ' ※奇数偶数比率は「交換後」(H列)を使用。「交換前」(G列)はスワップ提案前の実態を見るための参考列で、総合スコアには含めない
    Dim wsum As Double, wtotal As Double
    wsum = 0: wtotal = 0
    If ws.Cells(r, 3).Value <> "" Then wsum = wsum + ws.Cells(r, 3).Value * W_ABRATIO: wtotal = wtotal + W_ABRATIO
    If ws.Cells(r, 4).Value <> "" Then wsum = wsum + ws.Cells(r, 4).Value * W_ABOCCUPANCY: wtotal = wtotal + W_ABOCCUPANCY
    If ws.Cells(r, 8).Value <> "" Then wsum = wsum + ws.Cells(r, 8).Value * W_ODDEVEN: wtotal = wtotal + W_ODDEVEN
    If ws.Cells(r, 9).Value <> "" Then wsum = wsum + ws.Cells(r, 9).Value * W_CROSSFACE: wtotal = wtotal + W_CROSSFACE
    If ws.Cells(r, 10).Value <> "" Then wsum = wsum + ws.Cells(r, 10).Value * W_OTHER: wtotal = wtotal + W_OTHER

    If wtotal > 0 Then
        Dim finalScore As Double: finalScore = Application.WorksheetFunction.Round(wsum / wtotal, 0)
        ws.Cells(r, 11).Value = finalScore
        ws.Cells(r, 11).Font.Bold = True
        ColorByScore ws.Cells(r, 11), finalScore
        ws.Cells(r, 12).Value = GradeText(finalScore)
        ws.Cells(r, 12).Font.Bold = True
        ColorByScore ws.Cells(r, 12), finalScore
    Else
        ws.Cells(r, 11).Value = ""
        ws.Cells(r, 12).Value = ""
        ws.Cells(r, 11).Interior.ColorIndex = xlNone
        ws.Cells(r, 12).Interior.ColorIndex = xlNone
    End If

    ws.Cells(r, 13).Value = Now
    ' ※列幅は見出し作成時に固定済み（ここでAutoFitすると見出しの折り返しが崩れて再び横長になるため呼ばない）
End Sub

' --- 各モジュールから呼び出す記録用プロシージャ ---

' Module3から: AB号機使用比率・AB占有率(スコア＋理論比率／実績比率)・奇数偶数比率(交換前/交換後)・対面同時ヒットスコア・実績日
' ※対面同時ヒットスコアはModule3側で「理論最小対面ヒット数÷実績対面ヒット数」として計算済みの値を受け取るだけ
Public Sub LogFormationScore(ByVal oddTotalStart As Double, ByVal evenTotalStart As Double, ByVal oddTotal As Double, ByVal evenTotal As Double, ByVal crossFaceScore As Variant, ByVal abRatioScore As Variant, ByVal abOccupancyScore As Variant, ByVal abTheoreticalRatio As Variant, ByVal abActualRatio As Variant, ByVal actualDate As Date)
    Dim r As Long: r = GetRowForDate(actualDate)
    Dim ws As Worksheet: Set ws = GetKPISheet()

    Dim preBalanceScore As Double
    If (oddTotalStart + evenTotalStart) > 0 Then
        preBalanceScore = 100 * (1 - Abs(oddTotalStart - evenTotalStart) / (oddTotalStart + evenTotalStart))
    Else
        preBalanceScore = 100
    End If

    Dim balanceScore As Double
    If (oddTotal + evenTotal) > 0 Then
        balanceScore = 100 * (1 - Abs(oddTotal - evenTotal) / (oddTotal + evenTotal))
    Else
        balanceScore = 100
    End If

    If IsNumeric(abRatioScore) Then ws.Cells(r, 3).Value = abRatioScore
    If IsNumeric(abOccupancyScore) Then ws.Cells(r, 4).Value = abOccupancyScore
    If IsNumeric(abTheoreticalRatio) Then ws.Cells(r, 5).Value = abTheoreticalRatio
    If IsNumeric(abActualRatio) Then ws.Cells(r, 6).Value = abActualRatio
    ws.Cells(r, 7).Value = preBalanceScore
    ws.Cells(r, 8).Value = balanceScore
    If IsNumeric(crossFaceScore) Then ws.Cells(r, 9).Value = crossFaceScore
    RecalcRow r
End Sub

' Module1から: 交換提案の平均無駄歩行スコアを参考指標として記録（暫定正規化。実データが溜まったら調整）
Public Sub LogModule1Score(ByVal avgWasteScore As Double, ByVal actualDate As Date)
    Dim r As Long: r = GetRowForDate(actualDate)
    Dim ws As Worksheet: Set ws = GetKPISheet()
    Dim score As Double
    score = 100 - Application.WorksheetFunction.Min(100, (avgWasteScore / 20) * 100)
    ws.Cells(r, 10).Value = score
    RecalcRow r
End Sub
