Attribute VB_Name = "Module4"
'============================================================
' 計算（高速化版）
' マクロ記録日 : 2008/5/12  ユーザー名 : Kao
'
' 【遅くなっていた理由】
' ・Application.Calculation が既定の「自動」のままのため、
'   5回のRefreshTableのたびにブック全体（隠しシート「貼り付けｼｰﾄ2」
'   だけで約34,600個）の数式が毎回再計算されていた。
' ・Sheets(...).Select を5回繰り返すたびに画面が再描画されていた。
'
' 【対策】
' ・自動計算/画面更新を処理中だけOFFにし、5回分の再計算・再描画を
'   最後の1回にまとめた。
' ・ActiveSheet経由ではなく Sheets(...).PivotTables(...) を直接参照し、
'   不要なシート選択(.Select)を削除した。
' 出力結果・ロジックは元のまま。
'
' 【石狩版追加】
' 「就業時間 （石狩LC）」「日別 （石狩LC・バラ）」シートを作成後に
' 自動で更新対象へ含めるよう追加。シート未作成の間はSheetExists()で
' スキップするのでエラーにならない。
'============================================================
Sub 計算()
'マクロ記録日 : 2008/5/12  ユーザー名 : Kao

    Dim curCalc As XlCalculation

    On Error GoTo ErrHandler

    Application.ScreenUpdating = False
    curCalc = Application.Calculation
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Sheets("就業時間 （岩槻LC）").PivotTables("ﾋﾟﾎﾞｯﾄﾃｰﾌﾞﾙ1").RefreshTable
    Sheets("就業時間 （社員)").PivotTables("ﾋﾟﾎﾞｯﾄﾃｰﾌﾞﾙ1").RefreshTable
    Sheets("有給").PivotTables("ﾋﾟﾎﾞｯﾄﾃｰﾌﾞﾙ1").RefreshTable
    Sheets("日別").PivotTables("ﾋﾟﾎﾞｯﾄﾃｰﾌﾞﾙ1").RefreshTable
    Sheets("日別 (個人別)").PivotTables("ﾋﾟﾎﾞｯﾄﾃｰﾌﾞﾙ1").RefreshTable

    If SheetExists("就業時間 （石狩LC）") Then
        Sheets("就業時間 （石狩LC）").PivotTables("ﾋﾟﾎﾞｯﾄﾃｰﾌﾞﾙ1").RefreshTable
    End If

    If SheetExists("日別 （石狩LC・バラ）") Then
        Sheets("日別 （石狩LC・バラ）").PivotTables("ﾋﾟﾎﾞｯﾄﾃｰﾌﾞﾙ1").RefreshTable
    End If

ErrHandler:
    Application.Calculation = curCalc
    Application.CalculateFull '複数回分をまとめて最後に1回だけ再計算
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    Sheets("貼り付けシート").Select

    If Err.Number <> 0 Then
        MsgBox "エラーが発生しました：" & Err.Description, vbExclamation
    End If

End Sub

Private Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

'============================================================
' 月次ログ更新（月別集計の元データ蓄積）
'
' 「貼り付けシート」は毎月データを入れ替える運用のため、過去分を
' 月別に比較したい場合は別途履歴を持つ必要がある。このマクロは
' 「貼り付けシート」の現在の内容を「月次ログ」シートへ追記する。
'
' 日付（7列目、YYYYMMDD形式の数値）から年月（YYYYMM）を求め、
' 同じ年月の行が既に月次ログにあれば丸ごと入れ替える（＝同じ月を
' 再実行しても重複しない）。
'
' 月次ログが2か月分以上たまったら、そのデータ範囲を選択して
' 「挿入」→「ピボットテーブル」で月別の集計表を新規作成できる
' （行=氏名、列=年月、値=就業時間など、フィルタ=所属名）。
'============================================================
Sub 月次ログ更新()
    Const COLCNT As Long = 47 '所属ｺｰﾄﾞ～育児時間短（貼り付けシートの列数）

    Dim wsSrc As Worksheet, wsLog As Worksheet
    Dim srcData As Variant, oldData As Variant, newData As Variant
    Dim lastRow As Long, logLastRow As Long
    Dim i As Long, k As Long, ym As Long, keepCount As Long
    Dim ymSet As Object
    Dim keepRows() As Long

    On Error GoTo ErrHandler

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set wsSrc = Sheets("貼り付けシート")
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "貼り付けシートにデータがありません。", vbExclamation
        GoTo CleanUp
    End If
    srcData = wsSrc.Range(wsSrc.Cells(2, 1), wsSrc.Cells(lastRow, COLCNT)).Value

    '今回のデータに含まれる年月の一覧
    Set ymSet = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(srcData, 1)
        ym = Int(CLng(srcData(i, 7)) / 100)
        If Not ymSet.Exists(ym) Then ymSet.Add ym, True
    Next i

    '月次ログシートが無ければ作成（見出しをコピーし「年月」列を追加）
    If Not SheetExists("月次ログ") Then
        Set wsLog = Sheets.Add(After:=Sheets(Sheets.Count))
        wsLog.Name = "月次ログ"
        wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(1, COLCNT)).Copy
        wsLog.Range("A1").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        wsLog.Cells(1, COLCNT + 1) = "年月"
    Else
        Set wsLog = Sheets("月次ログ")
    End If

    '既存ログのうち、今回と重複しない年月の行だけ残す
    keepCount = 0
    logLastRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row
    If logLastRow >= 2 Then
        oldData = wsLog.Range(wsLog.Cells(2, 1), wsLog.Cells(logLastRow, COLCNT + 1)).Value
        ReDim keepRows(1 To UBound(oldData, 1))
        For i = 1 To UBound(oldData, 1)
            If Not ymSet.Exists(CLng(oldData(i, COLCNT + 1))) Then
                keepCount = keepCount + 1
                keepRows(keepCount) = i
            End If
        Next i
    End If

    '残す行＋今回のデータを1枚の配列にまとめる
    ReDim newData(1 To keepCount + UBound(srcData, 1), 1 To COLCNT + 1)
    For i = 1 To keepCount
        For k = 1 To COLCNT + 1
            newData(i, k) = oldData(keepRows(i), k)
        Next k
    Next i
    For i = 1 To UBound(srcData, 1)
        For k = 1 To COLCNT
            newData(keepCount + i, k) = srcData(i, k)
        Next k
        newData(keepCount + i, COLCNT + 1) = Int(CLng(srcData(i, 7)) / 100)
    Next i

    '月次ログを書き直す（一括読み書きなので蓄積が増えても高速）
    wsLog.Range(wsLog.Cells(2, 1), wsLog.Cells(wsLog.Rows.Count, COLCNT + 1)).ClearContents
    wsLog.Range(wsLog.Cells(2, 1), wsLog.Cells(UBound(newData, 1) + 1, COLCNT + 1)).Value = newData

    MsgBox "月次ログを更新しました（累計 " & UBound(newData, 1) & " 件）。" & vbCrLf & _
        "月別ピボットが未作成の場合は、月次ログのデータ範囲を選択して" & vbCrLf & _
        "「挿入」→「ピボットテーブル」で作成してください。", vbInformation

CleanUp:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "エラーが発生しました：" & Err.Description, vbExclamation
End Sub
