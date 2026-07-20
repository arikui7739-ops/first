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
