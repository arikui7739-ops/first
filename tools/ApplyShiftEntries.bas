Attribute VB_Name = "ShiftEntryImport"
Option Explicit

' シフト希望 入力ツールでコピーした内容を、いったん「取込用」シートに
' 普通に貼り付けた（Ctrl+V）あとにこのマクロを実行してください。
' 取込用シートの各行を氏名で本番のシフト表に照合し、値だけを
' PasteSpecial (SkipBlanks:=True) で反映します。空欄だった日は
' 本番シートの既存内容を変更しません。
Sub ApplyShiftEntries()
    Dim wsStaging As Worksheet
    Dim wsTarget As Worksheet
    Dim targetSheetName As String
    Dim nameCell As Range
    Dim firstDateCell As Range
    Dim i As Long, lastStagingRow As Long, lastStagingCol As Long
    Dim empName As String
    Dim foundCell As Range
    Dim srcRange As Range, destRange As Range
    Dim appliedCount As Long, missingNames As String

    On Error Resume Next
    Set wsStaging = ThisWorkbook.Sheets("取込用")
    On Error GoTo 0
    If wsStaging Is Nothing Then
        MsgBox "「取込用」という名前のシートが見つかりません。" & vbCrLf & _
               "シフト入力ツールでコピーした内容を貼り付けたシートの名前を" & _
               "「取込用」にしてから、もう一度実行してください。", vbExclamation
        Exit Sub
    End If

    targetSheetName = InputBox("貼り付け先のシート名を入力してください（例: 8月 ）", "貼り付け先シート")
    If targetSheetName = "" Then Exit Sub

    On Error Resume Next
    Set wsTarget = ThisWorkbook.Sheets(targetSheetName)
    On Error GoTo 0
    If wsTarget Is Nothing Then
        MsgBox "シート「" & targetSheetName & "」が見つかりません。", vbExclamation
        Exit Sub
    End If

    wsTarget.Activate

    On Error Resume Next
    Set nameCell = Application.InputBox( _
        "「" & targetSheetName & "」シートで、1人目の氏名が入っているセルをクリックしてOKを押してください。", _
        "氏名列の指定", Type:=8)
    On Error GoTo 0
    If nameCell Is Nothing Then Exit Sub

    On Error Resume Next
    Set firstDateCell = Application.InputBox( _
        "同じ行で、1日目の勤務欄（一番左の日付列）のセルをクリックしてOKを押してください。", _
        "日付列の指定", Type:=8)
    On Error GoTo 0
    If firstDateCell Is Nothing Then Exit Sub

    lastStagingRow = wsStaging.Cells(wsStaging.Rows.Count, 1).End(xlUp).Row
    appliedCount = 0
    missingNames = ""

    For i = 1 To lastStagingRow
        empName = Trim(wsStaging.Cells(i, 1).Value)
        If empName = "" Then GoTo NextRow

        Set foundCell = wsTarget.Columns(nameCell.Column).Find( _
            What:=empName, LookIn:=xlValues, LookAt:=xlWhole)

        If foundCell Is Nothing Then
            missingNames = missingNames & empName & vbCrLf
            GoTo NextRow
        End If

        lastStagingCol = wsStaging.Cells(i, wsStaging.Columns.Count).End(xlToLeft).Column
        If lastStagingCol < 2 Then GoTo NextRow

        Set srcRange = wsStaging.Range(wsStaging.Cells(i, 2), wsStaging.Cells(i, lastStagingCol))
        Set destRange = wsTarget.Cells(foundCell.Row, firstDateCell.Column)

        srcRange.Copy
        destRange.PasteSpecial Paste:=xlPasteValues, SkipBlanks:=True
        appliedCount = appliedCount + 1
NextRow:
    Next i

    Application.CutCopyMode = False

    Dim msg As String
    msg = appliedCount & " 名分を反映しました。"
    If missingNames <> "" Then
        msg = msg & vbCrLf & vbCrLf & "以下の氏名は「" & targetSheetName & "」に見つかりませんでした:" & vbCrLf & missingNames
    End If
    MsgBox msg, vbInformation
End Sub
