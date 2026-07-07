Attribute VB_Name = "ShiftCodeGenerator"
Option Explicit

' 「8月」等のシフト表の実績を「シフトマスタ出力」の記号に変換し、
' 「貼付用」シートのC3以降に書き出します。
' 実行前に「貼付用」シートのA1(年)・B1(月)が正しいことを確認してください。
'
' 休日系の割当:
'   法休→法休 / 所休→所休 / 有休→有休 / 休日→所休 /
'   欠勤→欠勤他 / 調整休→欠調整 / 夏休→夏期休
'   出張はそれ自体に時刻が記録されていないため、同じ人の前日→翌日の順で
'   実際の勤務時間が入っている方を借用してコード化します。前後とも
'   勤務時間が無い場合や、前後の時間が食い違う場合のみ実行後の一覧に
'   "要確認" として表示されます。

Private Const WORD_HOUJYUU As String = "法休"
Private Const WORD_SHOKYUU As String = "所休"
Private Const WORD_YUUKYUU As String = "有休"
Private Const WORD_KYUUJITSU As String = "休日"
Private Const WORD_KEKKIN As String = "欠勤"
Private Const WORD_CHOUSEI As String = "調整休"
Private Const WORD_NATSUYASUMI As String = "夏休"
Private Const WORD_SHUCCHOU As String = "出張"

Sub GenerateShiftCodes()
    Dim wsPaste As Worksheet, wsSrc As Worksheet, wsMaster As Worksheet
    Dim srcSheetName As String
    Dim dateHeaderCell As Range, nameHeaderCell As Range
    Dim dateHeaderRow As Long, srcFirstDataCol As Long
    Dim nameCol As Long, srcFirstDataRow As Long
    Dim srcLastDataRow As Long, srcLastDataCol As Long
    Dim pasteFirstRow As Long, pasteLastRow As Long
    Dim pasteFirstCol As Long, pasteLastCol As Long
    Dim splitCol As Long, srcSplitCol As Long
    Dim colMap() As Long
    Dim r As Long, c As Long, sc As Long
    Dim empName As String
    Dim foundNameCell As Range
    Dim rawVal As Variant, code As String
    Dim filledCount As Long, skippedEmpCount As Long
    Dim unresolved As String, unresolvedCount As Long
    Dim masterDict As Object

    Set wsPaste = ThisWorkbook.Sheets("貼付用")
    Set wsMaster = ThisWorkbook.Sheets("シフトマスタ出力")

    srcSheetName = InputBox("元になるシフト表のシート名を入力してください（例: 8月）", "参照シート")
    If srcSheetName = "" Then Exit Sub
    On Error Resume Next
    Set wsSrc = ThisWorkbook.Sheets(srcSheetName)
    On Error GoTo 0
    If wsSrc Is Nothing Then
        MsgBox "シート「" & srcSheetName & "」が見つかりません。", vbExclamation
        Exit Sub
    End If

    wsSrc.Activate
    On Error Resume Next
    Set dateHeaderCell = Application.InputBox( _
        "「" & srcSheetName & "」で、日付が入っている一番左のセル（1日目の日付）をクリックしてOKを押してください。", _
        "日付行の指定", Type:=8)
    On Error GoTo 0
    If dateHeaderCell Is Nothing Then Exit Sub

    On Error Resume Next
    Set nameHeaderCell = Application.InputBox( _
        "「" & srcSheetName & "」で、1人目の氏名が入っているセルをクリックしてOKを押してください。", _
        "氏名列の指定", Type:=8)
    On Error GoTo 0
    If nameHeaderCell Is Nothing Then Exit Sub

    dateHeaderRow = dateHeaderCell.Row
    srcFirstDataCol = dateHeaderCell.Column
    nameCol = nameHeaderCell.Column
    srcFirstDataRow = nameHeaderCell.Row
    srcLastDataRow = wsSrc.Cells(wsSrc.Rows.Count, nameCol).End(xlUp).Row
    srcLastDataCol = wsSrc.Cells(dateHeaderRow, wsSrc.Columns.Count).End(xlToLeft).Column

    srcSplitCol = 0
    For c = srcFirstDataCol + 1 To srcLastDataCol
        If IsNumeric(wsSrc.Cells(dateHeaderRow, c).Value) And IsNumeric(wsSrc.Cells(dateHeaderRow, c - 1).Value) Then
            If wsSrc.Cells(dateHeaderRow, c).Value < wsSrc.Cells(dateHeaderRow, c - 1).Value Then
                srcSplitCol = c
                Exit For
            End If
        End If
    Next c
    If srcSplitCol = 0 Then srcSplitCol = srcLastDataCol + 1

    pasteFirstRow = 3
    pasteLastRow = wsPaste.Cells(wsPaste.Rows.Count, 2).End(xlUp).Row
    pasteFirstCol = 3 ' C
    pasteLastCol = wsPaste.Cells(1, wsPaste.Columns.Count).End(xlToLeft).Column

    splitCol = 0
    For c = pasteFirstCol + 1 To pasteLastCol
        If IsNumeric(wsPaste.Cells(1, c).Value) And IsNumeric(wsPaste.Cells(1, c - 1).Value) Then
            If wsPaste.Cells(1, c).Value < wsPaste.Cells(1, c - 1).Value Then
                splitCol = c
                Exit For
            End If
        End If
    Next c
    If splitCol = 0 Then splitCol = pasteLastCol + 1

    ' 貼付用の各日付列 → 元シートの対応する列を先に1回だけ計算しておく
    ReDim colMap(pasteFirstCol To pasteLastCol)
    For c = pasteFirstCol To pasteLastCol
        Dim dayNum As Variant, isSecondHalf As Boolean, found As Long
        dayNum = wsPaste.Cells(1, c).Value
        isSecondHalf = (c >= splitCol)
        found = 0
        For sc = srcFirstDataCol To srcLastDataCol
            If (sc >= srcSplitCol) = isSecondHalf Then
                If wsSrc.Cells(dateHeaderRow, sc).Value = dayNum Then
                    found = sc
                    Exit For
                End If
            End If
        Next sc
        colMap(c) = found
    Next c

    Set masterDict = BuildMasterDictionary(wsMaster)

    filledCount = 0
    skippedEmpCount = 0
    unresolved = ""
    unresolvedCount = 0

    For r = pasteFirstRow To pasteLastRow
        empName = Trim(wsPaste.Cells(r, 2).Value)
        If empName = "" Then GoTo NextEmp

        Set foundNameCell = wsSrc.Range( _
            wsSrc.Cells(srcFirstDataRow, nameCol), wsSrc.Cells(srcLastDataRow, nameCol)).Find( _
            What:=empName, LookIn:=xlValues, LookAt:=xlWhole)

        If foundNameCell Is Nothing Then
            skippedEmpCount = skippedEmpCount + 1
            GoTo NextEmp
        End If

        For c = pasteFirstCol To pasteLastCol
            If colMap(c) = 0 Then GoTo NextDay
            rawVal = wsSrc.Cells(foundNameCell.Row, colMap(c)).Value
            If Trim(rawVal & "") = "" Then GoTo NextDay

            code = ConvertToShiftCode(CStr(rawVal), masterDict)

            If code = "" And Trim(CStr(rawVal)) = WORD_SHUCCHOU Then
                code = InferShuchoCode(wsSrc, foundNameCell.Row, colMap, c, pasteFirstCol, pasteLastCol, masterDict)
                If code = "AMBIGUOUS" Then
                    code = ""
                    unresolvedCount = unresolvedCount + 1
                    If unresolvedCount <= 30 Then
                        unresolved = unresolved & empName & " " & wsPaste.Cells(1, c).Value & "日: " & rawVal & "（前後で勤務時間が食い違うため要確認）" & vbCrLf
                    End If
                    GoTo NextDay
                End If
            End If

            If code = "" Then
                unresolvedCount = unresolvedCount + 1
                If unresolvedCount <= 30 Then
                    unresolved = unresolved & empName & " " & wsPaste.Cells(1, c).Value & "日: " & rawVal & vbCrLf
                End If
            Else
                wsPaste.Cells(r, c).Value = code
                filledCount = filledCount + 1
            End If
NextDay:
        Next c
NextEmp:
    Next r

    Dim msg As String
    msg = filledCount & " 件のコードを入力しました。"
    If skippedEmpCount > 0 Then
        msg = msg & vbCrLf & "「" & srcSheetName & "」に見つからなかった氏名: " & skippedEmpCount & " 名（対象外として空欄のままです）"
    End If
    If unresolvedCount > 0 Then
        msg = msg & vbCrLf & vbCrLf & "変換できなかったセル: " & unresolvedCount & " 件" & vbCrLf & unresolved
        If unresolvedCount > 30 Then msg = msg & "…ほか" & (unresolvedCount - 30) & "件"
    End If
    MsgBox msg, vbInformation
End Sub

Private Function BuildMasterDictionary(wsMaster As Worksheet) As Object
    Dim dict As Object
    Dim r As Long, lastRow As Long
    Dim sym As String
    Dim st As Variant, en As Variant
    Dim key As String

    Set dict = CreateObject("Scripting.Dictionary")
    lastRow = wsMaster.Cells(wsMaster.Rows.Count, 4).End(xlUp).Row

    For r = 2 To lastRow
        sym = Trim(wsMaster.Cells(r, 4).Value & "")
        If sym = "" Then GoTo NextRow
        st = wsMaster.Cells(r, 11).Value
        en = wsMaster.Cells(r, 13).Value
        If IsNumeric(st) And IsNumeric(en) Then
            If CDbl(st) <> 0 Or CDbl(en) <> 0 Then
                key = TimeKey(CDbl(st)) & "|" & TimeKey(CDbl(en))
                If Not dict.Exists(key) Then dict.Add key, sym
            End If
        End If
NextRow:
    Next r
    Set BuildMasterDictionary = dict
End Function

Private Function TimeKey(serialVal As Double) As String
    TimeKey = CStr(CLng(Application.WorksheetFunction.Round(serialVal * 1440, 0)))
End Function

Private Function ConvertToShiftCode(ByVal rawVal As String, masterDict As Object) As String
    Dim v As String
    v = Trim(rawVal)

    Select Case v
        Case WORD_HOUJYUU
            ConvertToShiftCode = "法休": Exit Function
        Case WORD_SHOKYUU
            ConvertToShiftCode = "所休": Exit Function
        Case WORD_YUUKYUU
            ConvertToShiftCode = "有休": Exit Function
        Case WORD_KYUUJITSU
            ConvertToShiftCode = "所休": Exit Function
        Case WORD_KEKKIN
            ConvertToShiftCode = "欠勤他": Exit Function
        Case WORD_CHOUSEI
            ConvertToShiftCode = "欠調整": Exit Function
        Case WORD_NATSUYASUMI
            ConvertToShiftCode = "夏期休": Exit Function
        Case WORD_SHUCCHOU
            ConvertToShiftCode = ""
            Exit Function
    End Select

    ConvertToShiftCode = TimeRangeToCode(v, masterDict)
End Function

' 「休日」等の言葉ではなく、実際の勤務時間の文字列だった場合にのみ
' コードを返す。言葉（休日など）や解析できない文字列には "" を返す。
Private Function TimeRangeToCode(ByVal rawVal As String, masterDict As Object) As String
    Dim startH As Long, startM As Long, endH As Long, endM As Long
    If Not ParseTimeRange(Trim(rawVal), startH, startM, endH, endM) Then
        TimeRangeToCode = ""
        Exit Function
    End If

    Dim startSerial As Double, endSerial As Double
    startSerial = (startH * 60 + startM) / 1440#
    endSerial = (endH * 60 + endM) / 1440#

    Dim key As String
    key = TimeKey(startSerial) & "|" & TimeKey(endSerial)

    If masterDict.Exists(key) Then
        TimeRangeToCode = masterDict(key)
    Else
        TimeRangeToCode = ""
    End If
End Function

' 出張の日を、同じ人の前日→翌日の順で実際の勤務時間から推定する。
' 前後とも勤務時間が取れなければ ""、前後で異なる時間なら "AMBIGUOUS" を返す。
Private Function InferShuchoCode(wsSrc As Worksheet, srcRow As Long, colMap() As Long, _
                                  c As Long, pasteFirstCol As Long, pasteLastCol As Long, _
                                  masterDict As Object) As String
    Dim prevCode As String, nextCode As String

    prevCode = ""
    If c - 1 >= pasteFirstCol Then
        If colMap(c - 1) <> 0 Then
            prevCode = TimeRangeToCode(CStr(wsSrc.Cells(srcRow, colMap(c - 1)).Value), masterDict)
        End If
    End If

    nextCode = ""
    If c + 1 <= pasteLastCol Then
        If colMap(c + 1) <> 0 Then
            nextCode = TimeRangeToCode(CStr(wsSrc.Cells(srcRow, colMap(c + 1)).Value), masterDict)
        End If
    End If

    If prevCode <> "" And nextCode <> "" Then
        If prevCode = nextCode Then
            InferShuchoCode = prevCode
        Else
            InferShuchoCode = "AMBIGUOUS"
        End If
    ElseIf prevCode <> "" Then
        InferShuchoCode = prevCode
    ElseIf nextCode <> "" Then
        InferShuchoCode = nextCode
    Else
        InferShuchoCode = ""
    End If
End Function

Private Function ParseTimeRange(ByVal s As String, ByRef startH As Long, ByRef startM As Long, ByRef endH As Long, ByRef endM As Long) As Boolean
    Dim norm As String
    Dim sep As String
    Dim parts() As String

    norm = Replace(s, "：", ":")

    If InStr(norm, "～") > 0 Then
        sep = "～"
    ElseIf InStr(norm, "－") > 0 Then
        sep = "－"
    ElseIf InStr(norm, "-") > 0 Then
        sep = "-"
    Else
        ParseTimeRange = False
        Exit Function
    End If

    parts = Split(norm, sep)
    If UBound(parts) <> 1 Then
        ParseTimeRange = False
        Exit Function
    End If

    If Not ParseHM(parts(0), startH, startM) Then
        ParseTimeRange = False
        Exit Function
    End If
    If Not ParseHM(parts(1), endH, endM) Then
        ParseTimeRange = False
        Exit Function
    End If

    ParseTimeRange = True
End Function

Private Function ParseHM(ByVal s As String, ByRef h As Long, ByRef m As Long) As Boolean
    Dim p() As String
    s = Trim(s)
    If InStr(s, ":") = 0 Then
        ParseHM = False
        Exit Function
    End If
    p = Split(s, ":")
    If UBound(p) <> 1 Then
        ParseHM = False
        Exit Function
    End If
    If Not IsNumeric(p(0)) Or Not IsNumeric(p(1)) Then
        ParseHM = False
        Exit Function
    End If
    h = CLng(p(0))
    m = CLng(p(1))
    ParseHM = True
End Function
