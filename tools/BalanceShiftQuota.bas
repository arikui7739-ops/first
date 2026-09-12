Attribute VB_Name = "BalanceShiftQuota"
Option Explicit

' 週契約日数（AL列）に実際の出勤日数をぴったり合わせるマクロ群。
'
' 基本方針（固定）:
'   ・AL列（週契約日数）＝絶対に合わせる対象
'   ・計画人数（AM/PM）＝参考値のみ。超過してもよい。このマクロでは考慮しない
'   ・「希望履歴」シートに記録された本人希望日は、絶対に変更しない
'   ・所休・法休・休日・休職のセルは、出勤日数にカウントしない
'     （欠勤・有休・調整休・夏休・出張・実際の勤務時間は出勤日数にカウントする）
'
' 使い方:
'   1) 対象月のシート（例:「11月」）をアクティブにする
'   2) BalanceShiftQuota を実行 → 日付行の先頭セル（1日目の日付）をクリック
'      → 前の期間のシート名を聞かれたら入力（先頭週が7日に満たない場合のみ）
'   3) 完了メッセージで、追加した件数と、埋めきれなかった件数（本人希望等と
'      衝突して調整できなかったケース）が表示される
'
'   確認だけしたい場合は CheckShiftQuota を実行（何も書き換えない）。
'
' 「希望履歴」シートの形式（1行目は見出し、2行目以降がデータ）:
'   A列: 氏名（対象シートのA列と完全一致する表記）
'   B列: 日付（"9/19" のような 月/日 の文字列）
'   C列: 内容（"休日" "有休" "欠勤" など。マクロ自体はC列を読みません。
'              B列にある日付＝本人希望として保護する、という使い方です）
'
' 前提（このブックのシート構成に合わせた決め事）:
'   ・日付の「日」が入っている行のすぐ下の行に、曜日（月・火・…・日）が入っている
'   ・年は3行目・対象シートの1列目付近、月は3行目の日付先頭列と同じ列、
'     月をまたぐ場合は日付が前の値より小さくなる列（切り替わり列）の3行目に次の月が入っている
'   ・AL列（週契約日数）は38列目固定

Private Const AL_COLUMN As Long = 38

Sub BalanceShiftQuota()
    Dim ws As Worksheet, wsHist As Worksheet, wsPrev As Worksheet
    Dim dateHeaderCell As Range
    Dim dateRow As Long, wdRow As Long, firstCol As Long, lastCol As Long, splitCol As Long
    Dim startMonth As Long, secondMonth As Long
    Dim empRows() As Long, empCount As Long
    Dim prevName As String
    Dim filled As Long, shortReport As String, shortCount As Long

    Set ws = ActiveSheet

    On Error Resume Next
    Set wsHist = ThisWorkbook.Sheets("希望履歴")
    On Error GoTo 0
    If wsHist Is Nothing Then
        MsgBox "「希望履歴」シートが見つかりません。氏名・日付・内容の3列を持つシートを" & _
               "「希望履歴」という名前で用意してから、もう一度実行してください。", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Set dateHeaderCell = Application.InputBox( _
        "「" & ws.Name & "」シートで、日付が入っている一番左のセル（1日目の日付）をクリックしてOKを押してください。", _
        "日付行の指定", Type:=8)
    On Error GoTo 0
    If dateHeaderCell Is Nothing Then Exit Sub

    dateRow = dateHeaderCell.Row
    wdRow = dateRow + 1
    firstCol = dateHeaderCell.Column
    lastCol = ws.Cells(dateRow, ws.Columns.Count).End(xlToLeft).Column

    FindSplitAndMonths ws, dateRow, firstCol, lastCol, splitCol, startMonth, secondMonth

    prevName = InputBox("先頭週が7日に満たない場合、不足分を補う前の期間のシート名を入力してください。" & vbCrLf & _
                         "不要であれば空欄のままOKを押してください。（例: 10月）", "前の期間のシート名")
    If prevName <> "" Then
        On Error Resume Next
        Set wsPrev = ThisWorkbook.Sheets(prevName)
        On Error GoTo 0
        If wsPrev Is Nothing Then
            MsgBox "シート「" & prevName & "」が見つからなかったため、先頭週の補完なしで進めます。", vbInformation
        End If
    End If

    CollectEmployeeRows ws, empRows, empCount
    If empCount = 0 Then
        MsgBox "対象者が見つかりませんでした（A列に氏名、AL列に数値がある行が対象です）。", vbExclamation
        Exit Sub
    End If

    Dim weeks As Collection
    Set weeks = BuildWeeks(ws, dateRow, wdRow, firstCol, lastCol)

    Dim histMap As Object
    Set histMap = BuildHistoryMap(wsHist, ws, dateRow, firstCol, lastCol, splitCol, startMonth, secondMonth)

    filled = 0
    shortCount = 0
    shortReport = ""

    Dim i As Long, wk As Collection, r As Long
    Dim al As Variant, cnt As Long, need As Long, got As Long
    Dim firstWeek As Collection
    Set firstWeek = weeks(1)

    For i = 1 To empCount
        r = empRows(i)
        al = ws.Cells(r, AL_COLUMN).Value
        If Not IsNumeric(al) Then GoTo NextEmp

        For Each wk In weeks
            cnt = CountQuotaDays(ws, r, wk)
            If wk.Item(1) = firstWeek.Item(1) And Not (wsPrev Is Nothing) And wk.Count < 7 Then
                cnt = cnt + PrevPeriodExtraDays(wsPrev, ws.Cells(r, 1).Value, 7 - wk.Count)
            End If

            If cnt < CLng(al) Then
                need = CLng(al) - cnt
                got = FillShortfall(ws, r, wk, wdRow, need, histMap, firstCol, lastCol)
                filled = filled + got
                If got < need Then
                    shortCount = shortCount + 1
                    shortReport = shortReport & ws.Cells(r, 1).Value & " (" & _
                        FormatWeekLabel(ws, dateRow, wk) & ") 実績" & (cnt + got) & _
                        "/契約" & al & vbCrLf
                End If
            End If
        Next wk
NextEmp:
    Next i

    Dim msg As String
    msg = "追加で埋めたセル数: " & filled
    If shortCount > 0 Then
        msg = msg & vbCrLf & vbCrLf & "本人希望等と衝突し、契約日数に届かなかった週: " & shortCount & "件" & vbCrLf & shortReport
    End If
    MsgBox msg, vbInformation
End Sub

Sub CheckShiftQuota()
    Dim ws As Worksheet
    Dim dateHeaderCell As Range
    Dim dateRow As Long, wdRow As Long, firstCol As Long, lastCol As Long
    Dim empRows() As Long, empCount As Long

    Set ws = ActiveSheet

    On Error Resume Next
    Set dateHeaderCell = Application.InputBox( _
        "「" & ws.Name & "」シートで、日付が入っている一番左のセル（1日目の日付）をクリックしてOKを押してください。", _
        "日付行の指定", Type:=8)
    On Error GoTo 0
    If dateHeaderCell Is Nothing Then Exit Sub

    dateRow = dateHeaderCell.Row
    wdRow = dateRow + 1
    firstCol = dateHeaderCell.Column
    lastCol = ws.Cells(dateRow, ws.Columns.Count).End(xlToLeft).Column

    CollectEmployeeRows ws, empRows, empCount
    If empCount = 0 Then
        MsgBox "対象者が見つかりませんでした。", vbExclamation
        Exit Sub
    End If

    Dim weeks As Collection
    Set weeks = BuildWeeks(ws, dateRow, wdRow, firstCol, lastCol)

    Dim report As String, mismatchCount As Long
    Dim i As Long, wk As Collection, r As Long, al As Variant, cnt As Long

    For i = 1 To empCount
        r = empRows(i)
        al = ws.Cells(r, AL_COLUMN).Value
        If Not IsNumeric(al) Then GoTo NextEmp
        For Each wk In weeks
            cnt = CountQuotaDays(ws, r, wk)
            If cnt <> CLng(al) Then
                mismatchCount = mismatchCount + 1
                report = report & ws.Cells(r, 1).Value & " (" & FormatWeekLabel(ws, dateRow, wk) & ") " & _
                         "実績" & cnt & " / 契約" & al & vbCrLf
            End If
        Next wk
NextEmp:
    Next i

    If mismatchCount = 0 Then
        MsgBox "全員・全週で週契約日数と一致しています。" & vbCrLf & _
               "（先頭週については前の期間シートの分を考慮していない単純チェックです）", vbInformation
    Else
        MsgBox "週契約日数と一致していない週: " & mismatchCount & "件" & vbCrLf & vbCrLf & report, vbExclamation
    End If
End Sub

' ============ 内部処理 ============

Private Sub CollectEmployeeRows(ws As Worksheet, ByRef empRows() As Long, ByRef empCount As Long)
    Dim lastRow As Long, r As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    ReDim empRows(1 To lastRow)
    empCount = 0
    For r = 1 To lastRow
        If Trim(ws.Cells(r, 1).Value & "") <> "" And IsNumeric(ws.Cells(r, AL_COLUMN).Value) Then
            empCount = empCount + 1
            empRows(empCount) = r
        End If
    Next r
End Sub

' 日付行を左から見て、直前までに見た数値より小さい値になった列を「切り替わり列」とする。
' （途中に日付の入っていない空白列（例:月末で存在しない日）が挟まっていても対応できるよう、
'   直前の列ではなく「それまでに見た最後の数値」と比較する）
Private Sub FindSplitAndMonths(ws As Worksheet, dateRow As Long, firstCol As Long, lastCol As Long, _
                                ByRef splitCol As Long, ByRef startMonth As Long, ByRef secondMonth As Long)
    Dim c As Long
    Dim lastVal As Double
    Dim haveLast As Boolean
    splitCol = 0
    haveLast = False
    For c = firstCol To lastCol
        If IsNumeric(ws.Cells(dateRow, c).Value) Then
            If haveLast Then
                If ws.Cells(dateRow, c).Value < lastVal Then
                    splitCol = c
                    Exit For
                End If
            End If
            lastVal = ws.Cells(dateRow, c).Value
            haveLast = True
        End If
    Next c

    startMonth = 0
    secondMonth = 0
    If IsNumeric(ws.Cells(3, firstCol).Value) Then startMonth = CLng(ws.Cells(3, firstCol).Value)
    If splitCol > 0 Then
        If IsNumeric(ws.Cells(3, splitCol).Value) Then secondMonth = CLng(ws.Cells(3, splitCol).Value)
    End If
End Sub

' 週(日～土)ごとの列番号を積んだ Collection の Collection を返す。
' 日付が入っていない列（例:月末で存在しない日の空白列）は週の日数に数えないよう読み飛ばす。
Private Function BuildWeeks(ws As Worksheet, dateRow As Long, wdRow As Long, firstCol As Long, lastCol As Long) As Collection
    Dim weeks As New Collection
    Dim wk As Collection
    Dim c As Long, wd As String

    Set wk = New Collection
    For c = firstCol To lastCol
        If IsNumeric(ws.Cells(dateRow, c).Value) Then
            wd = ws.Cells(wdRow, c).Value
            If wd = "日" And wk.Count > 0 Then
                weeks.Add wk
                Set wk = New Collection
            End If
            wk.Add c
        End If
    Next c
    If wk.Count > 0 Then weeks.Add wk
    Set BuildWeeks = weeks
End Function

Private Function IsOffOnly(v As Variant) As Boolean
    Dim s As String
    s = Trim(v & "")
    IsOffOnly = (s = "所休" Or s = "法休" Or s = "休日" Or s = "休職")
End Function

Private Function IsAdjustableOff(v As Variant) As Boolean
    Dim s As String
    s = Trim(v & "")
    IsAdjustableOff = (s = "所休" Or s = "休日")
End Function

Private Function CountQuotaDays(ws As Worksheet, r As Long, wk As Collection) As Long
    Dim c As Variant, cnt As Long
    For Each c In wk
        If Trim(ws.Cells(r, CLng(c)).Value & "") <> "" And Not IsOffOnly(ws.Cells(r, CLng(c)).Value) Then
            cnt = cnt + 1
        End If
    Next c
    CountQuotaDays = cnt
End Function

Private Function TypicalTime(ws As Worksheet, r As Long, firstCol As Long, lastCol As Long) As String
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim c As Long, v As String
    For c = firstCol To lastCol
        v = Trim(ws.Cells(r, c).Value & "")
        If v <> "" And Not IsOffOnly(v) And (InStr(v, ":") > 0 Or InStr(v, "：") > 0) Then
            If Not d.Exists(v) Then d.Add v, 0
            d(v) = d(v) + 1
        End If
    Next c
    Dim bestKey As String, bestCount As Long, k As Variant
    bestCount = -1
    For Each k In d.Keys
        If d(k) > bestCount Then
            bestCount = d(k)
            bestKey = k
        End If
    Next k
    TypicalTime = bestKey
End Function

Private Function BuildHistoryMap(wsHist As Worksheet, ws As Worksheet, dateRow As Long, _
                                  firstCol As Long, lastCol As Long, splitCol As Long, _
                                  startMonth As Long, secondMonth As Long) As Object
    Dim map As Object
    Set map = CreateObject("Scripting.Dictionary")

    Dim nameToRow As Object
    Set nameToRow = CreateObject("Scripting.Dictionary")
    Dim r As Long, lastRow As Long, nm As String
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 1 To lastRow
        nm = Trim(ws.Cells(r, 1).Value & "")
        If nm <> "" And Not nameToRow.Exists(nm) Then nameToRow.Add nm, r
    Next r

    Dim hLastRow As Long, i As Long
    hLastRow = wsHist.Cells(wsHist.Rows.Count, 1).End(xlUp).Row
    For i = 2 To hLastRow
        Dim hName As String, hDate As String
        hName = Trim(wsHist.Cells(i, 1).Value & "")
        hDate = Trim(wsHist.Cells(i, 2).Value & "")
        If hName = "" Or hDate = "" Then GoTo NextHist
        If Not nameToRow.Exists(hName) Then GoTo NextHist

        Dim slashPos As Long, monPart As String, dayPart As String
        slashPos = InStr(hDate, "/")
        If slashPos = 0 Then GoTo NextHist
        monPart = Left(hDate, slashPos - 1)
        dayPart = Mid(hDate, slashPos + 1)
        If Not IsNumeric(monPart) Or Not IsNumeric(dayPart) Then GoTo NextHist

        Dim targetCol As Long
        targetCol = FindDateColumn(ws, dateRow, firstCol, lastCol, splitCol, startMonth, secondMonth, _
                                    CLng(monPart), CLng(dayPart))
        If targetCol > 0 Then
            map(CStr(nameToRow(hName)) & "|" & CStr(targetCol)) = True
        End If
NextHist:
    Next i

    Set BuildHistoryMap = map
End Function

' 月・日の両方が一致する列を探す（月をまたぐシートでも同じ「日」の値が2回出るため必須）。
Private Function FindDateColumn(ws As Worksheet, dateRow As Long, firstCol As Long, lastCol As Long, _
                                 splitCol As Long, startMonth As Long, secondMonth As Long, _
                                 wantMonth As Long, wantDay As Long) As Long
    Dim c As Long, colMonth As Long
    For c = firstCol To lastCol
        If splitCol > 0 And c >= splitCol Then
            colMonth = secondMonth
        Else
            colMonth = startMonth
        End If
        If colMonth = wantMonth And ws.Cells(dateRow, c).Value = wantDay Then
            FindDateColumn = c
            Exit Function
        End If
    Next c
    FindDateColumn = 0
End Function

Private Function IsProtected(histMap As Object, r As Long, c As Long) As Boolean
    IsProtected = histMap.Exists(CStr(r) & "|" & CStr(c))
End Function

' 不足日数分を、調整可能な休み（所休/休日、本人希望以外）から埋める。埋めた件数を返す。
' 月・火を優先して埋める（それ以外は日付の早い順）。
' sheetFirstCol/sheetLastCol はシート全体の日付範囲（典型的な勤務時間を探す用）。
Private Function FillShortfall(ws As Worksheet, r As Long, wk As Collection, wdRow As Long, _
                                need As Long, histMap As Object, _
                                sheetFirstCol As Long, sheetLastCol As Long) As Long
    Dim cnt As Long
    Dim cols() As Long
    ReDim cols(1 To wk.Count)
    cnt = 0
    Dim c As Variant
    For Each c In wk
        If IsAdjustableOff(ws.Cells(r, CLng(c)).Value) And Not IsProtected(histMap, r, CLng(c)) Then
            cnt = cnt + 1
            cols(cnt) = CLng(c)
        End If
    Next c

    Dim i As Long, j As Long, tmpCol As Long
    For i = 1 To cnt - 1
        For j = i + 1 To cnt
            If WeekdayPriority(ws.Cells(wdRow, cols(j)).Value) < WeekdayPriority(ws.Cells(wdRow, cols(i)).Value) Then
                tmpCol = cols(i): cols(i) = cols(j): cols(j) = tmpCol
            End If
        Next j
    Next i

    Dim t As String
    t = TypicalTime(ws, r, sheetFirstCol, sheetLastCol)
    If t = "" Then
        FillShortfall = 0
        Exit Function
    End If

    Dim done As Long
    done = 0
    For i = 1 To cnt
        If done >= need Then Exit For
        ws.Cells(r, cols(i)).Value = t
        ApplyQuotaFontColor ws.Cells(r, cols(i)), t
        done = done + 1
    Next i
    FillShortfall = done
End Function

Private Function WeekdayPriority(ByVal wd As String) As Long
    Select Case wd
        Case "月": WeekdayPriority = 1
        Case "火": WeekdayPriority = 2
        Case "水": WeekdayPriority = 3
        Case "木": WeekdayPriority = 4
        Case "金": WeekdayPriority = 5
        Case "土": WeekdayPriority = 6
        Case "日": WeekdayPriority = 7
        Case Else: WeekdayPriority = 9
    End Select
End Function

Private Sub ApplyQuotaFontColor(cell As Range, ByVal value As String)
    Select Case Trim(value)
        Case "所休", "法休"
            cell.Font.Color = RGB(200, 30, 30)
        Case "有休"
            cell.Font.Color = RGB(29, 78, 216)
        Case Else
            cell.Font.Color = RGB(0, 0, 0)
    End Select
End Sub

' 前の期間シートで、同じ氏名の行の末尾 n 日分の「出勤日数」を数える。
Private Function PrevPeriodExtraDays(wsPrev As Worksheet, empName As String, n As Long) As Long
    If n <= 0 Then
        PrevPeriodExtraDays = 0
        Exit Function
    End If
    Dim r As Long, lastRow As Long, foundRow As Long
    lastRow = wsPrev.Cells(wsPrev.Rows.Count, 1).End(xlUp).Row
    foundRow = 0
    For r = 1 To lastRow
        If Trim(wsPrev.Cells(r, 1).Value & "") = Trim(empName) Then
            foundRow = r
            Exit For
        End If
    Next r
    If foundRow = 0 Then
        PrevPeriodExtraDays = 0
        Exit Function
    End If

    Dim dateRow As Long, lastCol As Long, firstCol As Long
    dateRow = 4
    firstCol = 2
    lastCol = wsPrev.Cells(dateRow, wsPrev.Columns.Count).End(xlToLeft).Column

    Dim c As Long, cnt As Long, counted As Long
    counted = 0
    cnt = 0
    For c = lastCol To firstCol Step -1
        If IsNumeric(wsPrev.Cells(dateRow, c).Value) Then
            If counted >= n Then Exit For
            If Trim(wsPrev.Cells(foundRow, c).Value & "") <> "" And Not IsOffOnly(wsPrev.Cells(foundRow, c).Value) Then
                cnt = cnt + 1
            End If
            counted = counted + 1
        End If
    Next c
    PrevPeriodExtraDays = cnt
End Function

Private Function FormatWeekLabel(ws As Worksheet, dateRow As Long, wk As Collection) As String
    FormatWeekLabel = ws.Cells(dateRow, CLng(wk.Item(1))).Value & "日～" & ws.Cells(dateRow, CLng(wk.Item(wk.Count))).Value & "日"
End Function
