Attribute VB_Name = "Module1"
Sub SwapLocationsByCorrelationFast_Fix()
    Dim fd As Office.FileDialog
    Dim filePath As String
    Dim fileNo As Integer, textLine As String

    Dim dictItemLoc As Object, dictItemHit As Object, dictItemMach As Object, dictItemZone As Object
    Set dictItemLoc = CreateObject("Scripting.Dictionary")
    Set dictItemHit = CreateObject("Scripting.Dictionary")
    Set dictItemMach = CreateObject("Scripting.Dictionary")
    Set dictItemZone = CreateObject("Scripting.Dictionary")

    Dim dictPairs As Object
    Set dictPairs = CreateObject("Scripting.Dictionary")
    Dim currentOrderItems As Object
    Set currentOrderItems = CreateObject("Scripting.Dictionary")

    ' 0. CFシートから ロケーション→品名・商品コード の対応表を先に作っておく（表示用、検索高速化のため辞書化）
    Dim dictLocName As Object: Set dictLocName = CreateObject("Scripting.Dictionary")
    Dim dictLocCode As Object: Set dictLocCode = CreateObject("Scripting.Dictionary")
    Dim wsCF As Worksheet
    On Error Resume Next
    Set wsCF = ActiveWorkbook.Sheets("CF")
    On Error GoTo 0
    If Not wsCF Is Nothing Then
        Dim lastCF As Long: lastCF = wsCF.Cells(wsCF.Rows.Count, "B").End(xlUp).row
        Dim cf As Long
        For cf = 2 To lastCF
            Dim locCode As String: locCode = Trim(CStr(wsCF.Cells(cf, 2).Value)) ' B列:ロケーション(号機*10000+段*100+列)
            If locCode <> "" And Not dictLocName.Exists(locCode) Then
                dictLocName.Add locCode, CStr(wsCF.Cells(cf, 9).Value) ' I列:品名
                Dim codeVal As Variant ' H列:品名コード（数値として保持。数値でない場合はそのまま）
                If IsNumeric(wsCF.Cells(cf, 8).Value) Then
                    codeVal = CLng(wsCF.Cells(cf, 8).Value)
                Else
                    codeVal = wsCF.Cells(cf, 8).Value
                End If
                dictLocCode.Add locCode, codeVal
            End If
        Next cf
    End If

    ' 1. ファイル選択（複数選択可・全ファイル形式）
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "ピッキング実績ファイルを選択（複数選択可）"
        .Filters.Clear
        .Filters.Add "すべてのファイル", "*.*"
        .AllowMultiSelect = True
        If .Show = False Then Exit Sub
    End With

    ' ★高速化＆画面更新停止
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    ' 2. データの読み込み（選択された全ファイルを順に処理し、集計を積み上げる）
    ' E行フォーマット: "E" + (号機2桁+段2桁+列2桁+出荷本数3桁=9桁)を最大3レコード、13文字間隔で連結
    ' H行はオーダー区切り（H行自体の内容は使わない。後半6桁はオリコンNoだが本マクロでは不使用）
    ' ※オーダーはファイルをまたがない前提のため、ファイルの切れ目でも必ず集計を締める
    Dim fIdx As Long
    Dim latestFileDate As Date: latestFileDate = DateSerial(1900, 1, 1) ' ファイル更新日時（B行から日付が読めない場合のフォールバック）
    Dim latestBDate As Date: latestBDate = DateSerial(1900, 1, 1) ' B行(先頭"B"+8桁日付)から読み取った実績日
    For fIdx = 1 To fd.SelectedItems.Count
        filePath = fd.SelectedItems(fIdx)
        Dim thisFileDate As Date: thisFileDate = FileDateTime(filePath)
        If thisFileDate > latestFileDate Then latestFileDate = thisFileDate
        fileNo = FreeFile
        Dim skipMode As Boolean: skipMode = False ' H99999(各ロケの在庫数サマリー行)配下は読み飛ばす
        Open filePath For Input As #fileNo
        Do While Not EOF(fileNo)
            Line Input #fileNo, textLine
            If Left(textLine, 1) = "B" And Len(textLine) >= 9 Then
                ' B行の2～9文字目(8桁)が実績日(YYYYMMDD)
                Dim bDateStr As String: bDateStr = Mid(textLine, 2, 8)
                If IsNumeric(bDateStr) Then
                    Dim bDate As Date
                    On Error Resume Next
                    bDate = DateSerial(CInt(Left(bDateStr, 4)), CInt(Mid(bDateStr, 5, 2)), CInt(Mid(bDateStr, 7, 2)))
                    On Error GoTo 0
                    If bDate > latestBDate Then latestBDate = bDate
                End If
            ElseIf Left(textLine, 1) = "H" Then
                Call RecordPairsFast(currentOrderItems, dictPairs, dictItemZone)
                currentOrderItems.RemoveAll
                skipMode = (Mid(textLine, 2, 5) = "99999") ' 在庫数サマリー行はオーダーではないため以降のE行を除外
            ElseIf Left(textLine, 1) = "E" And Len(textLine) >= 10 And Not skipMode Then
                Dim slotStart As Long
                For slotStart = 2 To Len(textLine) - 8 Step 13
                    Dim rec As String: rec = Mid(textLine, slotStart, 9)
                    If Trim(rec) <> "" And Len(Trim(rec)) = 9 And IsNumeric(rec) Then
                        Dim mach As Integer, dan As Integer, retsu As Integer
                        mach = Val(Mid(rec, 1, 2))
                        dan = Val(Mid(rec, 3, 2))
                        retsu = Val(Mid(rec, 5, 2))
                        ' 出荷本数(Mid(rec,7,3))は今回のロジックでは未使用（回数ベースでカウント）

                        If (mach >= 51 And mach <= 58) Or (mach >= 61 And mach <= 68) Then
                            ' 商品コードがファイルに存在しないため、ロケーション自体をアイテムキーとして扱う
                            Dim locKey As String: locKey = Format(mach, "00") & Format(dan, "00") & Format(retsu, "00")

                            dictItemLoc(locKey) = mach & "-" & Format(dan, "00") & "-" & Format(retsu, "00")
                            dictItemMach(locKey) = mach
                            dictItemZone(locKey) = IIf(mach <= 58, 50, 60)
                            dictItemHit(locKey) = dictItemHit(locKey) + 1
                            currentOrderItems(locKey) = 1
                        End If
                    End If
                Next slotStart
            End If
        Loop
        Call RecordPairsFast(currentOrderItems, dictPairs, dictItemZone)
        currentOrderItems.RemoveAll
        Close #fileNo
    Next fIdx

    ' 3. ペアスコアの計算と配列化
    Dim pairArr() As Variant
    Dim maxPairs As Long: maxPairs = dictPairs.Count
    If maxPairs = 0 Then maxPairs = 1
    ReDim pairArr(1 To maxPairs, 1 To 7)
    Dim pCnt As Long: pCnt = 0

    Dim pairKey As Variant
    For Each pairKey In dictPairs.Keys
        Dim items() As String: items = Split(pairKey, ",")
        Dim itemA As String: itemA = items(0)
        Dim itemB As String: itemB = items(1)

        Dim machA As Integer: machA = dictItemMach(itemA)
        Dim machB As Integer: machB = dictItemMach(itemB)
        Dim machDist As Integer: machDist = Abs(machA - machB)

        If machDist > 0 Then
            pCnt = pCnt + 1
            Dim coCount As Integer: coCount = dictPairs(pairKey)
            Dim wasteScore As Integer: wasteScore = coCount * machDist

            Dim pMachA As Integer: pMachA = GetMachPenaltyFast(machA)
            Dim pMachB As Integer: pMachB = GetMachPenaltyFast(machB)

            Dim anchorItem As String, moverItem As String
            If pMachA <= pMachB Then
                anchorItem = itemA: moverItem = itemB
            Else
                anchorItem = itemB: moverItem = itemA
            End If

            pairArr(pCnt, 1) = dictItemZone(anchorItem)
            pairArr(pCnt, 2) = wasteScore
            pairArr(pCnt, 3) = coCount
            pairArr(pCnt, 4) = anchorItem
            pairArr(pCnt, 5) = dictItemLoc(anchorItem)
            pairArr(pCnt, 6) = moverItem
            pairArr(pCnt, 7) = dictItemLoc(moverItem)
        End If
    Next pairKey

    If pCnt = 0 Then GoTo RestoreSettings

    ' ★ゼロ落ち防止処置をして一時シートでソート
    Dim wsTemp As Worksheet: Set wsTemp = Sheets.Add
    wsTemp.Columns("D:G").NumberFormat = "@" ' ロケーション列を文字列に

    wsTemp.Range("A1").Resize(pCnt, 7).Value = pairArr
    wsTemp.Sort.SortFields.Clear
    wsTemp.Sort.SortFields.Add Key:=wsTemp.Range("B1:B" & pCnt), Order:=xlDescending
    wsTemp.Sort.SetRange wsTemp.Range("A1:G" & pCnt)
    wsTemp.Sort.Apply
    pairArr = wsTemp.Range("A1:G" & pCnt).Value

    ' 4. ロケーションごとのヒット数昇順化
    Dim itemArr() As Variant, iCnt As Long
    iCnt = dictItemHit.Count
    ReDim itemArr(1 To iCnt, 1 To 4)
    Dim r As Long: r = 1

    Dim k As Variant
    For Each k In dictItemHit.Keys
        itemArr(r, 1) = k
        itemArr(r, 2) = dictItemMach(k)
        itemArr(r, 3) = dictItemZone(k)
        itemArr(r, 4) = dictItemHit(k)
        r = r + 1
    Next k

    wsTemp.Cells.Clear
    wsTemp.Columns("A:A").NumberFormat = "@" ' ロケーションキーを文字列に
    wsTemp.Range("A1").Resize(iCnt, 4).Value = itemArr
    wsTemp.Sort.SortFields.Clear
    wsTemp.Sort.SortFields.Add Key:=wsTemp.Range("D1:D" & iCnt), Order:=xlAscending
    wsTemp.Sort.SetRange wsTemp.Range("A1:D" & iCnt)
    wsTemp.Sort.Apply
    itemArr = wsTemp.Range("A1:D" & iCnt).Value
    wsTemp.Delete

    ' 号機ごとにコレクション化（検索高速化・型エラー防止）
    Dim machItems As Object: Set machItems = CreateObject("Scripting.Dictionary")
    For r = 1 To iCnt
        Dim mStr As String: mStr = CStr(itemArr(r, 2))
        If Not machItems.Exists(mStr) Then machItems.Add mStr, New Collection
        machItems(mStr).Add CStr(itemArr(r, 1))
    Next r

    ' 5. 最適な玉突き交換の決定処理
    Dim outArr() As Variant
    ReDim outArr(1 To pCnt, 1 To 14)
    Dim outCnt As Long: outCnt = 0
    Dim dictSwapped As Object: Set dictSwapped = CreateObject("Scripting.Dictionary")

    For r = 1 To pCnt
        Dim aItem As String: aItem = CStr(pairArr(r, 4))
        Dim mItem As String: mItem = CStr(pairArr(r, 6))

        If Not dictSwapped.Exists(aItem) And Not dictSwapped.Exists(mItem) Then
            Dim aMach As Integer: aMach = dictItemMach(aItem)
            Dim z As Integer: z = pairArr(r, 1)
            Dim targetItem As String: targetItem = ""

            ' 起点の号機、または隣接号機から検索
            Dim offsets As Variant: offsets = Array(0, -1, 1)
            Dim oIdx As Integer
            For oIdx = 0 To 2
                Dim searchM As String: searchM = CStr(aMach + offsets(oIdx))
                If machItems.Exists(searchM) Then
                    Dim candidate As Variant
                    For Each candidate In machItems(searchM)
                        If CStr(candidate) <> aItem And CStr(candidate) <> mItem And Not dictSwapped.Exists(CStr(candidate)) Then
                            targetItem = CStr(candidate)
                            Exit For
                        End If
                    Next candidate
                End If
                If targetItem <> "" Then Exit For
            Next oIdx

            ' 結果を格納
            If targetItem <> "" Then
                outCnt = outCnt + 1
                outArr(outCnt, 1) = z
                outArr(outCnt, 2) = pairArr(r, 2)
                outArr(outCnt, 3) = pairArr(r, 3)
                outArr(outCnt, 4) = GetLocName(dictLocName, aMach, aItem)
                outArr(outCnt, 5) = GetLocCode(dictLocCode, aMach, aItem)
                outArr(outCnt, 6) = dictItemLoc(aItem)
                outArr(outCnt, 7) = GetLocName(dictLocName, dictItemMach(mItem), mItem)
                outArr(outCnt, 8) = GetLocCode(dictLocCode, dictItemMach(mItem), mItem)
                outArr(outCnt, 9) = dictItemLoc(mItem)
                outArr(outCnt, 10) = "⇔"
                outArr(outCnt, 11) = GetLocName(dictLocName, dictItemMach(targetItem), targetItem)
                outArr(outCnt, 12) = GetLocCode(dictLocCode, dictItemMach(targetItem), targetItem)
                outArr(outCnt, 13) = dictItemLoc(targetItem)
                outArr(outCnt, 14) = dictItemHit(targetItem)

                dictSwapped(mItem) = True
                dictSwapped(targetItem) = True
            End If
        End If
    Next r

    ' 6. 一括出力
    If outCnt > 0 Then
        Dim wsOut As Worksheet
        On Error Resume Next
        Sheets("同時ピッキング交換指示書").Delete
        On Error GoTo 0

        ' 「操作パネル」シートがあればその左側に配置する
        Dim wsPanel1 As Worksheet
        On Error Resume Next
        Set wsPanel1 = ThisWorkbook.Sheets("操作パネル")
        On Error GoTo 0
        If Not wsPanel1 Is Nothing Then
            Set wsOut = ThisWorkbook.Sheets.Add(Before:=wsPanel1)
        Else
            Set wsOut = Sheets.Add
        End If
        wsOut.Name = "同時ピッキング交換指示書"

        ' ロケーション列は日付等への誤変換防止のため文字列に、商品コード列は数値表示に統一
        wsOut.Columns("E:E").NumberFormat = "0"
        wsOut.Columns("F:F").NumberFormat = "@"
        wsOut.Columns("H:H").NumberFormat = "0"
        wsOut.Columns("I:I").NumberFormat = "@"
        wsOut.Columns("L:L").NumberFormat = "0"
        wsOut.Columns("M:M").NumberFormat = "@"

        wsOut.Range("A1:N1").Value = Array("ゾーン", "無駄歩行スコア", "同時オーダー回数", "【起点品名】(動かさない)", "起点商品コード", "起点ロケーション", "【相関品名】(遠くから呼ぶ)", "相関商品コード", "相関ロケーション", "⇒交換⇒", "【入替対象品名】(近くの非稼働品)", "入替対象商品コード", "入替ロケーション(起点の近所)", "対象ヒット数")

        wsOut.Range("A2").Resize(outCnt, 14).Value = outArr

        wsOut.Range("A1:N1").Interior.Color = RGB(255, 230, 200)
        wsOut.Range("A1:N1").Font.Bold = True
        wsOut.Range("J:J").HorizontalAlignment = xlCenter
        wsOut.Range("J:J").Font.Bold = True
        wsOut.Columns("A:N").AutoFit

        ' KPI記録：平均無駄歩行スコア・実績日（B行の日付があれば優先、無ければファイル更新日時）
        Dim totalWaste As Double, wi As Long
        totalWaste = 0
        For wi = 1 To outCnt
            totalWaste = totalWaste + outArr(wi, 2)
        Next wi
        Dim reportDate As Date
        If latestBDate > DateSerial(1900, 1, 1) Then
            reportDate = latestBDate
        Else
            reportDate = DateSerial(Year(latestFileDate), Month(latestFileDate), Day(latestFileDate))
        End If
        On Error Resume Next
        Module7.LogModule1Score totalWaste / outCnt, reportDate
        On Error GoTo 0

        MsgBox "修正版の処理が完了しました。（" & fd.SelectedItems.Count & "ファイルを読込／" & outCnt & "件の交換候補）", vbInformation
    Else
        MsgBox "交換可能なペアが見つかりませんでした。", vbExclamation
    End If

RestoreSettings:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' --- 補助サブルーチン ---
Sub RecordPairsFast(currentItems As Object, dictPairs As Object, dictZone As Object)
    If currentItems.Count < 2 Then Exit Sub
    Dim itemsArr() As Variant: itemsArr = currentItems.Keys
    Dim i As Integer, j As Integer
    For i = 0 To UBound(itemsArr) - 1
        For j = i + 1 To UBound(itemsArr)
            Dim item1 As String: item1 = CStr(itemsArr(i))
            Dim item2 As String: item2 = CStr(itemsArr(j))
            If dictZone.Exists(item1) And dictZone.Exists(item2) Then
                If dictZone(item1) = dictZone(item2) Then
                    Dim pairKey As String
                    If item1 < item2 Then pairKey = item1 & "," & item2 Else pairKey = item2 & "," & item1
                    dictPairs(pairKey) = dictPairs(pairKey) + 1
                End If
            End If
        Next j
    Next i
End Sub

Function GetMachPenaltyFast(mach As Integer) As Integer
    Select Case mach
        Case 54, 64: GetMachPenaltyFast = 0
        Case 53, 55, 63, 65: GetMachPenaltyFast = 1
        Case 52, 56, 62, 66: GetMachPenaltyFast = 2
        Case 57, 67: GetMachPenaltyFast = 3
        Case 51, 61: GetMachPenaltyFast = 4
        Case 58, 68: GetMachPenaltyFast = 5
        Case Else: GetMachPenaltyFast = 99
    End Select
End Function

' ロケーションキー(locKey)からCFシートの品名を引く。見つからない場合はロケーション文字列をそのまま返す
Function GetLocName(dictLocName As Object, ByVal mach As Long, locKey As String) As String
    Dim dan As String, retsu As String
    dan = Mid(locKey, 3, 2)
    retsu = Mid(locKey, 5, 2)
    Dim locCode As String
    locCode = CStr(CLng(mach) * 10000& + CLng(Val(dan)) * 100& + CLng(Val(retsu)))
    If dictLocName.Exists(locCode) Then
        GetLocName = dictLocName(locCode)
    Else
        GetLocName = "(品名不明)"
    End If
End Function

' ロケーションキー(locKey)からCFシートの商品コードを引く。見つからない場合は空文字を返す
Function GetLocCode(dictLocCode As Object, ByVal mach As Long, locKey As String) As Variant
    Dim dan As String, retsu As String
    dan = Mid(locKey, 3, 2)
    retsu = Mid(locKey, 5, 2)
    Dim locCode As String
    locCode = CStr(CLng(mach) * 10000& + CLng(Val(dan)) * 100& + CLng(Val(retsu)))
    If dictLocCode.Exists(locCode) Then
        GetLocCode = dictLocCode(locCode)
    Else
        GetLocCode = ""
    End If
End Function
