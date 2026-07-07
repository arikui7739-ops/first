Attribute VB_Name = "Module5"
Sub GenerateLocationInstructions()
    Dim wsData As Worksheet, wsOut As Worksheet
    Dim wsTempItems As Worksheet, wsTempSlots As Worksheet
    Dim lastRow As Long, i As Long

    ' --- 1. シートの設定（既定は"CF"シート。操作パネルのボタンから実行しても正しく動作する） ---
    On Error Resume Next
    Set wsData = ActiveWorkbook.Sheets("CF")
    On Error GoTo 0

    If wsData Is Nothing Then
        Set wsData = ActiveSheet
        If MsgBox("「CF」シートが見つかりません。現在のシート（" & wsData.name & "）を処理しますか？", vbYesNo + vbQuestion) = vbNo Then
            Exit Sub
        End If
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    ' --- 2. 列番号の自動特定（重複列にも対応） ---
    Dim cMach As Long, cRow As Long, cCol As Long, cCode As Long, cName As Long, cFreq As Long
    Dim c As Long
    For c = 1 To wsData.Cells(1, wsData.Columns.Count).End(xlToLeft).Column
        Dim headerText As String
        ' 改行やスペースをすべて除去して判定
        headerText = Replace(Replace(Replace(GetSafeText(wsData.Cells(1, c)), " ", ""), "　", ""), vbLf, "")

        ' 最初に見つかった列を採用する（cMach = 0 等の条件で重複上書きを防止）
        If headerText Like "*号機*" And cMach = 0 Then cMach = c
        If headerText Like "*段*" And cRow = 0 Then cRow = c
        If headerText Like "*列*" And cCol = 0 Then cCol = c
        If headerText Like "*品名コード*" And cCode = 0 Then cCode = c
        If headerText Like "*品名*" And Not headerText Like "*コード*" And cName = 0 Then cName = c
        If headerText Like "*予測*" And headerText Like "*回数*" And cFreq = 0 Then cFreq = c
    Next c

    If cMach = 0 Or cRow = 0 Or cCol = 0 Or cCode = 0 Or cFreq = 0 Then
        MsgBox "現在のシートの1行目に必須項目が見つかりません。" & vbCrLf & _
               "「号機」「段」「列」「品名コード」「投入回数_予測」が含まれているか確認してください。", vbExclamation
        GoTo CleanUp
    End If

    ' --- 3. 高速処理用の一時シート作成 ---
    Set wsTempItems = ThisWorkbook.Sheets.Add
    Set wsTempSlots = ThisWorkbook.Sheets.Add
    wsTempItems.Range("A1:I1").Value = Array("Zone", "Freq", "Code", "Name", "OldMach", "OldRow", "OldCol", "OldScore", "Processed")
    wsTempSlots.Range("A1:E1").Value = Array("Zone", "Score", "Mach", "Row", "Col")

    ' 号機・段・列 → 現在の入居品(品名コード・品名) の対応表（移動先に既に何があるかを表示するため）
    Dim dictSlotCode As Object: Set dictSlotCode = CreateObject("Scripting.Dictionary")
    Dim dictSlotName As Object: Set dictSlotName = CreateObject("Scripting.Dictionary")

    ' --- 4. データ抽出と現在地（旧）スコア計算 ---
    lastRow = wsData.Cells(wsData.Rows.Count, cMach).End(xlUp).row
    Dim rItem As Long: rItem = 2

    For i = 2 To lastRow
        Dim mach As Integer: mach = Val(GetSafeText(wsData.Cells(i, cMach)))

        ' 50番台と60番台のエリアのみを対象
        If (mach >= 51 And mach <= 58) Or (mach >= 61 And mach <= 68) Then
            Dim zone As Integer: zone = IIf(mach <= 58, 50, 60)
            Dim rw As Integer: rw = Val(GetSafeText(wsData.Cells(i, cRow)))
            Dim cl As Integer: cl = Val(GetSafeText(wsData.Cells(i, cCol)))

            wsTempItems.Cells(rItem, 1).Value = zone
            wsTempItems.Cells(rItem, 2).Value = Val(GetSafeText(wsData.Cells(i, cFreq)))
            wsTempItems.Cells(rItem, 3).Value = GetCodeValue(GetSafeText(wsData.Cells(i, cCode)))
            wsTempItems.Cells(rItem, 4).Value = GetSafeText(wsData.Cells(i, cName))
            wsTempItems.Cells(rItem, 5).Value = mach
            wsTempItems.Cells(rItem, 6).Value = rw
            wsTempItems.Cells(rItem, 7).Value = cl
            wsTempItems.Cells(rItem, 8).Value = GetPenalty(mach, rw, cl)
            rItem = rItem + 1

            ' 現在このスロットに入居している品目を記録（同じスロットが複数行に出てきた場合は最初の1件を採用）
            Dim slotKey As String: slotKey = CStr(CLng(mach) * 10000& + CLng(rw) * 100& + CLng(cl))
            If Not dictSlotCode.Exists(slotKey) Then
                dictSlotCode.Add slotKey, GetCodeValue(GetSafeText(wsData.Cells(i, cCode)))
                dictSlotName.Add slotKey, GetSafeText(wsData.Cells(i, cName))
            End If
        End If
    Next i

    If rItem = 2 Then
        MsgBox "対象となる50号機・60号機エリアのデータが見つかりませんでした。", vbInformation
        GoTo CleanUp
    End If

    ' アイテムをゾーンごと、予測回数の降順にソート
    With wsTempItems.Sort
        .SortFields.Clear
        .SortFields.Add Key:=wsTempItems.Range("A2:A" & rItem - 1), Order:=xlAscending
        .SortFields.Add Key:=wsTempItems.Range("B2:B" & rItem - 1), Order:=xlDescending
        .SetRange wsTempItems.Range("A1:I" & rItem - 1)
        .Header = xlYes
        .Apply
    End With

    ' --- 5. 空き間口（スロット）の全生成とスコア計算 ---
    Dim m As Integer, r As Integer, co As Integer
    Dim rSlot As Long: rSlot = 2
    For m = 51 To 68
        If (m >= 51 And m <= 58) Or (m >= 61 And m <= 68) Then
            Dim sZone As Integer: sZone = IIf(m <= 58, 50, 60)
            For r = 1 To 4
                For co = 1 To 9
                    wsTempSlots.Cells(rSlot, 1).Value = sZone
                    wsTempSlots.Cells(rSlot, 2).Value = GetPenalty(m, r, co)
                    wsTempSlots.Cells(rSlot, 3).Value = m
                    wsTempSlots.Cells(rSlot, 4).Value = r
                    wsTempSlots.Cells(rSlot, 5).Value = co
                    rSlot = rSlot + 1
                Next co
            Next r
        End If
    Next m

    ' スロットをゾーンごと、スコア昇順にソート
    With wsTempSlots.Sort
        .SortFields.Clear
        .SortFields.Add Key:=wsTempSlots.Range("A2:A" & rSlot - 1), Order:=xlAscending
        .SortFields.Add Key:=wsTempSlots.Range("B2:B" & rSlot - 1), Order:=xlAscending
        .SortFields.Add Key:=wsTempSlots.Range("E2:E" & rSlot - 1), Order:=xlAscending ' 同じスコアなら手前列優先
        .SetRange wsTempSlots.Range("A1:E" & rSlot - 1)
        .Header = xlYes
        .Apply
    End With

    ' --- 6. 新規シートへマッチング結果の出力 ---
    On Error Resume Next
    ThisWorkbook.Sheets("Cバラ動線最適化").Delete
    On Error GoTo 0

    ' 「操作パネル」シートがあればその左側に配置する
    Dim wsPanel5 As Worksheet
    On Error Resume Next
    Set wsPanel5 = ThisWorkbook.Sheets("操作パネル")
    On Error GoTo 0
    If Not wsPanel5 Is Nothing Then
        Set wsOut = ThisWorkbook.Sheets.Add(Before:=wsPanel5)
    Else
        Set wsOut = ThisWorkbook.Sheets.Add
    End If
    wsOut.name = "Cバラ動線最適化"
    wsOut.Columns("A:A").NumberFormat = "0" ' 品名コードは数値表示
    wsOut.Columns("K:K").NumberFormat = "0" ' 現在の入居品コードも数値表示

    wsOut.Range("A1:N1").Value = Array("品名コード", "品名", "投入回数_予測", "旧_号機", "旧_段", "旧_列", "旧_スコア", "新_号機", "新_段", "新_列", "現在の入居品コード", "現在の入居品名", "新_スコア", "改善スコア")

    ' 50と60の開始行（ポインタ）をそれぞれ特定
    Dim ptr50 As Long, ptr60 As Long
    Dim k As Long
    For k = 2 To rSlot - 1
        If Val(wsTempSlots.Cells(k, 1).Value) = 50 And ptr50 = 0 Then ptr50 = k
        If Val(wsTempSlots.Cells(k, 1).Value) = 60 And ptr60 = 0 Then ptr60 = k
    Next k

    Dim outRow As Long: outRow = 2
    Dim iItem As Long

    ' 確実な割り当て処理
    For iItem = 2 To rItem - 1
        Dim targetZone As Integer: targetZone = Val(wsTempItems.Cells(iItem, 1).Value)
        Dim currentPtr As Long

        If targetZone = 50 Then
            currentPtr = ptr50
        Else
            currentPtr = ptr60
        End If

        If currentPtr > 0 And currentPtr < rSlot Then
            If Val(wsTempSlots.Cells(currentPtr, 1).Value) = targetZone Then
                wsOut.Cells(outRow, 1).Value = wsTempItems.Cells(iItem, 3).Value
                wsOut.Cells(outRow, 2).Value = wsTempItems.Cells(iItem, 4).Value
                wsOut.Cells(outRow, 3).Value = wsTempItems.Cells(iItem, 2).Value
                wsOut.Cells(outRow, 4).Value = wsTempItems.Cells(iItem, 5).Value
                wsOut.Cells(outRow, 5).Value = wsTempItems.Cells(iItem, 6).Value
                wsOut.Cells(outRow, 6).Value = wsTempItems.Cells(iItem, 7).Value
                wsOut.Cells(outRow, 7).Value = wsTempItems.Cells(iItem, 8).Value

                Dim newMach As Long, newRow As Long, newCol As Long
                newMach = wsTempSlots.Cells(currentPtr, 3).Value
                newRow = wsTempSlots.Cells(currentPtr, 4).Value
                newCol = wsTempSlots.Cells(currentPtr, 5).Value

                wsOut.Cells(outRow, 8).Value = newMach
                wsOut.Cells(outRow, 9).Value = newRow
                wsOut.Cells(outRow, 10).Value = newCol

                ' 移動先に現在入居している品目を表示（見つからなければ空欄＝現在空き間口）
                Dim newSlotKey As String: newSlotKey = CStr(CLng(newMach) * 10000& + CLng(newRow) * 100& + CLng(newCol))
                If dictSlotCode.Exists(newSlotKey) Then
                    wsOut.Cells(outRow, 11).Value = dictSlotCode(newSlotKey)
                    wsOut.Cells(outRow, 12).Value = dictSlotName(newSlotKey)
                Else
                    wsOut.Cells(outRow, 11).Value = ""
                    wsOut.Cells(outRow, 12).Value = "(空き間口)"
                End If

                wsOut.Cells(outRow, 13).Value = wsTempSlots.Cells(currentPtr, 2).Value

                ' 改善スコア
                wsOut.Cells(outRow, 14).Value = Val(wsTempItems.Cells(iItem, 8).Value) - Val(wsTempSlots.Cells(currentPtr, 2).Value)

                outRow = outRow + 1

                If targetZone = 50 Then
                    ptr50 = ptr50 + 1
                Else
                    ptr60 = ptr60 + 1
                End If
            End If
        End If
    Next iItem

    ' --- 7. 仕上げ（改善スコア降順で並び替え） ---
    If outRow > 2 Then
        With wsOut.Sort
            .SortFields.Clear
            .SortFields.Add Key:=wsOut.Range("N2:N" & outRow - 1), Order:=xlDescending
            .SetRange wsOut.Range("A1:N" & outRow - 1)
            .Header = xlYes
            .Apply
        End With
    End If

    wsOut.Range("A1:N1").Interior.Color = RGB(200, 230, 255)
    wsOut.Range("A1:N1").Font.Bold = True
    wsOut.Columns("A:N").AutoFit

    ' ※本マクロは予測データベースのため、KPI記録（実績データのみ対象）には記録しません

    MsgBox "「Cバラ動線最適化」の作成が完了しました。", vbInformation

CleanUp:
    On Error Resume Next
    wsTempItems.Delete
    wsTempSlots.Delete
    On Error GoTo 0

    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

' ----------------------------------------------------
' 補助関数群
' ----------------------------------------------------
Function GetSafeText(rng As Range) As String
    If IsError(rng.Value) Then
        GetSafeText = "0"
    ElseIf IsEmpty(rng.Value) Then
        GetSafeText = "0"
    Else
        GetSafeText = Trim(CStr(rng.Value))
    End If
End Function

' 品名コードを数値として保持する（数値でない場合は文字列のまま返す）
Function GetCodeValue(codeText As String) As Variant
    If IsNumeric(codeText) Then
        GetCodeValue = CLng(codeText)
    Else
        GetCodeValue = codeText
    End If
End Function

Function GetPenalty(mach As Integer, row As Integer, col As Integer) As Integer
    Dim pMach As Integer, pRow As Integer, pCol As Integer

    ' 号機:53,54(および63,64)を最良とし、そこからの近さでペナルティを決定
    Select Case mach
        Case 53, 54, 63, 64: pMach = 0
        Case 52, 55, 62, 65: pMach = 1
        Case 51, 56, 61, 66: pMach = 2
        Case 57, 67: pMach = 3
        Case 58, 68: pMach = 4
        Case Else: pMach = 99
    End Select

    Select Case row
        Case 2, 3: pRow = 0
        Case 4: pRow = 2
        Case 1: pRow = 3
        Case Else: pRow = 99
    End Select

    ' 列:1～3列目を最良、4～6列目を中間、7～9列目を最悪の3段階に変更
    Select Case col
        Case 1 To 3: pCol = 0
        Case 4 To 6: pCol = 1
        Case 7 To 9: pCol = 2
        Case Else: pCol = 99
    End Select

    GetPenalty = pMach + pRow + pCol
End Function
