Attribute VB_Name = "処理"
'============================================================
' 変換処理（高速化版）
' 元の処理内容・出力結果はそのまま、以下を変更して高速化：
'   ・Application.ScreenUpdating / Calculation / EnableEvents をOFF
'   ・Cells()の1セルずつ読み書き → 配列で一括読み込み／一括書き込み
'   ・対比表（集計区分・勤務区分）をループ内でシート参照 → 事前に配列化
'   ・末尾の無意味な自己コピペ（A1→A1）を削除
'   ・重複していたNumberFormatLocal設定（即上書きされる分）を削除
'
' 【重要】新ツール貼付けシートの時刻項目（開始・終了・休憩・実働時間等）は
' "08:30"のような文字列で入っている。元のマクロはCells()という実セルを
' 経由代入していたため、Excelの自動変換で数値の時刻に変わっていたが、
' 高速化版は配列のみで処理するため自動変換が働かない。そのため
' ToTime() で明示的に文字列→時刻の数値(日数の割合)へ変換している。
'============================================================
Sub 変換処理()
'作成：2015/7/11
'就業管理システム変更に伴うダウンロードデータの変換
'作成者：915812　泉　新
'高速化：2026対応

    Dim ten As Variant       '転記元データ配列（新ツール貼付け）
    Dim outArr As Variant    '出力データ配列
    Dim shu As Variant       '集計区分(対比)配列
    Dim kin As Variant       '勤務区分(対比)配列
    Dim i As Long, Z As Long, kenS As Long
    Dim che As Double, kyu As Double, v15 As Double, v16 As Double

    On Error GoTo ErrHandler

    '--- 高速化設定 ---
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    '出力シートクリア処理
    Sheets("出力").Cells.ClearContents

    '出力シート項目書出し
    Sheets("項目対比").Rows("4:4").Copy
    Sheets("出力").Range("A1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False

    '新ツール貼付けシートデータを一括取得（配列化）※最大1000人×30日=30,000行、40列
    ten = Sheets("新ツール貼付け").Range("A2:AN30001").Value

    '対比表を一括取得（配列化）
    shu = Sheets("集計区分 (対比)").Range("B4:E13").Value   '1:区分値 3:集計区分 4:集計区分名
    kin = Sheets("勤務区分 (対比)").Range("B4:F35").Value    '1:区分値 3:ｶﾚﾝﾀﾞｰ名 4:届出 5:届出名

    '件数カウント
    kenS = 0
    For i = 1 To 30000
        If ten(i, 1) <> "" Then
            kenS = kenS + 1
        Else
            Exit For
        End If
    Next i

    If kenS > 0 Then
        ReDim outArr(1 To kenS, 1 To 47)

        For i = 1 To kenS
            outArr(i, 1) = ten(i, 7) & ten(i, 9)                 '所属ｺｰﾄﾞ=事業所所属CD&セクションCD
            outArr(i, 2) = ten(i, 8) & " " & ten(i, 10)          '所属名=事業所所属略名
            outArr(i, 3) = ten(i, 1)                             '社員ｺｰﾄﾞ=社員CD
            outArr(i, 4) = ten(i, 2)                             '氏名=社員名（漢字）

            For Z = 1 To 10
                If shu(Z, 1) = ten(i, 5) Then
                    outArr(i, 5) = shu(Z, 3)  '集計区分
                    outArr(i, 6) = shu(Z, 4)  '集計区分名
                    Exit For
                End If
            Next Z

            outArr(i, 7) = Year(ten(i, 11)) * 10000 + Month(ten(i, 11)) * 100 + Day(ten(i, 11)) '日付

            For Z = 1 To 32
                If kin(Z, 1) = ten(i, 14) Then
                    outArr(i, 8) = kin(Z, 3)   'ｶﾚﾝﾀﾞｰ名
                    outArr(i, 9) = kin(Z, 4)   '届出
                    outArr(i, 10) = kin(Z, 5)  '届出名
                    Exit For
                End If
            Next Z

            outArr(i, 11) = "0" 'ｼﾌﾄｺｰﾄﾞ
            outArr(i, 12) = ""  'ｼﾌﾄ名

            outArr(i, 13) = ToTime(ten(i, 12)) '開始=予定開始
            outArr(i, 14) = ToTime(ten(i, 13)) '終了=予定終了

            '↓★休憩時間補正処理↓
            kyu = ToTime(ten(i, 18)) + ToTime(ten(i, 19))  '休憩＋深夜休憩
            v15 = outArr(i, 13) + ((outArr(i, 14) - outArr(i, 13)) / 2)
            che = v15 * 24 - Application.RoundDown(v15 * 24, 0)
            If che >= 0.75 Then
                che = 0.75
            ElseIf che >= 0.5 Then
                che = 0.5
            ElseIf che >= 0.25 Then
                che = 0.25
            Else
                che = 0
            End If
            v15 = (Application.RoundDown(v15 * 24, 0) + che) / 24 '休憩１開始
            v16 = v15 + kyu                                        '休憩１終了
            '↑★休憩時間補正処理↑
            outArr(i, 15) = v15
            outArr(i, 16) = v16

            outArr(i, 17) = 0                  '出勤１ﾌﾗｸﾞ
            outArr(i, 18) = ""                 '出勤１ﾌﾗｸﾞ名
            outArr(i, 19) = ToTime(ten(i, 16)) '出勤１時刻
            outArr(i, 20) = 0                  '退勤１ﾌﾗｸﾞ
            outArr(i, 21) = ""                 '退勤１ﾌﾗｸﾞ名
            outArr(i, 22) = ToTime(ten(i, 17)) '退勤１時刻

            '↓★退勤時間補正処理↓
            If outArr(i, 22) >= 1 Then
                outArr(i, 22) = outArr(i, 22) - 1
                outArr(i, 20) = 1
                outArr(i, 21) = "翌"
            End If
            '↑★退勤時間補正処理↑

            outArr(i, 23) = 0  '外出１ﾌﾗｸﾞ
            outArr(i, 24) = "" '外出１ﾌﾗｸﾞ名
            outArr(i, 25) = 0  '外出１時刻
            outArr(i, 26) = 0  '戻り１ﾌﾗｸﾞ
            outArr(i, 27) = "" '戻り１ﾌﾗｸﾞ名
            outArr(i, 28) = 0  '戻り１時刻

            outArr(i, 29) = ToTime(ten(i, 26))                        '就業時間
            outArr(i, 30) = ToTime(ten(i, 27)) + ToTime(ten(i, 38))   '実動時間
            outArr(i, 31) = ToTime(ten(i, 28)) + ToTime(ten(i, 29))   '時間外（補正込み）
            outArr(i, 32) = ToTime(ten(i, 30))                        '法定深夜
            outArr(i, 33) = ToTime(ten(i, 31)) + ToTime(ten(i, 32))   '法定休労（補正込み）
            outArr(i, 34) = 0                                          '法定休深夜
            outArr(i, 35) = ToTime(ten(i, 39))                        '遅刻
            outArr(i, 36) = 0                                          '早退
            outArr(i, 37) = 0                                          '外出
            outArr(i, 38) = ToTime(ten(i, 33))                        '早朝時間
            outArr(i, 39) = ToTime(ten(i, 34))                        '基本時間
            outArr(i, 40) = ToTime(ten(i, 35))                        '夜間時間
            outArr(i, 41) = ToTime(ten(i, 36))                        '深夜時間
            outArr(i, 42) = ToTime(ten(i, 37))                        '休日加算
            outArr(i, 43) = 0                                          '25%割増ｈ
            outArr(i, 44) = ToTime(ten(i, 23))                        '育児時間短
            outArr(i, 45) = ten(i, 22)                                 '年末年始
            outArr(i, 46) = ten(i, 20)                                 '運行手当
            outArr(i, 47) = ten(i, 21)                                 '大型手当

            For Z = 29 To 44
                outArr(i, Z) = outArr(i, Z) * 24
            Next Z

            '↓★シフト・休憩終了時刻変換　48　⇒　24↓
            If outArr(i, 14) >= 1 Then outArr(i, 14) = outArr(i, 14) - 1
            If outArr(i, 16) >= 1 Then outArr(i, 16) = outArr(i, 16) - 1
            '↑★シフト・休憩終了時刻変換　48　⇒　24↑
        Next i

        '出力シートへ一括書込み
        Sheets("出力").Range(Sheets("出力").Cells(2, 1), Sheets("出力").Cells(kenS + 1, 47)).Value = outArr
    End If

    '書式設定
    With Sheets("出力")
        .Columns("M:P").NumberFormatLocal = "[hh]:mm"
        .Range("S:S,V:V").NumberFormatLocal = "[hh]:mm"
        With .Columns("AC:AS")
            .Style = "Comma [0]"
            .NumberFormatLocal = "#,##0.00;[赤]-#,##0.00"
        End With
    End With

ErrHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True

    If Err.Number <> 0 Then
        MsgBox "エラーが発生しました：" & Err.Description & vbCrLf & _
            "（新ツール貼付け上の行番号：" & (i + 1) & "行目付近）", vbExclamation
    End If

End Sub

'新ツール貼付けの時刻項目（"08:30"等の文字列、または既に数値/空欄）を
'日数の割合（Excelの時刻シリアル値）に変換する。24時を超える文字列
'（"25:15"等の日またぎ表記）にも対応。
Private Function ToTime(v As Variant) As Double
    Dim s As String, p() As String
    Dim h As Double, m As Double, sec As Double

    If IsEmpty(v) Or v = "" Then
        ToTime = 0
    ElseIf IsNumeric(v) Then
        ToTime = CDbl(v)
    Else
        s = CStr(v)
        p = Split(s, ":")
        h = CDbl(p(0))
        If UBound(p) >= 1 Then m = CDbl(p(1)) Else m = 0
        If UBound(p) >= 2 Then sec = CDbl(p(2)) Else sec = 0
        ToTime = (h * 3600 + m * 60 + sec) / 86400
    End If
End Function
