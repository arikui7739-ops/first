Attribute VB_Name = "処理"
'============================================================
' 変換処理（高速化版・診断版）
' 「型が一致しません」の原因特定用に、エラー発生行(Erl)と
' その行の元データの中身をメッセージ表示するようにしています。
' 原因が分かったら診断コードは外して構いません。
'============================================================
Sub 変換処理()
'作成：2015/7/11
'就業管理システム変更に伴うダウンロードデータの変換
'作成者：915812　泉　新
'高速化：2026対応（診断版）

    Dim ten As Variant       '転記元データ配列（新ツール貼付け）
    Dim outArr As Variant    '出力データ配列
    Dim shu As Variant       '集計区分(対比)配列
    Dim kin As Variant       '勤務区分(対比)配列
    Dim i As Long, Z As Long, kenS As Long
    Dim che As Variant, kyu As Variant, v15 As Variant, v16 As Variant

    On Error GoTo ErrHandler

10  Application.ScreenUpdating = False
20  Application.Calculation = xlCalculationManual
30  Application.EnableEvents = False
40  Application.DisplayAlerts = False

50  Sheets("出力").Cells.ClearContents

60  Sheets("項目対比").Rows("4:4").Copy
70  Sheets("出力").Range("A1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
80  Application.CutCopyMode = False

90  ten = Sheets("新ツール貼付け").Range("A2:AN30001").Value

100 shu = Sheets("集計区分 (対比)").Range("B4:E13").Value
110 kin = Sheets("勤務区分 (対比)").Range("B4:F35").Value

120 kenS = 0
130 For i = 1 To 30000
140     If ten(i, 1) <> "" Then
150         kenS = kenS + 1
        Else
160         Exit For
        End If
170 Next i

180 If kenS > 0 Then
190     ReDim outArr(1 To kenS, 1 To 47)

200     For i = 1 To kenS
210         outArr(i, 1) = ten(i, 7) & ten(i, 9)
220         outArr(i, 2) = ten(i, 8) & " " & ten(i, 10)
230         outArr(i, 3) = ten(i, 1)
240         outArr(i, 4) = ten(i, 2)

250         For Z = 1 To 10
260             If shu(Z, 1) = ten(i, 5) Then
270                 outArr(i, 5) = shu(Z, 3)
280                 outArr(i, 6) = shu(Z, 4)
290                 Exit For
                End If
300         Next Z

310         outArr(i, 7) = Year(ten(i, 11)) * 10000 + Month(ten(i, 11)) * 100 + Day(ten(i, 11))

320         For Z = 1 To 32
330             If kin(Z, 1) = ten(i, 14) Then
340                 outArr(i, 8) = kin(Z, 3)
350                 outArr(i, 9) = kin(Z, 4)
360                 outArr(i, 10) = kin(Z, 5)
370                 Exit For
                End If
380         Next Z

390         outArr(i, 11) = "0"
400         outArr(i, 12) = ""

410         If ten(i, 12) = "" Then
420             outArr(i, 13) = 0
            Else
430             outArr(i, 13) = ten(i, 12)
            End If
440         If ten(i, 13) = "" Then
450             outArr(i, 14) = 0
            Else
460             outArr(i, 14) = ten(i, 13)
            End If

470         kyu = ten(i, 18) + ten(i, 19)
480         v15 = outArr(i, 13) + ((outArr(i, 14) - outArr(i, 13)) / 2)
490         che = v15 * 24 - Application.RoundDown(v15 * 24, 0)
500         If che >= 0.75 Then
510             che = 0.75
            ElseIf che >= 0.5 Then
520             che = 0.5
            ElseIf che >= 0.25 Then
530             che = 0.25
            Else
540             che = 0
            End If
550         v15 = (Application.RoundDown(v15 * 24, 0) + che) / 24
560         v16 = v15 + kyu
570         outArr(i, 15) = v15
580         outArr(i, 16) = v16

590         outArr(i, 17) = 0
600         outArr(i, 18) = ""
610         outArr(i, 19) = ten(i, 16)
620         outArr(i, 20) = 0
630         outArr(i, 21) = ""
640         outArr(i, 22) = ten(i, 17)

650         If outArr(i, 22) >= 1 Then
660             outArr(i, 22) = outArr(i, 22) - 1
670             outArr(i, 20) = 1
680             outArr(i, 21) = "翌"
            End If
690         If outArr(i, 19) = "" Then outArr(i, 19) = 0
700         If outArr(i, 22) = "" Then outArr(i, 22) = 0

710         outArr(i, 23) = 0
720         outArr(i, 24) = ""
730         outArr(i, 25) = 0
740         outArr(i, 26) = 0
750         outArr(i, 27) = ""
760         outArr(i, 28) = 0

770         outArr(i, 29) = ten(i, 26)
780         outArr(i, 30) = ten(i, 27) + ten(i, 38)
790         outArr(i, 31) = ten(i, 28) + ten(i, 29)
800         outArr(i, 32) = ten(i, 30)
810         outArr(i, 33) = ten(i, 31) + ten(i, 32)
820         outArr(i, 34) = 0
830         outArr(i, 35) = ten(i, 39)
840         outArr(i, 36) = 0
850         outArr(i, 37) = 0
860         outArr(i, 38) = ten(i, 33)
870         outArr(i, 39) = ten(i, 34)
880         outArr(i, 40) = ten(i, 35)
890         outArr(i, 41) = ten(i, 36)
900         outArr(i, 42) = ten(i, 37)
910         outArr(i, 43) = 0
920         outArr(i, 44) = ten(i, 23)
930         outArr(i, 45) = ten(i, 22)
940         outArr(i, 46) = ten(i, 20)
950         outArr(i, 47) = ten(i, 21)

960         For Z = 29 To 44
970             outArr(i, Z) = outArr(i, Z) * 24
980         Next Z

990         If outArr(i, 14) >= 1 Then outArr(i, 14) = outArr(i, 14) - 1
1000        If outArr(i, 16) >= 1 Then outArr(i, 16) = outArr(i, 16) - 1
1010    Next i

1020    Sheets("出力").Range(Sheets("出力").Cells(2, 1), Sheets("出力").Cells(kenS + 1, 47)).Value = outArr
    End If

1030 With Sheets("出力")
1040    .Columns("M:P").NumberFormatLocal = "[hh]:mm"
1050    .Range("S:S,V:V").NumberFormatLocal = "[hh]:mm"
1060    With .Columns("AC:AS")
1070        .Style = "Comma [0]"
1080        .NumberFormatLocal = "#,##0.00;[赤]-#,##0.00"
        End With
    End With

    GoTo CleanExit

ErrHandler:
    Dim msg As String, k As Long
    msg = "エラー：" & Err.Description & vbCrLf
    msg = msg & "発生行番号(Erl)：" & Erl & vbCrLf
    msg = msg & "新ツール貼付けの行：" & (i + 1) & "行目" & vbCrLf & vbCrLf

    If Not IsEmpty(ten) And i >= 1 Then
        msg = msg & "【その行(A" & (i + 1) & ":AN" & (i + 1) & ")の中身】" & vbCrLf
        For k = 1 To 40
            msg = msg & k & ":[" & SafeStr(ten(i, k)) & "](" & TypeName(ten(i, k)) & ")  "
            If k Mod 4 = 0 Then msg = msg & vbCrLf
        Next k
    End If

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True

    MsgBox msg, vbExclamation, "診断情報"
    Exit Sub

CleanExit:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True

End Sub

Private Function SafeStr(v As Variant) As String
    On Error Resume Next
    SafeStr = "<変換不可>"
    SafeStr = CStr(v)
End Function
