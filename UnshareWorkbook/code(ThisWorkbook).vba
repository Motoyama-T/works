Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Dim saveAddress() As Variant
    saveAddress = Array(ThisWorkbook.Name, ActiveSheet.Name, ActiveCell.Address)
    Dim Base_path As Variant
    Base_path = ThisWorkbook.Worksheets("入力マニュアル").Range("B2").Value
    Dim pos As Integer
    pos = InStrRev(Base_path, "\")
    Dim Base_name As String
    Base_name = Mid(Base_path, pos + 1)

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    If Dir(Base_path) <> "" And Dir(Base_path) Like "*データベース*" Then
        Workbooks.Open Base_path
    Else
        MsgBox "『 " & Base_path & " 』にはデータベースが存在しません", Buttons:=vbCritical, Title:="該当データが見つかりません！"
        Exit Sub
    End If
    Dim Base As Workbook
    Set Base = Workbooks(Base_name)
    If Base.ReadOnly Then
        'Debug.Print "読み取り専用です"
        MsgBox "時間を空けて再度お試しください。", Buttons:=vbCritical, Title:="他の人が開いています！"
    Else
        'Debug.Print "読み取り専用じゃないです"
        Dim rngDate As Range
        Dim rngFlag As Range
        Set rngDate = ThisWorkbook.Worksheets("入力マニュアル").Range("C2")
        Set rngFlag = ThisWorkbook.Worksheets("入力マニュアル").Range("D2")
        If rngDate = FileDateTime(Base_path) Then
            'Debug.Print "更新日時一致。上書きします"
            Base.Worksheets("作業進捗").Delete
            ThisWorkbook.Worksheets("作業進捗").Copy before:=Base.Worksheets("タイムスタンプ")
            Base.Worksheets("タイムスタンプ").Delete
            ThisWorkbook.Worksheets("タイムスタンプ").Copy after:=Base.Worksheets("作業進捗")
            Base.Worksheets("タイムスタンプ").Range("A1").Select
            Base.Worksheets("作業進捗").Activate
            Range("A1").Select
            Base.Save
            rngDate.Value = FileDateTime(Base_path)
            rngFlag.Value = "いいえ"
        ElseIf rngDate <> FileDateTime(Base_path) And rngFlag = "はい" Then
            'Debug.Print "更新日時が新しくなっています"
            Dim result As Integer
            result = MsgBox("最新のデータベースを取り込みますか？変更内容は破棄されます。", Buttons:=vbYesNo, Title:="古い情報のままです！")
            If result = vbYes Then
                ThisWorkbook.Worksheets("作業進捗").Delete
                Base.Worksheets("作業進捗").Copy before:=ThisWorkbook.Worksheets("タイムスタンプ")
                ThisWorkbook.Worksheets("タイムスタンプ").Delete
                Base.Worksheets("タイムスタンプ").Copy after:=ThisWorkbook.Worksheets("作業進捗")
                rngDate.Value = FileDateTime(Base_path)
                rngFlag.Value = "いいえ"
            Else
                GoTo continue:
            End If
        Else
            ThisWorkbook.Worksheets("作業進捗").Delete
            Base.Worksheets("作業進捗").Copy before:=ThisWorkbook.Worksheets("タイムスタンプ")
            ThisWorkbook.Worksheets("タイムスタンプ").Delete
            Base.Worksheets("タイムスタンプ").Copy after:=ThisWorkbook.Worksheets("作業進捗")
            rngDate.Value = FileDateTime(Base_path)
            rngFlag.Value = "いいえ"
        End If
        Set rngDate = Nothing
        Set rngFlag = Nothing
    End If
continue:
    Base.Close False
    Workbooks(saveAddress(0)).Worksheets(saveAddress(1)).Activate
    Range(saveAddress(2)).Select
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal target As Range)
    If Sh.Name = "作業進捗" Then
    If target.Count = 1 Then
        Dim x1 As Integer
        Dim y1 As Integer
        Dim x2 As Integer
        Dim y2 As Integer
        x1 = 4
        y1 = 4
        x2 = Sh.Cells.Find("セルの挿入はこのセルより上で行ってください。", lookat:=xlWhole).Row - 1
        y2 = Sh.Cells.Find("全体ファイル数", lookat:=xlWhole).Column - 2
        If target.Row >= x1 And target.Row <= x2 Then
        If target.Column >= y1 And target.Column <= y2 Then
        If target.Row Mod 2 = 0 Then
            If InStr(target.Value, "確認中") <> 0 Then
                ThisWorkbook.Worksheets("タイムスタンプ").Cells.Find(target.Address(False, False)).Offset(0, 1) = Now()
            ElseIf InStr(target.Value, "OK") <> 0 Then
                ThisWorkbook.Worksheets("タイムスタンプ").Cells.Find(target.Address(False, False)).Offset(0, 2) = Now()
            End If
        End If
        End If
        End If
    End If
        ThisWorkbook.Worksheets("入力マニュアル").Range("D2").Value = "はい"
    End If
End Sub


