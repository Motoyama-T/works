Private Sub PathButton_Click()
    Dim Base_path As Variant
    Base_path = Application.GetOpenFilename("エクセルファイル(*.xlsx),*.xlsx")
    If Base_path = False Then
        Exit Sub
    End If

    ThisWorkbook.Worksheets("入力マニュアル").Range("B2").Value = Base_path

End Sub
