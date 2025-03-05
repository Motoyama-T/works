Public LastRow1, LastRow2, MaxRow, MinRow, LastLen1, LastLen2, MaxLen, MinLen As Integer

Sub setResult()
 setVariable
 If LastLen1 = 0 Then
  MsgBox "テキスト1がありません"
  Exit Sub
 ElseIf LastLen2 = 0 Then
  MsgBox "テキスト2がありません"
  Exit Sub
 End If
 
 ReDim A1(1 To LastRow1, 1 To 53), A2(1 To LastRow2, 1 To 53) As Variant
 Dim a, b, c, d As Integer
 
 For x = 1 To LastRow1
  a = a + 1
  b = 0
  For y = 1 To Len(Sheets("テキスト1").Cells(x, 1))
   If y - b > 53 Then
    a = a + 1
    b = b + 53
   End If
   A1(a, y - b) = Mid(Sheets("テキスト1").Cells(x, 1), y, 1)
  Next
 Next
 For x = 1 To LastRow2
  c = c + 1
  d = 0
  For y = 1 To Len(Sheets("テキスト2").Cells(x, 1))
   If y - d > 53 Then
    c = c + 1
    d = d + 53
   End If
   A2(c, y - d) = Mid(Sheets("テキスト2").Cells(x, 1), y, 1)
  Next
 Next
 Sheets("結果").Range(Sheets("結果").Cells(2, 1), Sheets("結果").Cells(UBound(A1, 1) + 1, 53)) = A1
 Sheets("結果").Range(Sheets("結果").Cells(2, 55), Sheets("結果").Cells(UBound(A2, 1) + 1, 107)) = A2

 For x = 2 To MaxRow + 1
  For y1 = 1 To 53
  
  Dim y2 As Integer
  y2 = y1 + 54
  
'  If x > MinRow Then
'   Sheets("結果").Cells(x, y1).Interior.Color = RGB(150, 150, 150)
'  ElseIf Sheets("結果").Cells(x, y1) <> Sheets("結果").Cells(x, y2) Then
  If Sheets("結果").Cells(x, y1) <> Sheets("結果").Cells(x, y2) Then
   Sheets("結果").Cells(x, y1).Interior.Color = RGB(255, 100, 100)
   Sheets("結果").Cells(x, y2).Interior.Color = RGB(255, 100, 100)
  End If
  Next
 Next
End Sub

Sub setVariable()
With Application.WorksheetFunction

 LastRow1 = 0
 LastRow2 = 0
 MaxRow = 0
 MinRow = 0
 LastLen1 = 0
 LastLen2 = 0
 MaxLen = 0
 MinLen = 0
 
 LastRow1 = Sheets("テキスト1").Cells(Rows().count, 1).End(xlUp).Row
 LastRow2 = Sheets("テキスト2").Cells(Rows().count, 1).End(xlUp).Row
 MaxRow = .Max(LastRow1, LastRow2)
 MinRow = .Min(LastRow1, LastRow2)
 
 For i = 1 To MaxRow
  LastLen1 = .Max(LastLen1, Len(Sheets("テキスト1").Cells(i, 1)))
  LastLen2 = .Max(LastLen2, Len(Sheets("テキスト2").Cells(i, 1)))
  MinLen = .Min(LastLen1, LastLen2)
 Next
 MaxLen = .Max(MaxLen, LastLen1, LastLen2)
 
 
 Dim x, count As Integer
 
 x = 0
 count = 0
 For i = 1 To LastRow1
  x = Len(Sheets("テキスト1").Cells(i, 1))
  If x > 53 Then
   count = count + Int(x / 53)
  End If
 Next
 LastRow1 = .Max(LastRow1, count + i - 1)
 
 x = 0
 count = 0
 For i = 1 To LastRow2
  x = Len(Sheets("テキスト2").Cells(i, 1))
  If x > 53 Then
   count = count + Int(x / 53)
  End If
 Next
 LastRow2 = .Max(LastRow2, count + i - 1)
 
 MaxRow = .Max(LastRow1, LastRow2)
 MinRow = .Min(LastRow1, LastRow2)
 resetResult
End With
End Sub

Sub resetText1()
 Sheets("テキスト1").Columns(1).Clear
 Sheets("テキスト1").DrawingObjects.Delete
 Sheets("結果").Range("A2:BA" & MaxRow + 2).Clear
End Sub

Sub resetText2()
 Sheets("テキスト2").Columns(1).Clear
 Sheets("テキスト2").DrawingObjects.Delete
 Sheets("結果").Range("BC2:DC" & MaxRow + 2).Clear
End Sub

Sub resetResult()
 Sheets("結果").Range("A2:A" & Rows().count).EntireRow.Clear
 
 Dim i As Integer
 i = MaxRow + 2
 
 Sheets("結果").Range(Cells(2, 54), Cells(i, 54)) = "■"
 Sheets("結果").Range(Cells(2, 108), Cells(i, 108)) = "■"
End Sub

Sub resetAll()
 resetText1
 resetText2
 setVariable
End Sub
