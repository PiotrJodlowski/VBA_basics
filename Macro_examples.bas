Attribute VB_Name = "Module1"

Sub Wstaw()
 ActiveCell = Now
 Selection.Columns.AutoFit
End Sub

Sub Dodaj()
 Range("A11").Value = Range("A9").Value + Range("A10").Value
End Sub


Sub Zmien_nazwe()
 If (Range("A17").Value) <> "" Then
    ActiveSheet.Name = Range("A17").Value
    End If
End Sub


Sub Oblicz()
  Range("A32").Value = Switch(Range("A28").Value = 1, Range("A30").Value + Range("A31").Value, Range("A28").Value = 2, Range("A30").Value - Range("A31").Value, Range("A28").Value = 3, Range("A30").Value * Range("A31").Value)
End Sub
