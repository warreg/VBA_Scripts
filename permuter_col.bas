Sub Rearrange()
   With ActiveSheet.UsedRange
       .Value = Application.Index(.Value, .Worksheet.Evaluate("ROW(" & .Columns(1).Address & ")"), Array(7, 2, 4, 1, 6, 5, 3))
   End With
End Sub
