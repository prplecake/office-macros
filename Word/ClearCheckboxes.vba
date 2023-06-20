' https://stackoverflow.com/questions/51644991/clear-all-checkboxes-in-word-doc/51645189#51645189
Sub ClearCheckboxes()
'
' ClearCheckboxes Macro
'
  Dim ctrl As Word.ContentControl
  For Each ctrl In ActiveDocument.ContentControls
    If ctrl.Type = wdContentControlCheckBox Then
      ctrl.Checked = False
    End If
  Next ctrl
End Sub
