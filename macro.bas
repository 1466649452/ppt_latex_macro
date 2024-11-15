Sub PasteLatex()
  Application.CommandBars.ExecuteMso ("EquationInsertNew")
  Dim originalLatex As String
  Dim modifiedLatex As String
  
  originalLatex = InputBox("Equation")
  modifiedLatex = Replace(originalLatex, "\mathrm", "\text")
  
  With ActiveWindow.Selection.ShapeRange.TextFrame.TextRange
    .Characters(.Length - 1) = modifiedLatex
  End With
  Application.CommandBars.ExecuteMso ("EquationProfessional")
End Sub

Sub SwitchLatex()
  Application.CommandBars.ExecuteMso ("EquationInsertNew")
  With ActiveWindow.Selection.ShapeRange.TextFrame.TextRange
    .Characters(.Length - 1) = ChrW(&H24C9)
  End With
End Sub