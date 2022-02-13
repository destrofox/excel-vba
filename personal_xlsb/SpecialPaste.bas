Attribute VB_Name = "SpecialPaste"
Sub pasteValues()
Attribute pasteValues.VB_ProcData.VB_Invoke_Func = " \n14"

'
' pasteValues Macro
' Paste the clipboard content as values, discarding any formulas and/or formatting
'
' Touche de raccourci du clavier: Ctrl+Shift+V
'

On Error Resume Next
ActiveCell.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
End Sub
