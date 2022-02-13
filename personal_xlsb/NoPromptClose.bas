Attribute VB_Name = "NoPromptClose"
Sub closeWithoutPrompt()
Attribute closeWithoutPrompt.VB_Description = "Closes the current workbook without a save prompt, discarding changes"
Attribute closeWithoutPrompt.VB_ProcData.VB_Invoke_Func = "W\n14"
'
' NoPromptClose Macro
' Closes the current workbook without a save prompt, discarding changes
'
' Touche de raccourci du clavier: Ctrl+Shift+W
'

    'Set the save flag of the current workbook to disable the save prompt and close
    ActiveWorkbook.Saved = True
    ActiveWorkbook.Close
    
    'Exit Excel if all workbooks are closed (Personal.xlsb is always open by default)
    If Application.Workbooks.Count = 1 Then
        Application.Quit
    End If

End Sub
