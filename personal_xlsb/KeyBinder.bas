Attribute VB_Name = "KeyBinder"
Sub setShortcuts()

Application.OnKey "^+F", "clearFilters"
Application.OnKey "^+V", "pasteValues"
Application.OnKey "^+W", "closeWithoutPrompt"
Application.OnKey "^m", "setFilterOnCurrentColumn"

End Sub
