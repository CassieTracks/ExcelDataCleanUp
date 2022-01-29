Attribute VB_Name = "Module1"
Sub trimData()

    For Each c In Selection
        c.Value = Trim(c.Value)
        removeCharacters ("  ")
    Next

End Sub


Sub removeCharacters(removeText)

    For Each c In Selection
        Do While InStr(1, c.Value, removeText) <> 0
            c.Value = Replace(c.Value, removeText, "")
        Loop
    Next

End Sub
