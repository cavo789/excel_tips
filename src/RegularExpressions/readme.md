# Using RegEx in Excel VBA

> Simple demo code on how take advantages of the power of regex

Use late bindings so the user shouldn't reference a library.

The following function will keep only digits in a string so the code here below will show `12345678`:

```vbnet
Msgbox KeepDigitsOnly("123456 - TEST - 7-8")
```

The function:

```vbnet
Public Function KeepDigitsOnly(ByVal sValue As String) As String

Dim sReturn As String
Dim objRegex As Object

    sReturn = ""

    On Error Resume Next

    Set objRegex = CreateObject("vbscript.regexp")

    With objRegex
        .MultiLine = False
        .Global = True
        .IgnoreCase = False
        ' Match everything except a digit
        .Pattern = "[^\d+]"
    End With

    ' Replace matched characters by nothing
    sReturn = objRegex.Replace(sValue, vbNullString)

    Set objRegex = Nothing

    If Err.Number <> 0 Then
        sReturn = ""
        Err.Clear
    End If

    KeepDigitsOnly = sReturn

End Function
```
