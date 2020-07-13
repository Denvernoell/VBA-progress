# Files
## Check if file exists
```
Num = "1st"
Filename = (path & "\OldPdfs\" & Num & Name & ".pdf")
    check = Dir(Filename)
    If check = "" Then
    GoTo Process
    End If
```
