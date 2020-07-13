# Formatting

```
Sub rowheight()
ScreenUpdating = False
    Dim hgt As Variant
    Dim WorkRng As Range
    xTxt = ActiveWindow.RangeSelection.Address
    Set WorkRng = Application.InputBox("please select the data range:", "Kutools for Excel", xTxt, , , , , 8)
    'Set WorkRng = Range("A2")
    
For Each H In WorkRng
        If H.Value <> "" Then
            hgt = H.Value
            H.EntireRow.Select
            Selection.rowheight = hgt
        End If
    Next H
ScreenUpdating = True
End Sub
```

```
Sub rowheightdisplay()
ScreenUpdating = False
    Dim hgt As Variant
    Dim WorkRng As Range
    Dim i As Integer
    xTxt = ActiveWindow.RangeSelection.Address
    Set WorkRng = Application.InputBox("please select the data range:", "Kutools for Excel", xTxt, , , , , 8)
For Each H In WorkRng
        If H.rowheight <> "" Then
            hgt = H.rowheight
            H.Select
            Selection.Value = hgt
        End If
    Next H
ScreenUpdating = True
End Sub
```
```
Sub columnwidth()
ScreenUpdating = False
    Dim hgt As Variant
    Dim WorkRng As Range
    xTxt = ActiveWindow.RangeSelection.Address
    Set WorkRng = Application.InputBox("please select the data range:", "Kutools for Excel", xTxt, , , , , 8)
For Each C In WorkRng
        If C.Value <> "" Then
            W = C.Value
            C.EntireColumn.Select
            Selection.columnwidth = W
        End If
    Next C
ScreenUpdating = True
End Sub
```
```
Sub columnwidthdisplay()
ScreenUpdating = False

    Dim hgt As Variant
    Dim WorkRng As Range
    Dim i As Integer
    Dim iColumnWidth As Long
    xTxt = ActiveWindow.RangeSelection.Address
    Set WorkRng = Application.InputBox("please select the data range:", "Kutools for Excel", xTxt, , , , , 8)
For Each C In WorkRng
        If C.columnwidth <> "" Then
        'iColumnWidth = columns("a").ColumnWidth
            W = C.columnwidth
            C.Select
            Selection.Value = W
        End If
    Next C
    ScreenUpdating = True
End Sub
```
```
Sub ShapeDisplay()


W = Selection.ShapeRange.Width
H = Selection.ShapeRange.Height
MsgBox "Shape width is  " & W & vbNewLine & "Shape Height is  " & H


End Sub
```
```
Sub getRGB1()
    col = Selection.Font.Color
    MsgBox (col)
End Sub
```
```
Sub colored()

'Selection.Font.Color = 3421846

'MsgBox (RGB(255, 255, 204))
Selection.Interior.Color = RGB(255, 255, 204)

End Sub
```
```
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function            Color
'   Purpose             Determine the Background Color Of a Cell
'   @Param rng          Range to Determine Background Color of
'   @Param formatType   Default Value = 0
'                       0   Integer
'                       1   Hex
'                       2   RGB
'                       3   Excel Color Index
'   Usage               Color(A1)      -->   9507341
'                       Color(A1, 0)   -->   9507341
'                       Color(A1, 1)   -->   91120D
'                       Color(A1, 2)   -->   13, 18, 145
'                       Color(A1, 3)   -->   6
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function Color(rng As Range, Optional formatType As Integer = 0) As Variant
    Dim colorVal As Variant
    colorVal = Cells(rng.Row, rng.Column).Interior.Color
    Select Case formatType
        Case 1
            Color = Hex(colorVal)
        Case 2
            Color = (colorVal Mod 256) & ", " & ((colorVal \ 256) Mod 256) & ", " & (colorVal \ 65536)
        Case 3
            Color = Cells(rng.Row, rng.Column).Interior.ColorIndex
        Case Else
            Color = colorVal
    End Select
End Function
```
```
Sub FindColor()

'MsgBox (Color(activecell.Interior.color), 2))
MsgBox (Color(Range("F77"), 2))
MsgBox (Color(Range("J47"), 2))


End Sub
```

```
Function AkelColor(ColorType As String)

Select Case ColorType

Case "Yellow1"
AkelColor = RGB(255, 255, 204)

Case "Orange1"
AkelColor = RGB(253, 233, 217)

Case "Blue1"
AkelColor = RGB(79, 129, 189)

Case "Red1"
AkelColor = RGB(192, 80, 77)

Case "Grey1"
AkelColor = RGB(217, 217, 217)

Case "Grey2"
AkelColor = RGB(128, 128, 128)

End Select

End Function
```
```
Sub UseAkel()


Selection.Interior.Color = AkelColor("Blue1")



End Sub
```
