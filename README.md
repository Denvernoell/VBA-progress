# What is VBA
VBA (Visual Basic for Applications) is an available tool that can be used to automate repetitive tasks within programs like Microsoft Excel.

Here you can find what has been most useful to me as I have begun using it on a daily basis.

This is by no means all of the code that you will ever need but it is the code that I use most commonly.

## Quickstart.md

What I find most useful to start with is Movement

From there it might be helpful to learn if statements and loops which are in Logic

## Cell references
### Absolute
`Range("A1:C3").select`
### Relative
`Range("R1C1:R3C3").select`
### Cell number
`Range(Cells(1,1):Cells(3,3)).select`

## Movement
`activecell.offset(1,0).select`

## Selection
### To end
`Range(ActiveCell, ActiveCell.End(xlDown)).Select`
### Increase
`Selection.Resize(Selection.Rows.Count, Selection.Columns.Count + 1).Select`

## Fill
`Selection.FillRight`
