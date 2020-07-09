# VBA-progress
The VBA code that I use most commonly

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
`Selection.Resize(Selection.Rows.Count, Selection.Columns.Count + 1).Select'

## Fill
`Selection.FillRight`
