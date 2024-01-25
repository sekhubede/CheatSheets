# Excel Cheat Sheet

[document descrition]

## Add Worksheets to a Workbook from a list of cell values in the same Workbook

```markdown
1. Hold down ALT + F11 keys to open the Microsoft Visual Basic for Applications window.
2. Click Insert > Module, and paste the following code in the Module Window.
```

```vb
Sub AddSheets()
    Dim excelRange As Excel.Range
    Dim excelWorksheet As Excel.Worksheet
    Dim excelWorkbook As Excel.Workbook
    Set excelWorksheet = ActiveSheet
    Set excelWorkbook = ActiveWorkbook

    Application.ScreenUpdating = False

    For Each excelRange In excelWorksheet.Range("A1:A5")
        With excelWorkbook
            .Sheets.Add after:=.Sheets(.Sheets.Count)

            On Error Resume Next

            ActiveSheet.Name = excelRange.Value

            If Err.Number = 1004 Then
              Debug.Print excelRange.Value & " already used as a sheet name"
            End If

            On Error GoTo 0
        End With
    Next excelRange

    Application.ScreenUpdating = True

End Sub
```

```markdown
# Note: In the above code, A1:A5 is the cell range that you want to create sheets based on, please change it to your need.
```

```markdown
3. Press F5 key to run this code, and the new sheets will be created after all sheets in the current workbook.
```

## Add Worksheet names to cells in a Workbook

```markdown
1. Click Formulas, under the Defined Names tab click Name Manager, then click the New button.
2. Type ListSheets in the Name field of the pop-up window.
3. Add the below formula to the Refers to field, then click OK.
```

```vb
=REPLACE(GET.WORKBOOK(1),1,FIND("]",GET.WORKBOOK(1)),"")
```

```markdown
4. In the first cell you want the Worksheet names to start, paste the following formaula.
```

```vb
=INDEX(ListSheets,ROW(A2))
```

```markdown
## Note: In the above formula, A2 referes to the second Worksheet, assuming that the first Worksheet is the dashboard that will have the names of the Worksheets added to the Workbook.
```
