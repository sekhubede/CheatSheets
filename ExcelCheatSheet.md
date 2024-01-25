# Excel Cheat Sheet

[document descrition]

## Creating multiple worksheets from a list of cell values

```markdown
1. Hold down ALT + F11 keys to open the Microsoft Visual Basic for Applications window.
2. Click Insert > Module, and past the following code in the Module Window.
```

```vb
Sub AddSheets()
    Dim excelRange As Excel.Range
    Dim excelWorksheet As Excel.Worksheet
    Dim excelWorkbook As Excel.Workbook
    Set excelWorksheet = ActiveSheet
    Set excelWorkbook = ActiveWorkbook

    Application.ScreenUpdating = False

    For Each excelRange In excelWorksheet.Range("A1:A7")
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
<em>Note: In the aboce code, A1:A7 is the cell range that you want to create sheets based on, please change it to your need. </em>
```

```markdown
3. Press F5 key to run this code, and the new sheets have been created after all sheets in current workbook.
```
