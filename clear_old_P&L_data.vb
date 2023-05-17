Sub OPEN_MONSTER()

'it opens a .xlsx file and clears all prevous records. Useful for repetitive reporting tasks

Dim wb As Workbook
    Dim ch_monster As String

    ch_monster = "your_path"
    '~~> open the workbook and pass it to workbook object variable
    Set wb = Workbooks.Open(ch_monster)
End Sub

Sub CLEAR_DATA_FIELD()

'Jan-Dez Bereich

    Range("C6:N41").Select
    Selection.ClearContents
    
'YTD 21 Bereich

Range("P6:P41").Select
    Selection.ClearContents
    
'YTD vs. Budget Bereich

    Range("S6:S41").Select
    Selection.ClearContents
    
End Sub

Sub close_monster()

Dim wb As Workbook
    Dim ch_monster As String

    ch_monster = "your_path"
    '~~> open the workbook and pass it to workbook object variable
    Set wb = Workbooks.Close(ch_monster)
End Sub
