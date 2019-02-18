'This macro merges cells with the same value

Sub MergeColumnA()
    Dim i As Long
    Dim myLastRow As Long
        Application.DisplayAlerts = False
        myLastRow = Cells(Rows.Count, "A").End(xlUp).Row
            For i = myLastRow - 1 To 1 Step -1
                If Cells(i + 1, 1) <> "" Then If Cells(i, 1) = Cells(i + 1, 1) Then Range(Cells(i, 1), Cells(i + 1, 1)).Merge
                Next
            End
'                Next i Application.DisplayAlerts = True
End Sub
