Sub textreplace()
    Dim c As Range
            For Each c In Range("A2:L41")               'Define range
            If c.Value = "Unknown" Then c.Value = "---" '1st value is what you have, 2nd is what you want
            If c.Value = "N/A" Then c.Value = "---"     'Same as above
            Next
End Sub

'If you want to add text to the cell, then c.Value = "What to Add" & c.Value
