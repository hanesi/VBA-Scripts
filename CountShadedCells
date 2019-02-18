'This Macro counts the number of shaded cells in a defined range
'and outputs the value to a specific cell

Sub SumCountByConditionalFormat()
    Dim cellrngi, cellrngj, cellrngk, cellrngl As Range
    Dim cntresi, cntresj, cntresk, cntresl As Long
 
    cntresi = 0
    cntresj = 0
    cntresk = 0
    cntresl = 0
    
    Set cellrngi = Sheets("Sheet3").Range("I2:I81")
    Set cellrngj = Sheets("Sheet3").Range("J2:J81")
    Set cellrngk = Sheets("Sheet3").Range("K2:K81")
    Set cellrngl = Sheets("Sheet3").Range("L2:L81")
 
    For Each i In cellrngi
        If i.DisplayFormat.Interior.Color <> 16777215 Then
        cntresi = cntresi + 1
        End If
    Next i

    For Each j In cellrngj
        If j.DisplayFormat.Interior.Color <> 16777215 Then
        cntresj = cntresj + 1
        End If
    Next j
    
    For Each k In cellrngk
        If k.DisplayFormat.Interior.Color <> 16777215 Then
        cntresk = cntresk + 1
        End If
    Next k
    
    For Each l In cellrngl
        If l.DisplayFormat.Interior.Color <> 16777215 Then
        cntresl = cntresl + 1
        End If
    Next l
    
    MsgBox (cntresi)
    MsgBox (cntresj)
    MsgBox (cntresk)
    MsgBox (cntresl)
    
    Range("A89").Value = cntresi
    Range("A90").Value = cntresj
    Range("A88").Value = cntresk
    Range("A86").Value = cntresl
    Range("A87").Value = 90 - cntresl
    

End Sub
