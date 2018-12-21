Sub TestArray()
    'Vars:
    Dim TestArray As Variant, part1 As String, part2 As String, part3 As String, part4 As String, part5 As String
    
    'Parts:
    part1 = "<div class=""single-review"">" & vbCrLf & vbTab & "<div class=""review-content"">" & vbCrLf & vbTab & vbTab & "<div class=""stars-holder"">" & vbCrLf & vbTab & vbTab & vbTab & "<div class=""stars-image"">" & vbCrLf & vbTab & vbTab & vbTab & vbTab & "<div class=""stars-active"" style=""width: 100%;"">"
    part2 = "</div>" & vbCrLf & vbTab & vbTab & vbTab & "</div>" & vbCrLf & vbTab & vbTab & "</div>" & vbCrLf & vbTab & vbTab & "<div class=""reviewer"">" & vbCrLf & vbTab & vbTab & vbTab & "<div class=""reviewer-name"">"
    part3 = "</div>" & vbCrLf & vbTab & vbTab & vbTab & "<div class=""review-date"">"
    part4 = "</div>" & vbCrLf & vbTab & vbTab & vbTab & "<div class=""review-description"">"
    part5 = "</div>" & vbCrLf & vbTab & vbTab & "</div>" & vbCrLf & vbTab & "</div>" & vbCrLf & "</div>"
    
    'Array:
    TestArray = Range("A2:E15").Value
    
    'Debug:
    'MsgBox TestArray(5, 1)
    
    'Row Format:
    TestRow = part1 & TestArray(1, 1) & part2 & TestArray(1, 2) & " " & TestArray(1, 3) & part3 & TestArray(1, 4) & part4 & TestArray(1, 5) & part5
    MsgBox TestRow
    
    'Cycle
    For i = 1 To 5
        For j = 1 To 15
            
        Next j
    Next i
End Sub
