Sub ConvertToiCal()
    Dim i As Long, v As Long
    Dim RowStart As Long
    Dim objFSO, objFile
    Dim FilePath, FileName, TxtSummary, TxtDescr, TxtLoc, DateBorn, DateDied As String
    
    RowStart = 4    ' where to find first data
    FileName = "output"
    If ActiveSheet.Cells(1, 2).Value <> "" Then
        FileName = ActiveSheet.Cells(1, 2).Value
    End If
    FilePath = "C:\Users\Public\Documents\" & FileName & ".ics"
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFSO.CreateTextFile(FilePath)
 
 ' column 1=Name, 2=Description, 3=Date born, 4=Date died
    objFile.write "BEGIN:VCALENDAR" & vbCrLf & "CALSCALE:GREGORIAN" & vbCrLf
    i = RowStart
    While ActiveSheet.Cells(i, 1).Value <> ""           ' No more names? Then stop
        TxtSummary = "Geburtstag " & ActiveSheet.Cells(i, 1)    ' Name
        TxtDescr = ""
        If ActiveSheet.Cells(i, 8) <> "" Then
            TxtDescr = TxtDescr & "Partner: " & ActiveSheet.Cells(i, 8) & "\n"  ' Partner
        End If
        If ActiveSheet.Cells(i, 9) <> "" Then
            TxtDescr = TxtDescr & "Eltern: " & ActiveSheet.Cells(i, 9) & "\n"   ' Parents
        End If
        If ActiveSheet.Cells(i, 10) <> "" Then
            TxtDescr = TxtDescr & "Kinder: " & ActiveSheet.Cells(i, 10) & "\n"  ' Childs
        End If
        TxtDescr = TxtDescr & ActiveSheet.Cells(i, 2)           ' add orig. Description
        TxtLoc = ActiveSheet.Cells(i, 6)                        ' Location
        
        DateBorn = Format(ActiveSheet.Cells(i, 3), "yyyymmdd")  ' Date born
        objFile.write "BEGIN:VEVENT" & vbCrLf & _
                        "DTSTART;VALUE=DATE:" & DateBorn & vbCrLf & _
                        "DTEND;VALUE=DATE:" & DateBorn & vbCrLf & _
                        "RRULE:FREQ=YEARLY" & vbCrLf & _
                        "SUMMARY:" & TxtSummary & vbCrLf & _
                        "DESCRIPTION:" & TxtDescr & vbCrLf & _
                        "LOCATION:" & TxtLoc & vbCrLf & _
                        "END:VEVENT" & vbCrLf
                        
        DateDied = ActiveSheet.Cells(i, 4).Value                ' Date died available?
        If DateDied <> "" Then
            DateDied = Format(ActiveSheet.Cells(i, 4), "yyyymmdd")
            TxtSummary = "Todestag " & ActiveSheet.Cells(i, 1)
            objFile.write "BEGIN:VEVENT" & vbCrLf & _
                            "DTSTART;VALUE=DATE:" & DateDied & vbCrLf & _
                            "DTEND;VALUE=DATE:" & DateDied & vbCrLf & _
                            "RRULE:FREQ=YEARLY" & vbCrLf & _
                            "SUMMARY:" & TxtSummary & vbCrLf & _
                            "DESCRIPTION:" & TxtDescr & vbCrLf & _
                            "LOCATION:" & TxtLoc & vbCrLf & _
                            "END:VEVENT" & vbCrLf
        End If
        
        i = i + 1
    Wend
    objFile.write "END:VCALENDAR"
    MsgBox ("Ready: " & FilePath)
End Sub