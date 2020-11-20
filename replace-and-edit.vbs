Set objFS = CreateObject("Scripting.FileSystemObject")
strFile = "input.nc"
Set objFile = objFS.OpenTextFile(strFile)
Do Until objFile.AtEndOfStream
    strLine = objFile.ReadLine
    If InStr(strLine,"TOOL CALL 0     Z S1000")> 0 Then
        strLine = Replace(strLine,"TOOL CALL 0     Z S1000","TOOL CALL 0     Z S1600")
    Else 
        words = Split(strLine)
        for each word in words
            If word = "M4" Then
                strLine = Replace(strLine,"M4","M3")
            ElseIf word = "M8" Then
                strLine = Replace(strLine,"M8","M7")
            End If
        next
    End If 
    WScript.Echo strLine
Loop    

