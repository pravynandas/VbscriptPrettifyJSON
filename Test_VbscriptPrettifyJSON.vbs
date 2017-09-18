includeRelFile "./VbscriptPrettifyJSON.vbs"

data = createobject("Scripting.FileSystemObject").OpenTextFile("data.json", 1).ReadAll()
set oOut = createobject("Scripting.FileSystemObject").OpenTextFile("data_out.json", 2, True)
oOut.Writeline( VbscriptPrettifyJSON(data) )
oOut.close

Sub includeRelFile(fSpec)
    executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile(fSpec).readAll()
End Sub