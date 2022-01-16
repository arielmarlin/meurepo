Dim fso, outFile
Set fso = CreateObject("Scripting.FileSystemObject")
Set outFile = fso.CreateTextFile("output.txt", True)

set json = CreateObject("Chilkat_9_5_0.JsonObject")

'  This is the above JSON with whitespace chars removed (SPACE, TAB, CR, and LF chars).
'  The presence of whitespace chars for pretty-printing makes no difference to the Load
'  method.
jsonStr = "{ ""numbers"" : [ [""even"", 2, 4, 6, 8], [""prime"", 2, 3, 5, 7, 11, 13] ] }"

success = json.Load(jsonStr)
If (success <> 1) Then
    outFile.WriteLine(json.LastErrorText)
    WScript.Quit
End If

'  Get the value of the "numbers" object, which is an array that contains JSON arrays.
' outerArray is a Chilkat_9_5_0.JsonArray
Set outerArray = json.ArrayOf("numbers")
If (outerArray Is Nothing ) Then
    outFile.WriteLine("numbers array not found.")
    WScript.Quit
End If

numArrays = outerArray.Size

For i = 0 To numArrays - 1

    ' innerArray is a Chilkat_9_5_0.JsonArray
    Set innerArray = outerArray.ArrayAt(i)

    '  The first item in the innerArray is a string
    outFile.WriteLine(innerArray.StringAt(0) & ":")

    numInnerItems = innerArray.Size

    For j = 1 To numInnerItems - 1

        outFile.WriteLine("  " & innerArray.IntAt(j))

    Next

Next


outFile.Close
