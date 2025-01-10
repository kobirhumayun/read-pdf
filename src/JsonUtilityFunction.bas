Attribute VB_Name = "JsonUtilityFunction"

Private Function SaveDictionaryToJsonTextFile(dict As Object, filePath As String)

    ' Convert the dictionary to JSON
    Dim json As String
    json = JsonConverter.ConvertToJson(dict)
    
    ' Write the JSON to a file
    Open filePath For Output As #1
    Print #1, json
    Close #1

    Debug.Print "Dictionary save as Json"

End Function

Private Function SaveArrayOfDictionaryToJsonTextFile(arrayOfDict As Variant, filePath As String)

    ' Convert the array of dictionary to JSON
    Dim json As String
    json = JsonConverter.ConvertToJson(arrayOfDict)
    
    ' Write the JSON to a file
    Open filePath For Output As #1
    Print #1, json
    Close #1

    Debug.Print "Array of dictionary save as Json"

End Function

Private Function LoadDictionaryFromJsonTextFile(filePath As String) As Object

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim json As String

    ' Read the JSON from the file
    Open filePath For Input As #1
    json = Input(LOF(1), #1)
    Close #1
    
    ' Convert JSON to dictionary
    Set dict = JsonConverter.ParseJson(json)
    
    Set LoadDictionaryFromJsonTextFile = dict

End Function


