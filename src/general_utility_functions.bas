Attribute VB_Name = "general_utility_functions"
Option Explicit

Private Function InsertStringAtPosition(originalString As String, insertString As String, position As Integer) As Variant
    Dim length As Integer
    length = Len(originalString)
    
    If length >= position Then
        InsertStringAtPosition = Left(originalString, length - (position - 1)) & insertString & Right(originalString, (position - 1))
    Else
        InsertStringAtPosition = Null
    End If
End Function

Private Function RemoveInvalidChars(ByVal inputString As String) As String
    
    Static mapInitialized As Boolean
    Static invalidMap(0 To 255) As Boolean
    
    Dim invalidChars As String
    invalidChars = " ~`!@#$%^&*()-+=[]\{}|;':"",./<>?" & vbNewLine & Chr(10) & Chr(160)
    
    Dim i As Long
    Dim c As String
    Dim resultString As String
    
    ' Initialize the map once
    If Not mapInitialized Then
        For i = 1 To Len(invalidChars)
            invalidMap(Asc(Mid$(invalidChars, i, 1))) = True
        Next i
        mapInitialized = True
    End If
    
    ' Build the result by skipping invalid chars
    For i = 1 To Len(inputString)
        c = Mid$(inputString, i, 1)
        ' Only append if character is not marked invalid
        If Not invalidMap(Asc(c)) Then
            resultString = resultString & c
        End If
    Next i
    
    RemoveInvalidChars = resultString
End Function

Private Function oneDArrayConvertToTwoDArray(inputArray As Variant) As Variant
    Dim outputArray As Variant

    ReDim outputArray(LBound(inputArray) To UBound(inputArray), 1 To 1)

    Dim i As Long
    For i = LBound(inputArray) To UBound(inputArray)
        outputArray(i, 1) = inputArray(i)
    Next i

    oneDArrayConvertToTwoDArray = outputArray

End Function

Private Function regExReturnedObj(str As Variant, pattern As Variant, isGlobal As Boolean, isIgnoreCase As Boolean, isMultiLine As Boolean) As Object

    Dim regex As Object

    ' Convert the str to a string
    str = CStr(str)

    ' Convert the pattern to a string
    pattern = CStr(pattern)

    ' Create a RegExp object
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = isGlobal
        .IgnoreCase = isIgnoreCase
        .MultiLine = isMultiLine
        .pattern = pattern
    End With

    ' Return the test result
    Set regExReturnedObj = regex.Execute(str)

End Function

Private Function createRegExObj(pattern As String, isGlobal As Boolean, isIgnoreCase As Boolean, isMultiLine As Boolean) As Object

    Dim regEx As Object

    ' Create a RegExp object
    Set regEx = CreateObject("VBScript.RegExp")
    
    With regEx
        .Global = isGlobal
        .IgnoreCase = isIgnoreCase
        .MultiLine = isMultiLine
        .pattern = pattern
    End With

    ' Return regEx object
    Set createRegExObj = regEx

End Function

Private Function isStrPatternExist(str As Variant, pattern As Variant, isGlobal As Boolean, isIgnoreCase As Boolean, isMultiLine As Boolean) As Boolean

    Dim regex As Object

    ' Convert the str to a string
    str = CStr(str)

    ' Convert the pattern to a string
    pattern = CStr(pattern)

    ' Create a RegExp object
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = isGlobal
        .IgnoreCase = isIgnoreCase
        .MultiLine = isMultiLine
        .pattern = pattern
    End With

    ' Return the test result
    isStrPatternExist = regex.test(str)

End Function

Private Function ExtractLeftDigitWithRegex(number As Variant) As Variant
    Dim regex As Object
    Dim matches As Object
    Dim pattern As String
    Dim leftDigit As Variant
    
    ' Convert the number to a string
    Dim numberString As String
    numberString = CStr(number)
    
    ' Define the regular expression pattern to match the left digits
    pattern = "\d+"
    
    ' Create a RegExp object
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = True
        .IgnoreCase = True
        .MultiLine = False
        .pattern = pattern
    End With
    
    ' Get the matches
    Set matches = regex.Execute(numberString)
    
    ' Check if there's a match
    If matches.Count > 0 Then
        
        leftDigit = matches(0)
        
    Else
        ' Default to 0 if no match found
        leftDigit = 0
    End If
    
    ' Return the extracted left digit
    ExtractLeftDigitWithRegex = leftDigit
    
End Function

Private Function ExtractFirstLineWithRegex(str As Variant) As Variant

    Dim matches As Object
    Dim firstLine As Variant
    
    ' Get the matches
    Set matches = Application.Run("general_utility_functions.regExReturnedObj", str, ".+", True, True, True) ' extract first line
    
    ' Check if there's a match
    If matches.Count > 0 Then
        
        firstLine = matches(0)
        
    Else
        ' Default to 0 if no match found
        firstLine = 0
    End If
    
    ' Return the extracted first line
    ExtractFirstLineWithRegex = firstLine
    
End Function

Private Function ExtractRightDigitFromEnd(str As Variant) As Variant
    Dim regex As Object
    Dim matches As Object
    Dim pattern As String
    Dim rightDigit As Variant

    ' Convert the str to a string
    Dim numberString As String
    numberString = CStr(str)

    ' Define the regular expression pattern to match the right digits
    pattern = "\d+$"

    ' Create a RegExp object
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = True
        .IgnoreCase = True
        .MultiLine = True
        .pattern = pattern
    End With

    ' Get the matches
    Set matches = regex.Execute(numberString)

    ' Check if there's a match
    If matches.Count > 0 Then

        rightDigit = matches(0)

    Else
        ' Default to 0 if no match found
        rightDigit = 0
    End If

    ' Return the extracted right digit
    ExtractRightDigitFromEnd = rightDigit

End Function

Private Function dictKeyGeneratorWithProvidedArrayElements(ByVal arr As Variant) As String

    Dim tempDictKeyStr As String
    Dim elements As Variant

    tempDictKeyStr = ""

    For Each elements In arr
       tempDictKeyStr = tempDictKeyStr & "_" & elements
    Next elements

    tempDictKeyStr = Right(tempDictKeyStr, Len(tempDictKeyStr) - 1)

    dictKeyGeneratorWithProvidedArrayElements = tempDictKeyStr
    
End Function
  
Private Function CopyFileToFolderUsingFSO(sourceFilePath As String, targetFolderPath As String, overwrite As Boolean)

    On Error Resume Next

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(sourceFilePath) Then
        Dim fileName As String
        fileName = fso.GetFileName(sourceFilePath)

        Dim targetPath As String
        targetPath = fso.BuildPath(targetFolderPath, fileName)

        fso.CopyFile sourceFilePath, targetPath, overwrite

        ' Check if the copy was successful
        If Err.number = 0 Then
            MsgBox "File " & sourceFilePath & " copied successfully!"
        Else
            MsgBox "Target " & targetFolderPath & " " & Err.Description
        End If
    Else
        MsgBox "Source file " & sourceFilePath & " not found."
    End If

End Function

Private Function CopyFileAsNewFileFSO(sourceFilePath As String, newFilePath As String, overwrite As Boolean)

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(sourceFilePath) Then
        
        fso.CopyFile sourceFilePath, newFilePath, overwrite

    Else
        MsgBox "Source file " & sourceFilePath & " not found."
    End If

End Function


Private Function sequentiallyRelateTwoArraysAsDictionary(properties_1 As String, properties_2 As String, properties_1_Arr As Variant, properties_2_Arr As Variant) As Variant
    'this function take two str as properties & two arr contain values then return a dictionary, dictionary use first arr elements as keys
    'and all keys are also dictionaries that contain same sequential elements of two arr

    Dim mainDictionary As Object
    Set mainDictionary = CreateObject("Scripting.Dictionary")

    Dim subDictionary As Object
    Dim removedAllInvalidChrFromKeys As String
    
    If LBound(properties_1_Arr) <> LBound(properties_2_Arr) Or UBound(properties_1_Arr) <> UBound(properties_2_Arr) Then
    
        MsgBox "Both array length are not same"
        Exit Function
        
    End If
    
    Dim i As Long

    ' create sub dictionary and add to main dictionary
    For i = LBound(properties_1_Arr) To UBound(properties_1_Arr)

        Set subDictionary = Application.Run("dictionary_utility_functions.CreateDicWithProvidedKeysAndValues", Array(properties_1, properties_2), Array(properties_1_Arr(i), properties_2_Arr(i)))

        removedAllInvalidChrFromKeys = Application.Run("general_utility_functions.RemoveInvalidChars", properties_1_Arr(i))   'remove all invalid characters for use dic keys
        mainDictionary.Add removedAllInvalidChrFromKeys, subDictionary
    Next i

    Set sequentiallyRelateTwoArraysAsDictionary = mainDictionary

End Function

Private Function ExcludeElements(arr1 As Variant, arr2 As Variant) As Variant
    'exclude all the elements from first array which elements exist in second array
    Dim i As Long
    Dim j As Long

    Dim arr2Dictionary As Object
    Set arr2Dictionary = CreateObject("Scripting.Dictionary")

    Dim excludedDictionary As Object
    Set excludedDictionary = CreateObject("Scripting.Dictionary")
        
    ' Loop through the elements of arr2
    For i = LBound(arr2) To UBound(arr2)
        
        arr2Dictionary(arr2(i)) = arr2(i)
        
    Next i

    ' Loop through the elements of arr1
    For j = LBound(arr1) To UBound(arr1)

        If Not arr2Dictionary.Exists(arr1(j)) Then
            excludedDictionary(arr1(j)) = arr1(j)
        End If
        
    Next j
        
    ' Return the result array
    ExcludeElements = excludedDictionary.keys
        
End Function
