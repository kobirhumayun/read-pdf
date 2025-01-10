Attribute VB_Name = "dictionary_utility_functions"
Option Explicit

Private Function CreateDicWithProvidedKeysAndValues( _
    ByVal keysArray As Variant, _
    ByVal valuesArray As Variant) As Object
    
    ' Ensure keysArray and valuesArray are arrays
    If Not IsArray(keysArray) Or Not IsArray(valuesArray) Then
        MsgBox "Both parameters must be arrays."
        Exit Function
    End If
    
    ' Check that both arrays have the same dimensions
    If (LBound(keysArray) <> LBound(valuesArray)) Or (UBound(keysArray) <> UBound(valuesArray)) Then
        MsgBox "Keys and values arrays have different sizes."
        Exit Function
    End If
    
    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    Dim tmpKey As String
    
    ' Loop through all elements
    For i = LBound(keysArray) To UBound(keysArray)
    
        tmpKey = Application.Run("general_utility_functions.RemoveInvalidChars", CStr(keysArray(i)))
        
        ' Check for duplicates
        If dic.Exists(tmpKey) Then
        
            MsgBox "Duplicate key found: '" & tmpKey & "'."
            Exit Function
            
        Else
        
            dic.Add key:=tmpKey, Item:=valuesArray(i)
            
        End If
        
    Next i
    
    ' Return the dictionary
    Set CreateDicWithProvidedKeysAndValues = dic
    
End Function

Private Function AddKeysWithPrimary(dictionary As Object, primaryKey As Variant, keysArray As Variant) As Object

    Dim removedAllInvalidChrFromKeys As Variant
    Dim i As Long

    ' Add new keys with primary value
    For i = LBound(keysArray) To UBound(keysArray)

        removedAllInvalidChrFromKeys = Application.Run("general_utility_functions.RemoveInvalidChars", keysArray(i))   'remove all invalid characters for use dic keys
        dictionary(removedAllInvalidChrFromKeys) = primaryKey

    Next i

    ' Return the modified dictionary
    Set AddKeysWithPrimary = dictionary
End Function

Private Function AddKeysAndValueSame(dictionary As Object, keysArray As Variant) As Object
        
    Dim removedAllInvalidChrFromKeys As Variant
        
    Dim i As Long
    ' Add same keys and value
    For i = LBound(keysArray) To UBound(keysArray)
        removedAllInvalidChrFromKeys = Application.Run("general_utility_functions.RemoveInvalidChars", keysArray(i))   'remove all invalid characters for use dic keys
        dictionary(removedAllInvalidChrFromKeys) = keysArray(i)
    Next i

    ' Return the modified dictionary
    Set AddKeysAndValueSame = dictionary
End Function

Private Function addKeysAndValueToDic(dictionary As Object, key As Variant, value As Variant) As Object

    Dim removedAllInvalidChrFromKeys As Variant

    ' Add key with values

        removedAllInvalidChrFromKeys = Application.Run("general_utility_functions.RemoveInvalidChars", key)   'remove all invalid characters for use dic keys
        
        If dictionary.Exists(removedAllInvalidChrFromKeys) Then
            MsgBox "Dictionary Key """ & removedAllInvalidChrFromKeys & """ Already Exists"
            Exit Function
        Else
            dictionary.Add removedAllInvalidChrFromKeys, value
        End If

    ' Return the modified dictionary
    Set addKeysAndValueToDic = dictionary

End Function
 
Private Function PutDictionaryValuesIntoWorksheet(wsRange As Range, dict As Object, keysPrint As Boolean, itemsPrint As Boolean, printOnColumn As Boolean)
    ' wsRange is just starting one cell address, the function dynamically resizes the range
    
    If dict.Count > 0 Then
    
        If (keysPrint And itemsPrint And printOnColumn) Then
    
            wsRange.Resize(dict.Count, 1).value = Application.Run("general_utility_functions.oneDArrayConvertToTwoDArray", dict.keys)
            wsRange.Offset(0, 1).Resize(dict.Count, 1).value = Application.Run("general_utility_functions.oneDArrayConvertToTwoDArray", dict.items)
            
        ElseIf (keysPrint And printOnColumn) Then
    
            wsRange.Resize(dict.Count, 1).value = Application.Run("general_utility_functions.oneDArrayConvertToTwoDArray", dict.keys)
    
        ElseIf (itemsPrint And printOnColumn) Then
    
            wsRange.Resize(dict.Count, 1).value = Application.Run("general_utility_functions.oneDArrayConvertToTwoDArray", dict.items)
    
        ElseIf (keysPrint And itemsPrint) Then
    
            wsRange.Resize(1, dict.Count).value = dict.keys
            wsRange.Offset(1, 0).Resize(1, dict.Count).value = dict.items
    
        ElseIf (keysPrint) Then
    
            wsRange.Resize(1, dict.Count).value = dict.keys
    
        ElseIf (itemsPrint) Then
    
            wsRange.Resize(1, dict.Count).value = dict.items
    
        End If
    
    End If

End Function

Private Function SortDictionaryByKey(dict As Object _
                  , Optional sortorder As XlSortOrder = xlAscending) As Object
    
    Dim arrList As Object
    Set arrList = CreateObject("System.Collections.ArrayList")
    
    ' Put keys in an ArrayList
    Dim key As Variant, coll As New Collection
    For Each key In dict
        arrList.Add key
    Next key
    
    ' Sort the keys
    arrList.Sort
    
    ' For descending order, reverse
    If sortorder = xlDescending Then
        arrList.Reverse
    End If
    
    ' Create new dictionary
    Dim dictNew As Object
    Set dictNew = CreateObject("Scripting.Dictionary")
    
    ' Read through the sorted keys and add to new dictionary
    For Each key In arrList
        dictNew.Add key, dict(key)
    Next key
    
    ' Clean up
    Set arrList = Nothing
    Set dict = Nothing
    
    ' Return the new dictionary
    Set SortDictionaryByKey = dictNew
        
End Function

Private Function mergeDict(mainDict As Object, addingDict As Object) As Object
    'this function received two dictionaries and merge them, then return merged dictionary

    Dim dictKey As Variant
    Dim i As Long
    For i = 0 To addingDict.Count - 1

        dictKey = addingDict.keys()(i)

        If mainDict.Exists(dictKey) Then
            MsgBox "Dictionary Key """ & dictKey & """ Already Exists"
            Exit Function
        Else
            mainDict.Add dictKey, addingDict(dictKey)
        End If

    Next i

    Set mergeDict = mainDict

End Function

Private Function sumOfProvidedKeys(dict As Object, arrOfKeys As Variant) As Variant
    'this function received a dictionary and a array of keys then
    ' sum of provided key's value and return the sum

    Dim element  As Variant
    Dim removedAllInvalidChrFromKeys As Variant
    Dim sum As Variant
    sum = 0

    For Each element In arrOfKeys

        removedAllInvalidChrFromKeys = Application.Run("general_utility_functions.RemoveInvalidChars", element)    'remove all invalid characters for use dic keys

        If dict.Exists(removedAllInvalidChrFromKeys) Then

            sum = sum + dict(removedAllInvalidChrFromKeys)

        Else

            MsgBox "Dictionary Key """ & removedAllInvalidChrFromKeys & """ Not Found"
            Exit Function

        End If

    Next

    sumOfProvidedKeys = sum

End Function

Private Function sumOfInnerDictOfProvidedKeys(dict As Object, arrOfKeys As Variant) As Variant
    'received a one level nested dictionary and a array of keys then
    ' sum of all inner dictionary of provided key's value and return the sum

    Dim sum As Variant
    sum = 0

    Dim dicKey As Variant

    For Each dicKey In dict.keys
        
        sum = sum + Application.Run("dictionary_utility_functions.sumOfProvidedKeys", dict(dicKey), arrOfKeys)

    Next dicKey

    sumOfInnerDictOfProvidedKeys = sum

End Function

Function ConvertDictToArrayOfDict(dict As Object) As Variant
    Dim dictArray() As Object
    Dim key As Variant
    Dim i As Long
    
    ' Initialize the array with the size of the dictionary
    ReDim dictArray(0 To dict.Count - 1)
    
    ' Iterate through each key in the dictionary
    i = 0
    For Each key In dict.keys
        ' Create a new dictionary for each key-value pair
        Dim newDict As Object
        Set newDict = CreateObject("Scripting.Dictionary")
        
        ' Add the key-value pair to the new dictionary
        newDict.Add key, dict(key)
        
        ' Assign the new dictionary to the array
        Set dictArray(i) = newDict
        i = i + 1
    Next key
    
    ' Return the array of dictionary objects
    ConvertDictToArrayOfDict = dictArray
End Function
