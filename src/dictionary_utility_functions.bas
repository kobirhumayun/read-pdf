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
    
    Dim dic As New Scripting.dictionary
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