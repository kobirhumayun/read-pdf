Attribute VB_Name = "Brac"
Option Explicit

Private Function ExtractPdfLcBrac(lcPath As String) As Object
    
    Dim lcText As String
    lcText = Application.Run("readPdf.ExtractTextFromPdfUsingAcrobatAcroHiliteList", lcPath)
    
    Dim resultDict As Object
    Set resultDict = CreateObject("Scripting.Dictionary")
    
    resultDict.Add "lcNo", Application.Run("Brac.ExtractLcNoBrac", lcText)
    resultDict.Add "lcDt", Application.Run("Brac.ExtractLcDtBrac", lcText)
    
    Set ExtractPdfLcBrac = resultDict
    
End Function

Private Function ExtractLcNoBrac(lcText As String) As String
    Dim lcPortionObj As Object
    Dim lcNo As String

    ' First regex match to extract the portion containing LC number
    Set lcPortionObj = Application.Run("general_utility_functions.regExReturnedObj", lcText, "20.+\n.+\n31c", True, True, True)
    If lcPortionObj Is Nothing Or lcPortionObj.Count = 0 Then
        ExtractLcNoBrac = vbNullString
        Exit Function
    End If

    ' Second regex match to extract the LC number from the portion
    Set lcPortionObj = Application.Run("general_utility_functions.regExReturnedObj", lcPortionObj(0), ".+", True, True, True)
    If lcPortionObj Is Nothing Or lcPortionObj.Count < 2 Then
        ExtractLcNoBrac = vbNullString
        Exit Function
    End If

    ' Extract the LC number
    lcNo = lcPortionObj(1)
    ExtractLcNoBrac = lcNo
End Function

Private Function ExtractLcDtBrac(lcText As String) As String
    Dim lcDtPortionObj As Object
    Dim lcDt As String

    ' First regex match to extract the portion containing LC Dt
    Set lcDtPortionObj = Application.Run("general_utility_functions.regExReturnedObj", lcText, "31c.+\n.+\n40e", True, True, True)
    If lcDtPortionObj Is Nothing Or lcDtPortionObj.Count = 0 Then
        ExtractLcDtBrac = vbNullString
        Exit Function
    End If

    ' Second regex match to extract the LC Dt from the portion
    Set lcDtPortionObj = Application.Run("general_utility_functions.regExReturnedObj", lcDtPortionObj(0), "\d+", True, True, True)
    If lcDtPortionObj Is Nothing Or lcDtPortionObj.Count < 2 Then
        ExtractLcDtBrac = vbNullString
        Exit Function
    End If

    ' Extract the Dt
    lcDt = Application.Run("utils.ReformatDateString", lcDtPortionObj(1))
    ExtractLcDtBrac = lcDt
End Function


