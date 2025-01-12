Attribute VB_Name = "Brac"
Option Explicit

Private Function ExtractPdfLcBrac(lcPath As String) As Object
    
    Dim lcText As String
    lcText = Application.Run("readPdf.ExtractTextFromPdfUsingAcrobatAcroHiliteList", lcPath)
    
    Dim resultDict As Object
    Set resultDict = CreateObject("Scripting.Dictionary")
    
    resultDict.Add "lcNo", Application.Run("Brac.ExtractLcNoBrac", lcText)
    resultDict.Add "lcDt", Application.Run("Brac.ExtractLcDtBrac", lcText)
    resultDict.Add "expiryDt", Application.Run("Brac.ExtractExpiryDtBrac", lcText)
    resultDict.Add "beneficiary", Application.Run("Brac.ExtractBeneficiaryBrac", lcText)

    Set ExtractPdfLcBrac = resultDict
    
End Function

Private Function ExtractLcNoBrac(lcText As String) As String
    
    ExtractLcNoBrac = Application.Run("utils.ExtractTextWithExcludeLines", lcText, "20.+\n.+\n31c", 1, 1)

End Function

Private Function ExtractLcDtBrac(lcText As String) As String
    Dim lcDtPortionObj As Object
    Dim lcDt As String

    lcDt = Application.Run("utils.ExtractTextWithExcludeLines", lcText, "31c.+\n.+\n40e", 1, 1)
    Set lcDtPortionObj = Application.Run("general_utility_functions.regExReturnedObj", lcDt, "\d+", True, True, True)
   
    If lcDtPortionObj Is Nothing Or lcDtPortionObj.Count <> 1 Then
        ExtractLcDtBrac = vbNullString
        Exit Function
    End If

    ExtractLcDtBrac = Application.Run("utils.ReformatDateString", lcDtPortionObj(0))
    
End Function

Private Function ExtractExpiryDtBrac(lcText As String) As String
    Dim expiryDtPortionObj As Object
    Dim expiryDt As String

    expiryDt = Application.Run("utils.ExtractTextWithExcludeLines", lcText, "31d.+\n.+\n50", 1, 1)
    Set expiryDtPortionObj = Application.Run("general_utility_functions.regExReturnedObj", expiryDt, "\d+", True, True, True)
   
    If expiryDtPortionObj Is Nothing Or expiryDtPortionObj.Count <> 1 Then
        ExtractExpiryDtBrac = vbNullString
        Exit Function
    End If

    ExtractExpiryDtBrac = Application.Run("utils.ReformatDateString", expiryDtPortionObj(0))

End Function

Private Function ExtractBeneficiaryBrac(lcText As String) As String
    Dim beneficiaryPortionObj As Object
    Dim beneficiary As String

    ' First regex match to extract the portion containing Beneficiary
    Set beneficiaryPortionObj = Application.Run("general_utility_functions.regExReturnedObj", lcText, "59([\s\S]*?)32B", True, True, True)
    If beneficiaryPortionObj Is Nothing Or beneficiaryPortionObj.Count = 0 Then
        ExtractBeneficiaryBrac = vbNullString
        Exit Function
    End If

    beneficiary = beneficiaryPortionObj(0)

    ' Second regex match to extract the Beneficiary from the portion
    Set beneficiaryPortionObj = Application.Run("general_utility_functions.regExReturnedObj", beneficiaryPortionObj(0), ".+", True, True, True)

    beneficiary = Replace(beneficiary, beneficiaryPortionObj(0) & Chr(10), "")
    beneficiary = Replace(beneficiary, beneficiaryPortionObj(beneficiaryPortionObj.Count - 1), "")
    beneficiary = Left(beneficiary,Len(beneficiary)-2) 'remove extra two line breck

    ExtractBeneficiaryBrac = beneficiary
End Function


