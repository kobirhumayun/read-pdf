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

Private Function ExtractExpiryDtBrac(lcText As String) As String
    Dim expiryDtPortionObj As Object
    Dim expiryDt As String

    ' First regex match to extract the portion containing expiry Dt
    Set expiryDtPortionObj = Application.Run("general_utility_functions.regExReturnedObj", lcText, "31d.+\n.+\n50", True, True, True)
    If expiryDtPortionObj Is Nothing Or expiryDtPortionObj.Count = 0 Then
        ExtractExpiryDtBrac = vbNullString
        Exit Function
    End If

    ' Second regex match to extract the expiry Dt from the portion
    Set expiryDtPortionObj = Application.Run("general_utility_functions.regExReturnedObj", expiryDtPortionObj(0), "\d+", True, True, True)
    If expiryDtPortionObj Is Nothing Or expiryDtPortionObj.Count < 2 Then
        ExtractExpiryDtBrac = vbNullString
        Exit Function
    End If

    ' Extract the Dt
    expiryDt = Application.Run("utils.ReformatDateString", expiryDtPortionObj(1))
    ExtractExpiryDtBrac = expiryDt
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


