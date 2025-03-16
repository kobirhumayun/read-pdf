Attribute VB_Name = "Mtb1"
Option Explicit

Private Function ReadMtb1Lcs(b2bPaths As Object) As Object
    Dim resultDict As Object
    Set resultDict = CreateObject("Scripting.Dictionary")

    Dim dicKey As Variant
    
    For Each dicKey In b2bPaths.Keys
        resultDict.Add resultDict.Count + 1, Application.Run("Mtb1.ExtractPdfLcMtb1", b2bPaths(dicKey))
    Next dicKey
    
    Set ReadMtb1Lcs = resultDict

End Function

Private Function ExtractPdfLcMtb1(lcPath As String) As Object
    
    Dim readPdf As Object
    Set readPdf = Application.Run("readPdf.ExtractTextFromPdfUsingAcrobatAcroHiliteList", lcPath)

    Dim lcText As String
    lcText = readPdf("totalText")
    
    Dim resultDict As Object
    Set resultDict = CreateObject("Scripting.Dictionary")
    
    resultDict.Add "lcNo", Application.Run("Mtb1.ExtractLcNoMtb1", lcText)
    resultDict.Add "lcDt", Application.Run("Mtb1.ExtractLcDtMtb1", lcText)
    resultDict.Add "expiryDt", Application.Run("Mtb1.ExtractExpiryDtMtb1", lcText)
    resultDict.Add "beneficiary", Application.Run("Mtb1.ExtractBeneficiaryMtb1", lcText)
    resultDict.Add "amount", Application.Run("Mtb1.ExtractAmountMtb1", lcText)
    resultDict.Add "shipmentDt", Application.Run("Mtb1.ExtractShipmentDtMtb1", lcText)
    resultDict.Add "pi", Application.Run("Mtb1.ExtractPiMtb1", lcText)

    resultDict.Add "pdfProperties", readPdf

    Set ExtractPdfLcMtb1 = resultDict
    
End Function

Private Function ExtractLcNoMtb1(lcText As String) As String
    
    ExtractLcNoMtb1 = Application.Run("utils.ExtractTextWithExcludeLines", lcText, "20.+\n.+\n31c", 1, 1)

End Function

Private Function ExtractLcDtMtb1(lcText As String) As String
    Dim lcDtPortionObj As Object
    Dim lcDt As String

    lcDt = Application.Run("utils.ExtractTextWithExcludeLines", lcText, "31c.+\n.+\n40e", 1, 1)
    Set lcDtPortionObj = Application.Run("general_utility_functions.regExReturnedObj", lcDt, "\d+", True, True, True)
   
    If lcDtPortionObj Is Nothing Or lcDtPortionObj.Count <> 1 Then
        ExtractLcDtMtb1 = vbNullString
        Exit Function
    End If

    ExtractLcDtMtb1 = Application.Run("utils.ReformatDateString", lcDtPortionObj(0))

End Function

Private Function ExtractExpiryDtMtb1(lcText As String) As String
    Dim expiryDtPortionObj As Object
    Dim expiryDt As String

    expiryDt = Application.Run("utils.ExtractTextWithExcludeLines", lcText, "31d.+\n.+\n50", 1, 1)
    Set expiryDtPortionObj = Application.Run("general_utility_functions.regExReturnedObj", expiryDt, "\d+", True, True, True)
   
    If expiryDtPortionObj Is Nothing Or expiryDtPortionObj.Count <> 1 Then
        ExtractExpiryDtMtb1 = vbNullString
        Exit Function
    End If

    ExtractExpiryDtMtb1 = Application.Run("utils.ReformatDateString", expiryDtPortionObj(0))

End Function

Private Function ExtractBeneficiaryMtb1(lcText As String) As String

    ExtractBeneficiaryMtb1 = Application.Run("utils.ExtractTextWithExcludeLines", lcText, "59 Beneficiary([\s\S]*?)32B", 1, 1)
    
End Function

Private Function ExtractAmountMtb1(lcText As String) As Variant
    Dim amountLineObj As Object
    Dim amountLine As String

    amountLine = Application.Run("utils.ExtractTextWithExcludeLines", lcText, "32B.+\n.+\n41D", 1, 1)

    Set amountLineObj = Application.Run("general_utility_functions.regExReturnedObj", amountLine, "(\d+\,\d+)|(\d+)", True, True, True)
   
    If amountLineObj Is Nothing Or amountLineObj.Count <> 1 Then
        ExtractAmountMtb1 = 0
        Exit Function
    End If

    ExtractAmountMtb1 = Replace(amountLineObj(0), ",", ".")
    
End Function

Private Function ExtractShipmentDtMtb1(lcText As String) As String
    Dim shipmentDtPortionObj As Object
    Dim shipmentDt As String

    shipmentDt = Application.Run("utils.ExtractTextWithExcludeLines", lcText, "44C.+\n.+\n45A", 1, 1)
    Set shipmentDtPortionObj = Application.Run("general_utility_functions.regExReturnedObj", shipmentDt, "\d+", True, True, True)
   
    If shipmentDtPortionObj Is Nothing Or shipmentDtPortionObj.Count <> 1 Then
        ExtractShipmentDtMtb1 = vbNullString
        Exit Function
    End If

    ExtractShipmentDtMtb1 = Application.Run("utils.ReformatDateString", shipmentDtPortionObj(0))

End Function

Private Function ExtractPiMtb1(lcText As String) As String
    Dim piPortionObj As Object
    Dim piPortion As String
    Dim piConcat As String
    Dim i As Long
    piConcat = vbNullString

    piPortion = Application.Run("utils.ExtractTextWithExcludeLines", lcText, "45A([\s\S]*?)46A", 1, 1)
    Set piPortionObj = Application.Run("general_utility_functions.regExReturnedObj", piPortion, "btl\/\d{2}\/\d{4}", True, True, True)
   
    If piPortionObj Is Nothing Or piPortionObj.Count < 1 Then
        ExtractPiMtb1 = vbNullString
        Exit Function
    End If

    for i = 0 To piPortionObj.Count - 1
        piConcat = piConcat & piPortionObj(i) & ", "
    Next i

    piConcat = Left(piConcat, Len(piConcat) - 2)

    ExtractPiMtb1 = piConcat

End Function
