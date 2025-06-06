Attribute VB_Name = "Mtb"
Option Explicit

Private Function ReadMtbLcs(b2bPaths As Object) As Object
    Dim resultDict As Object
    Set resultDict = CreateObject("Scripting.Dictionary")

    Dim dicKey As Variant
    
    For Each dicKey In b2bPaths.Keys
        Dim readPdf As Object
        Set readPdf = Application.Run("readPdf.ExtractTextFromPdfUsingAcrobatAcroHiliteList", b2bPaths(dicKey))
        resultDict.Add resultDict.Count + 1, Application.Run("Mtb.ExtractPdfLcMtb", readPdf)
    Next dicKey
    
    Set ReadMtbLcs = resultDict

End Function

Private Function ExtractPdfLcMtb(readPdf As Object) As Object

    Dim lcText As String
    lcText = readPdf("totalText")
    
    Dim resultDict As Object
    Set resultDict = CreateObject("Scripting.Dictionary")
    
    resultDict.Add "lcNo", Application.Run("Mtb.ExtractLcNoMtb", lcText)
    resultDict.Add "lcDt", Application.Run("Mtb.ExtractLcDtMtb", lcText)
    resultDict.Add "expiryDt", Application.Run("Mtb.ExtractExpiryDtMtb", lcText)
    resultDict.Add "beneficiary", Application.Run("Mtb.ExtractBeneficiaryMtb", lcText)
    resultDict.Add "amount", Application.Run("Mtb.ExtractAmountMtb", lcText)
    resultDict.Add "shipmentDt", Application.Run("Mtb.ExtractShipmentDtMtb", lcText)
    resultDict.Add "pi", Application.Run("Mtb.ExtractPiMtb", lcText)

    resultDict.Add "pdfProperties", readPdf

    Set ExtractPdfLcMtb = resultDict
    
End Function

Private Function ExtractLcNoMtb(lcText As String) As String
    
    ExtractLcNoMtb = Application.Run("utils.ExtractTextWithExcludeLines", lcText, "f20.+\n.+\nf31c", 1, 1)

End Function

Private Function ExtractLcDtMtb(lcText As String) As String
    Dim lcDtPortionObj As Object
    Dim lcDt As String

    lcDt = Application.Run("utils.ExtractTextWithExcludeLines", lcText, "f31c.+\n.+\nf40e", 1, 1)
    Set lcDtPortionObj = Application.Run("general_utility_functions.regExReturnedObj", lcDt, "\d{6}", False, True, True)
   
    If lcDtPortionObj Is Nothing Or lcDtPortionObj.Count <> 1 Then
        ExtractLcDtMtb = vbNullString
        Exit Function
    End If

    ExtractLcDtMtb = Application.Run("utils.ReformatDateString", lcDtPortionObj(0))

End Function

Private Function ExtractExpiryDtMtb(lcText As String) As String
    Dim expiryDtPortionObj As Object
    Dim expiryDt As String

    expiryDt = Application.Run("utils.ExtractTextWithExcludeLines", lcText, "f31d.+\n.+\n.+\nf50", 1, 2)
    Set expiryDtPortionObj = Application.Run("general_utility_functions.regExReturnedObj", expiryDt, "\d{6}", False, True, True)
   
    If expiryDtPortionObj Is Nothing Or expiryDtPortionObj.Count <> 1 Then
        ExtractExpiryDtMtb = vbNullString
        Exit Function
    End If

    ExtractExpiryDtMtb = Application.Run("utils.ReformatDateString", expiryDtPortionObj(0))

End Function

Private Function ExtractBeneficiaryMtb(lcText As String) As String

    ExtractBeneficiaryMtb = Application.Run("utils.ExtractTextWithExcludeLines", lcText, "f59.+Beneficiary([\s\S]*?)f32B", 2, 1)
    
End Function

Private Function ExtractAmountMtb(lcText As String) As Variant
    Dim amountLineObj As Object
    Dim amountLine As String

    amountLine = Application.Run("utils.ExtractTextWithExcludeLines", lcText, "f32B.+\n.+\n.+\nf41D", 2, 1)

    Set amountLineObj = Application.Run("general_utility_functions.regExReturnedObj", amountLine, "(\d+\,\d+)|(\d+)", False, True, True)
   
    If amountLineObj Is Nothing Or amountLineObj.Count <> 1 Then
        ExtractAmountMtb = 0
        Exit Function
    End If

    ExtractAmountMtb = Replace(amountLineObj(0), ",", ".")
    
End Function

Private Function ExtractShipmentDtMtb(lcText As String) As String
    Dim shipmentDtPortionObj As Object
    Dim shipmentDt As String

    shipmentDt = Application.Run("utils.ExtractTextWithExcludeLines", lcText, "f44C.+\n.+\nf45A", 1, 1)
    Set shipmentDtPortionObj = Application.Run("general_utility_functions.regExReturnedObj", shipmentDt, "\d{6}", False, True, True)
   
    If shipmentDtPortionObj Is Nothing Or shipmentDtPortionObj.Count <> 1 Then
        ExtractShipmentDtMtb = vbNullString
        Exit Function
    End If

    ExtractShipmentDtMtb = Application.Run("utils.ReformatDateString", shipmentDtPortionObj(0))

End Function

Private Function ExtractPiMtb(lcText As String) As String
    Dim piPortionObj As Object
    Dim piPortion As String
    Dim piConcat As String
    Dim i As Long
    piConcat = vbNullString

    piPortion = Application.Run("utils.ExtractTextWithExcludeLines", lcText, "f45A([\s\S]*?)f46A", 1, 1)
    Set piPortionObj = Application.Run("general_utility_functions.regExReturnedObj", piPortion, "btl\/\d{2}\/\d{4}", True, True, True)
   
    If piPortionObj Is Nothing Or piPortionObj.Count < 1 Then
        ExtractPiMtb = vbNullString
        Exit Function
    End If

    for i = 0 To piPortionObj.Count - 1
        piConcat = piConcat & piPortionObj(i) & ", "
    Next i

    piConcat = Left(piConcat, Len(piConcat) - 2)

    ExtractPiMtb = piConcat

End Function
