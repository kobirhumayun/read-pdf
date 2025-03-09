Attribute VB_Name = "Scb"
Option Explicit

Private Function ReadScbLcs(b2bPaths As Object) As Object
    Dim resultDict As Object
    Set resultDict = CreateObject("Scripting.Dictionary")

    Dim dicKey As Variant
    
    For Each dicKey In b2bPaths.Keys
        resultDict.Add resultDict.Count + 1, Application.Run("Scb.ExtractPdfLcScb", b2bPaths(dicKey))
    Next dicKey
    
    Set ReadScbLcs = resultDict

End Function

Private Function ExtractPdfLcScb(lcPath As String) As Object
    
    Dim readPdf As Object
    Set readPdf = Application.Run("readPdf.ExtractTextFromPdfUsingAcrobatAcroHiliteList", lcPath)

    Dim lcText As String
    lcText = readPdf("totalText")
    
    Dim resultDict As Object
    Set resultDict = CreateObject("Scripting.Dictionary")
    
    resultDict.Add "lcNo", Application.Run("Scb.ExtractLcNoScb", lcText)
    resultDict.Add "lcDt", Application.Run("Scb.ExtractLcDtScb", lcText)
    resultDict.Add "expiryDt", Application.Run("Scb.ExtractExpiryDtScb", lcText)
    resultDict.Add "beneficiary", Application.Run("Scb.ExtractBeneficiaryScb", lcText)
    resultDict.Add "amount", Application.Run("Scb.ExtractAmountScb", lcText)
    resultDict.Add "shipmentDt", Application.Run("Scb.ExtractShipmentDtScb", lcText)
    resultDict.Add "pi", Application.Run("Scb.ExtractPiScb", lcText)

    resultDict.Add "pdfProperties", readPdf

    Set ExtractPdfLcScb = resultDict
    
End Function

Private Function ExtractLcNoScb(lcText As String) As String
    
    ExtractLcNoScb = Application.Run("utils.ExtractTextWithExcludeLines", lcText, "20.+\n.+\n31c", 1, 1)

End Function

Private Function ExtractLcDtScb(lcText As String) As String
    Dim lcDtPortionObj As Object
    Dim lcDt As String

    lcDt = Application.Run("utils.ExtractTextWithExcludeLines", lcText, "31c.+\n.+\n40e", 1, 1)
    Set lcDtPortionObj = Application.Run("general_utility_functions.regExReturnedObj", lcDt, "\d+", True, True, True)
   
    If lcDtPortionObj Is Nothing Or lcDtPortionObj.Count <> 1 Then
        ExtractLcDtScb = vbNullString
        Exit Function
    End If

    ExtractLcDtScb = Application.Run("utils.ReformatDateString", lcDtPortionObj(0))

End Function

Private Function ExtractExpiryDtScb(lcText As String) As String
    Dim expiryDtPortionObj As Object
    Dim expiryDt As String

    expiryDt = Application.Run("utils.ExtractTextWithExcludeLines", lcText, "31d.+\n.+\n50", 1, 1)
    Set expiryDtPortionObj = Application.Run("general_utility_functions.regExReturnedObj", expiryDt, "\d+", True, True, True)
   
    If expiryDtPortionObj Is Nothing Or expiryDtPortionObj.Count <> 1 Then
        ExtractExpiryDtScb = vbNullString
        Exit Function
    End If

    ExtractExpiryDtScb = Application.Run("utils.ReformatDateString", expiryDtPortionObj(0))

End Function

Private Function ExtractBeneficiaryScb(lcText As String) As String

    ExtractBeneficiaryScb = Application.Run("utils.ExtractTextWithExcludeLines", lcText, "59 Beneficiary([\s\S]*?)32B", 1, 1)
    
End Function

Private Function ExtractAmountScb(lcText As String) As Variant
    Dim amountLineObj As Object
    Dim amountLine As String

    amountLine = Application.Run("utils.ExtractTextWithExcludeLines", lcText, "32B.+\n.+\n41D", 1, 1)

    Set amountLineObj = Application.Run("general_utility_functions.regExReturnedObj", amountLine, "(\d+\,\d+)|(\d+)", True, True, True)
   
    If amountLineObj Is Nothing Or amountLineObj.Count <> 1 Then
        ExtractAmountScb = 0
        Exit Function
    End If

    ExtractAmountScb = Replace(amountLineObj(0), ",", ".")
    
End Function

Private Function ExtractShipmentDtScb(lcText As String) As String
    Dim shipmentDtPortionObj As Object
    Dim shipmentDt As String

    shipmentDt = Application.Run("utils.ExtractTextWithExcludeLines", lcText, "44C.+\n.+\n45A", 1, 1)
    Set shipmentDtPortionObj = Application.Run("general_utility_functions.regExReturnedObj", shipmentDt, "\d+", True, True, True)
   
    If shipmentDtPortionObj Is Nothing Or shipmentDtPortionObj.Count <> 1 Then
        ExtractShipmentDtScb = vbNullString
        Exit Function
    End If

    ExtractShipmentDtScb = Application.Run("utils.ReformatDateString", shipmentDtPortionObj(0))

End Function

Private Function ExtractPiScb(lcText As String) As String
    Dim piPortionObj As Object
    Dim piPortion As String
    Dim piConcat As String
    Dim i As Long
    piConcat = vbNullString

    piPortion = Application.Run("utils.ExtractTextWithExcludeLines", lcText, "45A([\s\S]*?)46A", 1, 1)
    Set piPortionObj = Application.Run("general_utility_functions.regExReturnedObj", piPortion, "btl\/\d{2}\/\d{4}", True, True, True)
   
    If piPortionObj Is Nothing Or piPortionObj.Count < 1 Then
        ExtractPiScb = vbNullString
        Exit Function
    End If

    for i = 0 To piPortionObj.Count - 1
        piConcat = piConcat & piPortionObj(i) & ", "
    Next i

    piConcat = Left(piConcat, Len(piConcat) - 2)

    ExtractPiScb = piConcat

End Function
