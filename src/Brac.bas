Attribute VB_Name = "Brac"
Option Explicit

Private Function ReadBracLcs(b2bPaths As Object) As Object
    Dim resultDict As Object
    Set resultDict = CreateObject("Scripting.Dictionary")

    Dim dicKey As Variant
    
    For Each dicKey In b2bPaths.Keys
        Dim readPdf As Object
        Set readPdf = Application.Run("readPdf.ExtractTextFromPdfUsingAcrobatAcroHiliteList", b2bPaths(dicKey))
        resultDict.Add resultDict.Count + 1, Application.Run("Brac.ExtractPdfLcBrac", readPdf)
    Next dicKey
    
    Set ReadBracLcs = resultDict

End Function

Private Function ExtractPdfLcBrac(readPdf As Object) As Object

    Dim lcText As String
    lcText = readPdf("totalText")
    
    Dim resultDict As Object
    Set resultDict = CreateObject("Scripting.Dictionary")
    
    resultDict.Add "lcNo", Application.Run("Brac.ExtractLcNoBrac", lcText)
    resultDict.Add "lcDt", Application.Run("Brac.ExtractLcDtBrac", lcText)
    resultDict.Add "expiryDt", Application.Run("Brac.ExtractExpiryDtBrac", lcText)
    resultDict.Add "beneficiary", Application.Run("Brac.ExtractBeneficiaryBrac", lcText)
    resultDict.Add "amount", Application.Run("Brac.ExtractAmountBrac", lcText)
    resultDict.Add "shipmentDt", Application.Run("Brac.ExtractShipmentDtBrac", lcText)
    resultDict.Add "pi", Application.Run("Brac.ExtractPiBrac", lcText)

    resultDict.Add "pdfProperties", readPdf

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

    ExtractBeneficiaryBrac = Application.Run("utils.ExtractTextWithExcludeLines", lcText, "59 Beneficiary([\s\S]*?)32B", 1, 1)
    
End Function

Private Function ExtractAmountBrac(lcText As String) As Variant
    Dim amountLineObj As Object
    Dim amountLine As String

    amountLine = Application.Run("utils.ExtractTextWithExcludeLines", lcText, "32B.+\n.+\n41D", 1, 1)

    Set amountLineObj = Application.Run("general_utility_functions.regExReturnedObj", amountLine, "(\d+\,\d+)|(\d+)", True, True, True)
   
    If amountLineObj Is Nothing Or amountLineObj.Count <> 1 Then
        ExtractAmountBrac = 0
        Exit Function
    End If

    ExtractAmountBrac = Replace(amountLineObj(0), ",", ".")
    
End Function

Private Function ExtractShipmentDtBrac(lcText As String) As String
    Dim shipmentDtPortionObj As Object
    Dim shipmentDt As String

    shipmentDt = Application.Run("utils.ExtractTextWithExcludeLines", lcText, "44C.+\n.+\n45A", 1, 1)
    Set shipmentDtPortionObj = Application.Run("general_utility_functions.regExReturnedObj", shipmentDt, "\d+", True, True, True)
   
    If shipmentDtPortionObj Is Nothing Or shipmentDtPortionObj.Count <> 1 Then
        ExtractShipmentDtBrac = vbNullString
        Exit Function
    End If

    ExtractShipmentDtBrac = Application.Run("utils.ReformatDateString", shipmentDtPortionObj(0))

End Function

Private Function ExtractPiBrac(lcText As String) As String
    Dim piPortionObj As Object
    Dim piPortion As String
    Dim piConcat As String
    Dim i As Long
    piConcat = vbNullString

    piPortion = Application.Run("utils.ExtractTextWithExcludeLines", lcText, "45A([\s\S]*?)46A", 1, 1)
    Set piPortionObj = Application.Run("general_utility_functions.regExReturnedObj", piPortion, "btl\/\d{2}\/\d{4}", True, True, True)
   
    If piPortionObj Is Nothing Or piPortionObj.Count < 1 Then
        ExtractPiBrac = vbNullString
        Exit Function
    End If

    for i = 0 To piPortionObj.Count - 1
        piConcat = piConcat & piPortionObj(i) & ", "
    Next i

    piConcat = Left(piConcat, Len(piConcat) - 2)

    ExtractPiBrac = piConcat

End Function
