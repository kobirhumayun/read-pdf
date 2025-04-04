Attribute VB_Name = "City"
Option Explicit

Private Function ReadCityLcs(b2bPaths As Object) As Object
    Dim resultDict As Object
    Set resultDict = CreateObject("Scripting.Dictionary")

    Dim dicKey As Variant
    
    For Each dicKey In b2bPaths.Keys
        Dim readPdf As Object
        Set readPdf = Application.Run("readPdf.ExtractTextFromPdfUsingAcrobatAcroHiliteList", b2bPaths(dicKey))
        resultDict.Add resultDict.Count + 1, Application.Run("City.ExtractPdfLcCity", readPdf)
    Next dicKey
    
    Set ReadCityLcs = resultDict

End Function

Private Function ExtractPdfLcCity(readPdf As Object) As Object

    Dim lcText As String
    lcText = readPdf("totalText")
    
    Dim resultDict As Object
    Set resultDict = CreateObject("Scripting.Dictionary")
    
    resultDict.Add "lcNo", Application.Run("City.ExtractLcNoCity", lcText)
    resultDict.Add "lcDt", Application.Run("City.ExtractLcDtCity", lcText)
    resultDict.Add "expiryDt", Application.Run("City.ExtractExpiryDtCity", lcText)
    resultDict.Add "beneficiary", Application.Run("City.ExtractBeneficiaryCity", lcText)
    resultDict.Add "amount", Application.Run("City.ExtractAmountCity", lcText)
    resultDict.Add "shipmentDt", Application.Run("City.ExtractShipmentDtCity", lcText)
    resultDict.Add "pi", Application.Run("City.ExtractPiCity", lcText)

    resultDict.Add "pdfProperties", readPdf

    Set ExtractPdfLcCity = resultDict
    
End Function

Private Function ExtractLcNoCity(lcText As String) As String
    
    ExtractLcNoCity = Application.Run("utils.ExtractTextWithExcludeLines", lcText, "20.+\n.+\n31c", 1, 1)

End Function

Private Function ExtractLcDtCity(lcText As String) As String
    Dim lcDtPortionObj As Object
    Dim lcDt As String

    lcDt = Application.Run("utils.ExtractTextWithExcludeLines", lcText, "31c.+\n.+\n40e", 1, 1)
    Set lcDtPortionObj = Application.Run("general_utility_functions.regExReturnedObj", lcDt, "\d+", True, True, True)
   
    If lcDtPortionObj Is Nothing Or lcDtPortionObj.Count <> 1 Then
        ExtractLcDtCity = vbNullString
        Exit Function
    End If

    ExtractLcDtCity = Application.Run("utils.ReformatDateString", lcDtPortionObj(0))

End Function

Private Function ExtractExpiryDtCity(lcText As String) As String
    Dim expiryDtPortionObj As Object
    Dim expiryDt As String

    expiryDt = Application.Run("utils.ExtractTextWithExcludeLines", lcText, "31d.+\n.+\n50", 1, 1)
    Set expiryDtPortionObj = Application.Run("general_utility_functions.regExReturnedObj", expiryDt, "\d+", True, True, True)
   
    If expiryDtPortionObj Is Nothing Or expiryDtPortionObj.Count <> 1 Then
        ExtractExpiryDtCity = vbNullString
        Exit Function
    End If

    ExtractExpiryDtCity = Application.Run("utils.ReformatDateString", expiryDtPortionObj(0))

End Function

Private Function ExtractBeneficiaryCity(lcText As String) As String

    ExtractBeneficiaryCity = Application.Run("utils.ExtractTextWithExcludeLines", lcText, "59.+Beneficiary([\s\S]*?)32B", 1, 1)
    
End Function

Private Function ExtractAmountCity(lcText As String) As Variant
    Dim amountLineObj As Object
    Dim amountLine As String

    amountLine = Application.Run("utils.ExtractTextWithExcludeLines", lcText, "32B.+\n.+\n41D", 1, 1)

    Set amountLineObj = Application.Run("general_utility_functions.regExReturnedObj", amountLine, "(\d+\,\d+)|(\d+)", True, True, True)
   
    If amountLineObj Is Nothing Or amountLineObj.Count <> 1 Then
        ExtractAmountCity = 0
        Exit Function
    End If

    ExtractAmountCity = Replace(amountLineObj(0), ",", ".")
    
End Function

Private Function ExtractShipmentDtCity(lcText As String) As String
    Dim shipmentDtPortionObj As Object
    Dim shipmentDt As String

    shipmentDt = Application.Run("utils.ExtractTextWithExcludeLines", lcText, "44C.+\n.+\n45A", 1, 1)
    Set shipmentDtPortionObj = Application.Run("general_utility_functions.regExReturnedObj", shipmentDt, "\d+", True, True, True)
   
    If shipmentDtPortionObj Is Nothing Or shipmentDtPortionObj.Count <> 1 Then
        ExtractShipmentDtCity = vbNullString
        Exit Function
    End If

    ExtractShipmentDtCity = Application.Run("utils.ReformatDateString", shipmentDtPortionObj(0))

End Function

Private Function ExtractPiCity(lcText As String) As String
    Dim piPortionObj As Object
    Dim piPortion As String
    Dim piConcat As String
    Dim i As Long
    piConcat = vbNullString

    piPortion = Application.Run("utils.ExtractTextWithExcludeLines", lcText, "45A([\s\S]*?)46A", 1, 1)
    Set piPortionObj = Application.Run("general_utility_functions.regExReturnedObj", piPortion, "btl\/\d{2}\/\d{4}", True, True, True)
   
    If piPortionObj Is Nothing Or piPortionObj.Count < 1 Then
        ExtractPiCity = vbNullString
        Exit Function
    End If

    for i = 0 To piPortionObj.Count - 1
        piConcat = piConcat & piPortionObj(i) & ", "
    Next i

    piConcat = Left(piConcat, Len(piConcat) - 2)

    ExtractPiCity = piConcat

End Function
