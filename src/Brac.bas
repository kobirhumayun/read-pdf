Attribute VB_Name = "Brac"
Option Explicit

Private Function ExtractPdfLcBrac(lcPath As String) As Object
    
    Dim lcText As String
    lcText = Application.Run("readPdf.ExtractTextFromPdfUsingAcrobatAcroHiliteList", lcPath)
    
    Dim resultDict As Object
    Set resultDict = CreateObject("Scripting.Dictionary")
    
    resultDict.Add "lcNo", Application.Run("Brac.ExtractLcNoBrac", lcText)
    
    Set ExtractPdfLcBrac = resultDict
    
End Function

Private Function ExtractLcNoBrac(lcText As String) As String

    Dim lcPortionObj As Object
    Set lcPortionObj = Application.Run("general_utility_functions.regExReturnedObj", lcText, "20.+\n.+\n31c", True, True, True)
    Set lcPortionObj = Application.Run("general_utility_functions.regExReturnedObj", lcPortionObj(0), ".+", True, True, True)
    
    Dim lcNo As String
    lcNo = lcPortionObj(1)

    ExtractLcNoBrac = lcNo
    
End Function

