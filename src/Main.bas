Attribute VB_Name = "Main"
Option Explicit

Sub BracLc()

    Dim b2bPaths As Object
    Set b2bPaths = Application.Run("utils.GetSelectedFilePaths", "C:\Users\Humayun\Downloads\B2B", "Select Brac LC Only", "pdf")
    
    Dim resultDict As Object
    Set resultDict = CreateObject("Scripting.Dictionary")

    Dim tempDict As Object
    
    Dim dicKey As Variant
    
    For Each dicKey In b2bPaths.Keys
    
        Set tempDict = Application.Run("Brac.ExtractPdfLcBrac", b2bPaths(dicKey))
        Debug.Print tempDict("lcNo")
        resultDict.Add resultDict.Count + 1, tempDict
    
    Next dicKey
    
End Sub
