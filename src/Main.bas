Attribute VB_Name = "Main"
Option Explicit

Sub BracLc()

    Dim b2bPaths As Object
    Set b2bPaths = Application.Run("utils.GetSelectedFilePaths", "G:\PDL Customs\Export LC, Import LC & UP\Import LC With Related Doc\YEAR-2025", "Select Brac LC Only", "pdf")
    
    Dim resultDict As Object
    Set resultDict = Application.Run("Brac.ReadBracLcs", b2bPaths)

    Dim ws As Worksheet
    Set ws = ActiveWorkbook.ActiveSheet

    Dim printB2bInfo As Boolean
    printB2bInfo = Application.Run("utils.PutB2bDataToWs", resultDict, ws, True, 1, 1, True)
    
End Sub
