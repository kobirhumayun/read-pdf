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
    printB2bInfo = Application.Run("utils.PutB2bDataToWs", resultDict, ws, True, 1, 1, True, True)
    
End Sub

Sub AlArafahLc()

    Dim b2bPaths As Object
    Set b2bPaths = Application.Run("utils.GetSelectedFilePaths", "G:\PDL Customs\Export LC, Import LC & UP\Import LC With Related Doc\YEAR-2025", "Select AlArafah LC Only", "pdf")
    
    Dim resultDict As Object
    Set resultDict = Application.Run("AlArafah.ReadAlArafahLcs", b2bPaths)

    Dim ws As Worksheet
    Set ws = ActiveWorkbook.ActiveSheet

    Dim printB2bInfo As Boolean
    printB2bInfo = Application.Run("utils.PutB2bDataToWs", resultDict, ws, True, 1, 1, True, True)
    
End Sub

Sub CityLc()

    Dim b2bPaths As Object
    Set b2bPaths = Application.Run("utils.GetSelectedFilePaths", "G:\PDL Customs\Export LC, Import LC & UP\Import LC With Related Doc\YEAR-2025", "Select City LC Only", "pdf")
    
    Dim resultDict As Object
    Set resultDict = Application.Run("City.ReadCityLcs", b2bPaths)

    Dim ws As Worksheet
    Set ws = ActiveWorkbook.ActiveSheet

    Dim printB2bInfo As Boolean
    printB2bInfo = Application.Run("utils.PutB2bDataToWs", resultDict, ws, True, 1, 1, True, True)
    
End Sub

Sub MtbLc()

    Dim b2bPaths As Object
    Set b2bPaths = Application.Run("utils.GetSelectedFilePaths", "G:\PDL Customs\Export LC, Import LC & UP\Import LC With Related Doc\YEAR-2025", "Select Mtb LC Only", "pdf")
    
    Dim resultDict As Object
    Set resultDict = Application.Run("Mtb.ReadMtbLcs", b2bPaths)

    Dim ws As Worksheet
    Set ws = ActiveWorkbook.ActiveSheet

    Dim printB2bInfo As Boolean
    printB2bInfo = Application.Run("utils.PutB2bDataToWs", resultDict, ws, True, 1, 1, True, True)
    
End Sub

Sub Mtb1Lc()

    Dim b2bPaths As Object
    Set b2bPaths = Application.Run("utils.GetSelectedFilePaths", "G:\PDL Customs\Export LC, Import LC & UP\Import LC With Related Doc\YEAR-2025", "Select Mtb1 LC Only", "pdf")
    
    Dim resultDict As Object
    Set resultDict = Application.Run("Mtb1.ReadMtb1Lcs", b2bPaths)

    Dim ws As Worksheet
    Set ws = ActiveWorkbook.ActiveSheet

    Dim printB2bInfo As Boolean
    printB2bInfo = Application.Run("utils.PutB2bDataToWs", resultDict, ws, True, 1, 1, True, True)
    
End Sub

Sub ScbLc()

    Dim b2bPaths As Object
    Set b2bPaths = Application.Run("utils.GetSelectedFilePaths", "G:\PDL Customs\Export LC, Import LC & UP\Import LC With Related Doc\YEAR-2025", "Select Scb LC Only", "pdf")
    
    Dim resultDict As Object
    Set resultDict = Application.Run("Scb.ReadScbLcs", b2bPaths)

    Dim ws As Worksheet
    Set ws = ActiveWorkbook.ActiveSheet

    Dim printB2bInfo As Boolean
    printB2bInfo = Application.Run("utils.PutB2bDataToWs", resultDict, ws, True, 1, 1, True, True)
    
End Sub

Sub AnyBankLc()

    Dim b2bPaths As Object
    Set b2bPaths = Application.Run("utils.GetSelectedFilePaths", "G:\PDL Customs\Export LC, Import LC & UP\Import LC With Related Doc\YEAR-2025", "Select any bank LC", "pdf")
    
    Dim resultDict As Object
    Set resultDict = CreateObject("Scripting.Dictionary")

    Dim dicKey As Variant
    
    For Each dicKey In b2bPaths.Keys
        Dim readPdf As Object
        Set readPdf = Application.Run("readPdf.ExtractTextFromPdfUsingAcrobatAcroHiliteList", b2bPaths(dicKey))

       resultDict.Add resultDict.Count + 1, Application.Run("utils.ExtractAnyBankLc", readPdf)

    Next dicKey

    Dim ws As Worksheet
    Set ws = ActiveWorkbook.ActiveSheet

    Dim printB2bInfo As Boolean
    printB2bInfo = Application.Run("utils.PutB2bDataToWs", resultDict, ws, True, 1, 1, False, False)
    
    
End Sub

Sub PrintB2BLc()

    Dim b2bPaths As Object
    Set b2bPaths = Application.Run("utils.GetSelectedFilePaths", "G:\PDL Customs\Export LC, Import LC & UP\Import LC With Related Doc\YEAR-2025", "Select any bank LC", "pdf")
    
    Dim resultDict As Object
    Set resultDict = CreateObject("Scripting.Dictionary")
    Dim lcNoMismatchWithFileNameDict As Object
    Set lcNoMismatchWithFileNameDict = CreateObject("Scripting.Dictionary")
    Dim tempDict As Object
    Dim printPageRangeDict As Object
    Dim printStatus As Boolean

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim fileName As String

    Dim dicKey As Variant
    
    For Each dicKey In b2bPaths.Keys
        Dim readPdf As Object
        Set readPdf = Application.Run("readPdf.ExtractTextFromPdfUsingAcrobatAcroHiliteList", b2bPaths(dicKey))

        Set tempDict = Application.Run("utils.ExtractAnyBankLc", readPdf)

        If tempDict("bankName") = "Unknown" Then
            tempDict.Add "lcNo",  b2bPaths(dicKey) ' set file path to inspect
        End If

        fileName = fso.GetBaseName(b2bPaths(dicKey))

        If tempDict("lcNo") <> fileName Then
            lcNoMismatchWithFileNameDict.Add lcNoMismatchWithFileNameDict.Count + 1, fileName
        End If

        Set printPageRangeDict = Application.Run("utils.GetPageRangeForPrint", tempDict)
        printStatus = Application.Run("utils.PrintPdfPageRange", b2bPaths(dicKey), printPageRangeDict("startPage"), printPageRangeDict("endPage"))
        ' MsgBox "start page " & printPageRangeDict("startPage") & "end page " & printPageRangeDict("endPage")

       resultDict.Add resultDict.Count + 1, tempDict

    Next dicKey

    Application.Run "JsonUtilityFunction.SaveDictionaryToJsonTextFile", resultDict, "D:\Temp\UP Draft\Draft 2025\b2b-print-data\" & _
       "printed-b2b-" & Application.Run("utils.GetTimestampForFilename") & ".json"

    Dim ws As Worksheet
    Set ws = ActiveWorkbook.ActiveSheet

    Dim printB2bInfo As Boolean
    printB2bInfo = Application.Run("utils.PutB2bDataToWs", resultDict, ws, True, 1, 1, False, False)

    Dim msgStr As String

    If lcNoMismatchWithFileNameDict.Count > 0 Then
        For Each dicKey In lcNoMismatchWithFileNameDict.Keys
            msgStr = msgStr & Chr(10) & lcNoMismatchWithFileNameDict(1)
        Next dicKey
            MsgBox msgStr, ,"List of LC that mismatch with file name"
    End If
    
    
End Sub
