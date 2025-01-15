Attribute VB_Name = "readPdf"
Option Explicit

Function ExtractTextFromPdfUsingAcrobatJsMethod(pdfPath As String) As String
    Dim AcroApp As Object
    Dim AcroAVDoc As Object
    Dim AcroPDDoc As Object
    Dim jsObj As Object
    Dim extractedText As String
    Dim numPages As Long
    Dim pageNum As Long
    Dim numWords As Long
    Dim wordNum As Long
    Dim wordText As String

    On Error GoTo ErrorHandler

    ' Initialize Acrobat objects using late binding
    Set AcroApp = CreateObject("AcroExch.App")
    Set AcroAVDoc = CreateObject("AcroExch.AVDoc")
    
    ' Open the PDF file
    If AcroAVDoc.Open(pdfPath, "") Then
        Set AcroPDDoc = AcroAVDoc.GetPDDoc
        Set jsObj = AcroPDDoc.GetJSObject

        ' Get the total number of pages
        numPages = jsObj.numPages

        ' Initialize the extracted text
        extractedText = ""

        ' Loop through each page in the PDF
        For pageNum = 0 To numPages - 1
            ' Get the number of words on the current page
            numWords = jsObj.GetPageNumWords(pageNum)
            
            ' Loop through each word on the page
            For wordNum = 0 To numWords - 1
                wordText = jsObj.getPageNthWord(pageNum, wordNum, False)
                extractedText = extractedText & wordText
            Next wordNum
            ' Add a line break after each page
            extractedText = extractedText & vbCrLf & vbCrLf
        Next pageNum

        ' Close the document and exit Acrobat
        AcroAVDoc.Close True
        AcroApp.Exit

        ' Return the extracted text
        ExtractTextFromPdfUsingAcrobatJsMethod = extractedText
    Else
        MsgBox "Failed to open PDF file.", vbExclamation
    End If

CleanUp:
    ' Release the objects
    If Not AcroAVDoc Is Nothing Then Set AcroAVDoc = Nothing
    If Not AcroApp Is Nothing Then Set AcroApp = Nothing
    If Not AcroPDDoc Is Nothing Then Set AcroPDDoc = Nothing
    If Not jsObj Is Nothing Then Set jsObj = Nothing
    Exit Function

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    Resume CleanUp
End Function

Private Function ExtractTextFromPdfUsingAcrobatAcroHiliteList(filePath As String) As Object
    Dim AcroApp As Object
    Dim AcroDoc As Object
    Dim AcroPage As Object
    Dim AcroHiliteList As Object
    Dim AcroTextSelect As Object
    Dim pageNumber As Long
    Dim pageText As String
    Dim totalText As String
    Dim totalPages As Long
    Dim resultDict As Object
    Dim i As Long

    Set resultDict = CreateObject("Scripting.Dictionary")

    ' Initialize the total text variable
    totalText = ""

    ' Create Acrobat Application object
    Set AcroApp = CreateObject("AcroExch.App")
    Set AcroDoc = CreateObject("AcroExch.PDDoc")

    ' Open the PDF file
    If AcroDoc.Open(filePath) Then
        totalPages = AcroDoc.GetNumPages() ' Get total number of pages in the PDF
        resultDict.Add "totalPageCount", totalPages

        ' Loop through each page and extract text
        For pageNumber = 0 To totalPages - 1
            Set AcroPage = AcroDoc.AcquirePage(pageNumber)
            Set AcroHiliteList = CreateObject("AcroExch.HiliteList")
            AcroHiliteList.Add 0, 32767 ' Highlight all text on the page

            Set AcroTextSelect = AcroPage.CreatePageHilite(AcroHiliteList)
            If Not AcroTextSelect Is Nothing Then
                pageText = ""
                If Not resultDict.Exists("textPagesCount") Then resultDict.Add "textPagesCount", 0
                If Not resultDict.Exists("textPagesList") Then resultDict.Add "textPagesList", ""
                resultDict("textPagesCount") = resultDict("textPagesCount") + 1
                resultDict("textPagesList") = resultDict("textPagesList") & IIf(resultDict("textPagesList") = "", "", ",") & (pageNumber + 1)

                For i = 0 To AcroTextSelect.GetNumText - 1
                    pageText = pageText & AcroTextSelect.GetText(i) ' Extract text
                Next i
                totalText = totalText & vbCrLf & pageText
            Else
                If Not resultDict.Exists("blankPagesCount") Then resultDict.Add "blankPagesCount", 0
                If Not resultDict.Exists("blankPagesList") Then resultDict.Add "blankPagesList", ""
                resultDict("blankPagesCount") = resultDict("blankPagesCount") + 1
                resultDict("blankPagesList") = resultDict("blankPagesList") & IIf(resultDict("blankPagesList") = "", "", ",") & (pageNumber + 1)
            End If
        Next pageNumber

        ' Close the document
        AcroDoc.Close
    Else
        MsgBox "Failed to open PDF file.", vbExclamation
    End If

    ' Quit Acrobat
    AcroApp.Exit
    Set AcroApp = Nothing
    Set AcroDoc = Nothing
    Set AcroPage = Nothing
    Set AcroHiliteList = Nothing
    Set AcroTextSelect = Nothing

    resultDict.Add "totalText", totalText

    ' Return the extracted text
    Set ExtractTextFromPdfUsingAcrobatAcroHiliteList = resultDict
End Function
