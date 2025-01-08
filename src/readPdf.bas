Attribute VB_Name = "readPdf"
Option Explicit

Function ExtractTextFromPdfUsingAcrobatJsMethod(pdfPath As String) As String
    Dim AcroApp As Object
    Dim AcroAVDoc As Object
    Dim AcroPDDoc As Object
    Dim jsObj As Object
    Dim extractedText As String
    Dim numPages As Integer
    Dim pageNum As Integer
    Dim numWords As Integer
    Dim wordNum As Integer
    Dim wordText As String
    Dim sb As Object ' Use StringBuilder for better performance

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

        ' Initialize StringBuilder
        Set sb = CreateObject("System.Text.StringBuilder")

        ' Loop through each page in the PDF
        For pageNum = 0 To numPages - 1
            ' Get the number of words on the current page
            numWords = jsObj.GetPageNumWords(pageNum)
            
            ' Loop through each word on the page
            For wordNum = 0 To numWords - 1
                wordText = jsObj.getPageNthWord(pageNum, wordNum, False)
                sb.Append wordText & " "
            Next wordNum
            ' Add a line break after each page
            sb.Append vbCrLf & vbCrLf
        Next pageNum

        ' Close the document and exit Acrobat
        AcroAVDoc.Close True
        AcroApp.Exit

        ' Return the extracted text
        ExtractTextFromPdfUsingAcrobatJsMethod = sb.ToString
    Else
        MsgBox "Failed to open PDF file.", vbExclamation
    End If

CleanUp:
    ' Release the objects
    If Not AcroAVDoc Is Nothing Then Set AcroAVDoc = Nothing
    If Not AcroApp Is Nothing Then Set AcroApp = Nothing
    If Not AcroPDDoc Is Nothing Then Set AcroPDDoc = Nothing
    If Not jsObj Is Nothing Then Set jsObj = Nothing
    If Not sb Is Nothing Then Set sb = Nothing
    Exit Function

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    Resume CleanUp
End Function
