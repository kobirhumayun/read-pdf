Attribute VB_Name = "utils"
Option Explicit

Function GetSelectedFilePaths(initialFolderPath As String, dialogTitle As String, Optional fileType As String = "") As Object

  ' Declare variables
  Dim fileDialog As fileDialog
  Dim selectedFile As Variant
  Dim filePaths As Object

  ' Create a FileDialog object
  Set fileDialog = Application.fileDialog(msoFileDialogFilePicker)

  ' Set FileDialog properties
  With fileDialog
    .Title = dialogTitle
    .InitialFileName = initialFolderPath
    .AllowMultiSelect = True
    
    ' Set file type filter if provided
    If fileType <> "" Then
      .Filters.Add "Files", "*." & fileType, 1
    End If
    
  End With

  ' Show the File Picker dialog
  If fileDialog.Show = -1 Then ' -1 indicates a file was selected

    ' Create a dictionary to store file paths
    Set filePaths = CreateObject("Scripting.Dictionary")

    ' Loop through each selected file
    For Each selectedFile In fileDialog.SelectedItems
      ' Add the full file path to the dictionary
      filePaths.Add selectedFile, selectedFile
    Next selectedFile

  End If

  ' Return the dictionary of file paths
  Set GetSelectedFilePaths = filePaths

End Function

Private Function WriteStringToTexFile(text As String, filePath As String) As Boolean
    On Error GoTo ErrorHandler ' Enable error handling

    Dim fileNumber As Integer
    fileNumber = FreeFile ' Get a free file number

    ' Open the file for output
    Open filePath For Output As #fileNumber
    Print #fileNumber, text
    Close #fileNumber
    
    WriteStringToTexFile = True ' Indicate success
    Exit Function

ErrorHandler:
    WriteStringToTexFile = False ' Indicate failure
    If fileNumber <> 0 Then Close #fileNumber ' Ensure the file is closed if an error occurred
End Function

Private Function ReadTextFile(filePath As String) As String
    On Error GoTo ErrorHandler ' Enable error handling

    Dim text As String
    Dim fileNumber As Integer
    fileNumber = FreeFile ' Get a free file number

    ' Open the file for input
    Open filePath For Input As #fileNumber
    text = Input$(LOF(fileNumber), fileNumber) ' Read the entire file
    Close #fileNumber
    
    ReadTextFile = text ' Return the read text
    Exit Function

ErrorHandler:
    ReadTextFile = "" ' Return an empty string in case of an error
    If fileNumber <> 0 Then Close #fileNumber ' Ensure the file is closed if an error occurred
End Function

Function ReformatDateString(dateStr As String) As String
  Dim dayPart As String, monthPart As String, yearPart As String
  Dim result As String

    ' Check if the input string is exactly 6 characters and numeric
  If Len(dateStr) = 6 And IsNumeric(dateStr) Then
      ' Extract day, month, and year parts
    yearPart = Mid(dateStr, 1, 2)
    monthPart = Mid(dateStr, 3, 2)
    dayPart = Mid(dateStr, 5, 2)
    
      ' Construct the date string in DD/MM/YY format
    result = dayPart & "/" & monthPart & "/" & yearPart
    
      ' Return the result
    ReformatDateString = result
  Else
    ReformatDateString = "Invalid date format."
  End If
End Function

Private Function ExtractTextWithExcludeLines(InputText As String, Pattern As String, ExcludeLinesStart As Long, ExcludeLinesEnd As Long) As String
  Dim regEx As Object
  Dim matches As Object
  Dim resultText As String
  Dim lines() As String
  Dim i As Long
  Dim startIndex As Long, endIndex As Long
  
  ' Create a RegExp object
  Set regEx = CreateObject("VBScript.RegExp")
  regEx.Global = False        ' Only find the first occurrence
  regEx.IgnoreCase = True
  regEx.Multiline = True      ' Treat input as multiple lines
  regEx.Pattern = Pattern
    
  ' Execute the regex on the input text
  Set matches = regEx.Execute(InputText)
  
  ' Check if a match was found
  If matches.Count > 0 Then
    ' Get the full matched text
    resultText = matches(0).Value
    
    ' Normalize line breaks to Line Feed (vbLf) to handle different newline characters
    resultText = Replace(resultText, vbCrLf, vbLf)
    resultText = Replace(resultText, vbCr, vbLf)
    
    ' Split the text into lines
    lines = Split(resultText, vbLf)
    
    ' Calculate start and end indices after excluding specified lines
    startIndex = LBound(lines) + ExcludeLinesStart
    endIndex = UBound(lines) - ExcludeLinesEnd
    
    ' Ensure indices are within bounds
    If startIndex > endIndex Or startIndex > UBound(lines) Or endIndex < LBound(lines) Then
      ' No lines to return
      ExtractTextWithExcludeLines = ""
      Exit Function
    ElseIf startIndex < LBound(lines) Then
      startIndex = LBound(lines)
    End If

    If endIndex > UBound(lines) Then
      endIndex = UBound(lines)
    End If
    
    ' Concatenate the selected lines
    Dim selectedLines As String
    For i = startIndex To endIndex
      selectedLines = selectedLines & lines(i) & vbLf
    Next i
    
    ' Remove the trailing newline character and replace Line Feeds with Carriage Return + Line Feed for Windows compatibility
    If Right(selectedLines, 1) = vbLf Then
      selectedLines = Left(selectedLines, Len(selectedLines) - 1)
    End If
    ExtractTextWithExcludeLines = Replace(selectedLines, vbLf, vbCrLf)
  Else
    ' No match found
    ExtractTextWithExcludeLines = ""
  End If
    
End Function

Private Function PutB2bDataToWs(resultDict As Object, ws As Worksheet, printHeaders As Boolean, startRow As Long, startColumn As Long, printShipExpAndOthers As Boolean, printPdfProperties As Boolean) As Boolean
  ' Declare variables
  Dim tempDict As Object
  
  Dim dicKey As Variant
  
  Dim row As Long
  row = startRow

  Dim columns As Long
  columns = startColumn
  
  ' Print headers if printHeaders is True
  If printHeaders Then
    ws.Cells(row, columns).Value = "LC No"
    ws.Cells(row, columns + 1).Value = "LC Date"
    ws.Cells(row, columns + 2).Value = "Amount"

    If printShipExpAndOthers Then
      ws.Cells(row, columns + 3).Value = "Shipment Date"
      ws.Cells(row, columns + 4).Value = "Expiry Date"
      ws.Cells(row, columns + 5).Value = "Beneficiary"
      ws.Cells(row, columns + 6).Value = "PI"
    End If

    If printPdfProperties Then
      ws.Cells(row, columns + 7).Value = "Page Count"
      ws.Cells(row, columns + 8).Value = "Text Page Count"
      ws.Cells(row, columns + 9).Value = "Text Page List"
      ws.Cells(row, columns + 10).Value = "Blank Page Count"
      ws.Cells(row, columns + 11).Value = "Blank Page List"
    End If
    
    row = row + 1
  End If
  
  For Each dicKey In resultDict.Keys
  
      Set tempDict = resultDict(dicKey)
      
      ws.Cells(row, columns).Value = tempDict("lcNo")
      If tempDict("lcDt") <> "" Then
          ws.Cells(row, columns + 1).Value = IIf(IsDate(tempDict("lcDt")), CDate(tempDict("lcDt")), tempDict("lcDt"))
      End If
      ws.Cells(row, columns + 2).Value = tempDict("amount")

      If printShipExpAndOthers Then

        If tempDict("shipmentDt") <> "" Then
          ws.Cells(row, columns + 3).Value = IIf(IsDate(tempDict("shipmentDt")), CDate(tempDict("shipmentDt")), tempDict("shipmentDt"))
        End If
        If tempDict("expiryDt") <> "" Then
            ws.Cells(row, columns + 4).Value = IIf(IsDate(tempDict("expiryDt")), CDate(tempDict("expiryDt")), tempDict("expiryDt"))
        End If
        ws.Cells(row, columns + 5).Value = tempDict("beneficiary")
        ws.Cells(row, columns + 6).Value = tempDict("pi")
        
      End If

      If printPdfProperties Then
        ws.Cells(row, columns + 7).Value = tempDict("pdfProperties")("totalPageCount")
        ws.Cells(row, columns + 8).Value = tempDict("pdfProperties")("textPagesCount")
        ws.Cells(row, columns + 9).Value = tempDict("pdfProperties")("textPagesList")
        ws.Cells(row, columns + 10).Value = tempDict("pdfProperties")("blankPagesCount")
        ws.Cells(row, columns + 11).Value = tempDict("pdfProperties")("blankPagesList")
      End If

      row = row + 1
  
  Next dicKey 

  PutB2bDataToWs = True ' Indicate success

End Function

Private Function LcOfWhichBank(readPdf As Object) As Object

  Dim lcText As String
  lcText = readPdf("totalText")
  
  Dim resultDict As Object
  Set resultDict = CreateObject("Scripting.Dictionary")
  
  Dim bankName As String

  bankName =  Application.Run("AlArafah.ExtractLcNoAlArafah", lcText)

  If bankName <> "" And Left(bankName, 4) = "1080" Then

    bankName = "AlArafah"
    resultDict.Add "bankName", bankName
    Set LcOfWhichBank = resultDict
    Exit Function
    
  End If
  
  bankName =  Application.Run("Brac.ExtractLcNoBrac", lcText)

  If bankName <> "" And Left(bankName, 4) = "3085" Then

    bankName = "Brac"
    resultDict.Add "bankName", bankName
    Set LcOfWhichBank = resultDict
    Exit Function

  End If

  bankName =  Application.Run("City.ExtractLcNoCity", lcText)

  If bankName <> "" And Left(bankName, 5) = "07422" Then

    bankName = "City"
    resultDict.Add "bankName", bankName
    Set LcOfWhichBank = resultDict
    Exit Function

  End If

  bankName =  Application.Run("Mtb.ExtractLcNoMtb", lcText)

  If bankName <> "" And Left(bankName, 7) = "0002228" Then

    bankName = "Mtb"
    resultDict.Add "bankName", bankName
    Set LcOfWhichBank = resultDict
    Exit Function

  End If

  bankName =  Application.Run("Mtb1.ExtractLcNoMtb1", lcText)

  If bankName <> "" And Left(bankName, 7) = "0002228" Then

    bankName = "Mtb1"
    resultDict.Add "bankName", bankName
    Set LcOfWhichBank = resultDict
    Exit Function

  End If

  bankName =  Application.Run("Scb.ExtractLcNoScb", lcText)

  If bankName <> "" And Left(bankName, 4) = "4110" Then

    bankName = "Scb"
    resultDict.Add "bankName", bankName
    Set LcOfWhichBank = resultDict
    Exit Function

  End If

  resultDict.Add "bankName", "Unknown"
  Set LcOfWhichBank = resultDict

End Function

Private Function ExtractAnyBankLc(readPdf As Object) As Object
  
  Dim resultDict As Object
  Set resultDict = CreateObject("Scripting.Dictionary")
  Dim bankName As String
  Dim bankNameDict As Object
  Set bankNameDict = Application.Run("utils.LcOfWhichBank", readPdf)

  bankName = bankNameDict("bankName")

  If bankName = "AlArafah" Then
      Set resultDict = Application.Run("AlArafah.ExtractPdfLcAlArafah", readPdf)
  ElseIf bankName = "Brac" Then
      Set resultDict = Application.Run("Brac.ExtractPdfLcBrac", readPdf)
  ElseIf bankName = "City" Then
      Set resultDict = Application.Run("City.ExtractPdfLcCity", readPdf)
  ElseIf bankName = "Mtb" Then
      Set resultDict = Application.Run("Mtb.ExtractPdfLcMtb", readPdf)
  ElseIf bankName = "Mtb1" Then
      Set resultDict = Application.Run("Mtb1.ExtractPdfLcMtb1", readPdf)
  ElseIf bankName = "Scb" Then
      Set resultDict = Application.Run("Scb.ExtractPdfLcScb", readPdf)
  End If

  resultDict.Add "bankName", bankName
  
  Set ExtractAnyBankLc = resultDict

End Function

Private Function PrintPdfPageRange(ByVal filePath As String, ByVal startPage As Integer, ByVal endPage As Integer) As Boolean
    ' Purpose: Prints a specified page range of a PDF file silently using Adobe Acrobat SDK.
    ' Returns: True if printing was initiated successfully, False otherwise.
    ' Notes:
    ' - Requires Adobe Acrobat (Standard or Pro, *not* just Reader) to be installed.
    ' - Uses Late Binding for better compatibility across Acrobat versions.
    ' - Assumes startPage and endPage are 1-based page numbers.

    Dim avDoc As Object
    Dim pdDoc As Object
    Dim printSuccess As Boolean
    Dim actualStartPage As Integer
    Dim actualEndPage As Integer
    Dim numPages As Long

    ' --- Initialize ---
    printSuccess = False ' Default to failure
    Set avDoc = Nothing
    Set pdDoc = Nothing

    ' --- Input Validation ---
    If Len(Dir(filePath)) = 0 Then ' Check if file exists
        MsgBox "Error: PDF file not found." & vbCrLf & filePath, vbCritical, "File Not Found"
        GoTo Cleanup ' Exit point
    End If

    If startPage < 1 Or endPage < startPage Then
        MsgBox "Error: Invalid page range specified (" & startPage & " to " & endPage & "). Start page must be >= 1 and End page must be >= Start page.", vbCritical, "Invalid Page Range"
        GoTo Cleanup ' Exit point
    End If

    ' --- Set up Error Handling ---
    On Error GoTo ErrorHandler

    ' --- Create Acrobat Application Document Object (AVDoc) ---
    ' This is usually sufficient to interact with a document view/printing
    Set avDoc = CreateObject("AcroExch.AVDoc")

    ' --- Open the PDF Document Silently ---
    If avDoc.Open(filePath, "") Then

        ' --- Get the PDDoc to check page count ---
        Set pdDoc = avDoc.GetPDDoc()
        numPages = pdDoc.GetNumPages()

        ' --- Validate page range against actual document pages ---
        If startPage > numPages Then
             MsgBox "Error: Start page (" & startPage & ") exceeds the total number of pages (" & numPages & ") in the document.", vbCritical, "Invalid Page Range"
             GoTo Cleanup ' Exit point
        End If
        If endPage > numPages Then
             MsgBox "Warning: End page (" & endPage & ") exceeds the total pages (" & numPages & "). Adjusting to print up to the last page.", vbExclamation, "Page Range Adjusted"
             endPage = numPages ' Adjust end page to the maximum valid page
        End If

        ' --- Adjust page numbers to 0-based index for Adobe SDK ---
        actualStartPage = startPage - 1
        actualEndPage = endPage - 1

        ' --- Print the specified page range silently ---
        ' Parameters for PrintPagesSilent:
        ' nFirstPage (0-based)
        ' nLastPage (0-based)
        ' nPrintLevel (2 = Level 2 PostScript with annotations, common default; 3=PrintAsImage)
        ' bShrinkToFit (0 = False)
        ' bBinaryOk (0 = False, use True (1) if printer supports binary PostScript)
        ' Return value: Non-zero on success, 0 on failure (or throws error)
        If avDoc.PrintPagesSilent(actualStartPage, actualEndPage, 2, 0, 0) <> 0 Then
            printSuccess = True ' Method returned success code
        Else
            ' Even if it returns 0, sometimes it works but might indicate a minor issue.
            ' Primarily rely on error handling, but log this potential issue if needed.
            Debug.Print "PrintPagesSilent returned 0 for file: " & filePath & ". Print may have failed silently."
            ' Keep printSuccess = False here as a precaution
        End If

    Else
        ' Failed to open the document
        MsgBox "Error: Could not open the PDF file." & vbCrLf & filePath, vbCritical, "File Open Error"
        ' avDoc.Open failed, avDoc object might still exist but is useless
        GoTo Cleanup
    End If

' --- Cleanup: Close document and release objects ---
Cleanup:
    On Error Resume Next ' Ignore errors during cleanup itself

    ' Close the document without saving changes
    If Not avDoc Is Nothing Then
        avDoc.Close True ' Use True to close silently without prompting
    End If

    ' Release COM objects
    Set pdDoc = Nothing
    Set avDoc = Nothing

    ' Optional: Suggest garbage collection (rarely needed, but can help in loops)
    ' VBA.Collection.Remove "AcroExch.AVDoc" ' This syntax isn't correct
    ' Consider Application.StatusBar updates if process is long

    On Error GoTo 0 ' Restore default error handling
    PrintPdfPageRange = printSuccess ' Return the final status
    Exit Function ' Normal exit

' --- Error Handler ---
ErrorHandler:
    MsgBox "An error occurred during PDF printing:" & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description & vbCrLf & _
           "Source: " & Err.Source, vbCritical, "Adobe Automation Error"
    printSuccess = False ' Ensure function returns False on error
    Resume Cleanup ' Go to cleanup routine to release any acquired resources

End Function

Private Function GetTimestampForFilename() As String
    ' Returns the current date and time formatted as YYYY-MM-DD_HH-MM-SS.
    ' More readable for some, still filename-safe and sortable.
    ' Example Output: 2025-04-08_10-54-21 (using nn for minutes)

    GetTimestampForFilename = Format(Now, "yyyy-mm-dd_hh-nn-ss")

End Function

Private Function GetPageRangeForPrint(extractedLcDict As Object) As Object
    
  Dim resultDict As Object
  Set resultDict = CreateObject("Scripting.Dictionary")

  Dim piArr As Variant
  piArr = Split(extractedLcDict("pi"), ",")

  Debug.Print (UBound(piArr) + 1)

  Dim bankName As String
  bankName = extractedLcDict("bankName")

  If bankName = "AlArafah" Then

    resultDict.Add "startPage", 1
    resultDict.Add "endPage", 3 + (UBound(piArr) + 1)
      
  ElseIf bankName = "Brac" Then

    resultDict.Add "startPage", 1
    resultDict.Add "endPage", 2 + (UBound(piArr) + 1)
      
  ElseIf bankName = "City" Then

    resultDict.Add "startPage", 1
    resultDict.Add "endPage", 4 + (UBound(piArr) + 1)
      
  ElseIf bankName = "Mtb" Then

    resultDict.Add "startPage", 1
    resultDict.Add "endPage", 4 + (UBound(piArr) + 1)
      
  ElseIf bankName = "Mtb1" Then

    resultDict.Add "startPage", 1
    resultDict.Add "endPage", 4 + (UBound(piArr) + 1)
      
  ElseIf bankName = "Scb" Then

    resultDict.Add "startPage", 1
    resultDict.Add "endPage", 3 + (UBound(piArr) + 1)
      
  End If

  Set GetPageRangeForPrint = resultDict

End Function