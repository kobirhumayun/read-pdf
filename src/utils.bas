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

Private Function PutB2bDataToWs(resultDict As Object, ws As Worksheet, printHeaders As Boolean, startRow As Long, startColumn As Long, printPdfProperties As Boolean) As Boolean
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
    ws.Cells(row, columns + 2).Value = "Expiry Date"
    ws.Cells(row, columns + 3).Value = "Beneficiary"
    ws.Cells(row, columns + 4).Value = "Amount"
    ws.Cells(row, columns + 5).Value = "Shipment Date"
    ws.Cells(row, columns + 6).Value = "PI"
    ws.Cells(row, columns + 7).Value = "Page Count"
    ws.Cells(row, columns + 8).Value = "Text Page Count"
    ws.Cells(row, columns + 9).Value = "Text Page List"
    ws.Cells(row, columns + 10).Value = "Blank Page Count"
    ws.Cells(row, columns + 11).Value = "Blank Page List"
    row = row + 1
  End If
  
  For Each dicKey In resultDict.Keys
  
      Set tempDict = resultDict(dicKey)
      
      ws.Cells(row, columns).Value = tempDict("lcNo")
      If tempDict("lcDt") <> "" Then
          ws.Cells(row, columns + 1).Value = IIf(IsDate(tempDict("lcDt")), CDate(tempDict("lcDt")), tempDict("lcDt"))
      End If
      If tempDict("expiryDt") <> "" Then
          ws.Cells(row, columns + 2).Value = IIf(IsDate(tempDict("expiryDt")), CDate(tempDict("expiryDt")), tempDict("expiryDt"))
      End If
      ws.Cells(row, columns + 3).Value = tempDict("beneficiary")
      ws.Cells(row, columns + 4).Value = tempDict("amount")
      If tempDict("shipmentDt") <> "" Then
          ws.Cells(row, columns + 5).Value = IIf(IsDate(tempDict("shipmentDt")), CDate(tempDict("shipmentDt")), tempDict("shipmentDt"))
      End If
      ws.Cells(row, columns + 6).Value = tempDict("pi")

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

  Set ExtractAnyBankLc = resultDict

End Function
