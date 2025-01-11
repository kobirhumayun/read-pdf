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
