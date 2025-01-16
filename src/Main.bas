Attribute VB_Name = "Main"
Option Explicit

Sub BracLc()

    Dim b2bPaths As Object
    Set b2bPaths = Application.Run("utils.GetSelectedFilePaths", "G:\PDL Customs\Export LC, Import LC & UP\Import LC With Related Doc\YEAR-2025", "Select Brac LC Only", "pdf")
    
    Dim resultDict As Object
    Set resultDict = Application.Run("Brac.ReadBracLcs", b2bPaths)

    Dim tempDict As Object
    
    Dim dicKey As Variant
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.ActiveSheet
    
    Dim row As Long
    row = 1
    
    ' Print headers
    ws.Cells(row, 1).Value = "LC No"
    ws.Cells(row, 2).Value = "LC Date"
    ws.Cells(row, 3).Value = "Expiry Date"
    ws.Cells(row, 4).Value = "Beneficiary"
    ws.Cells(row, 5).Value = "Amount"
    ws.Cells(row, 6).Value = "Shipment Date"
    ws.Cells(row, 7).Value = "PI"

    ws.Cells(row, 8).Value = "Page Count"
    ws.Cells(row, 9).Value = "Text Page Count"
    ws.Cells(row, 10).Value = "Text Page List"
    ws.Cells(row, 11).Value = "Blank Page Count"
    ws.Cells(row, 12).Value = "Blank Page List"

    row = row + 1
    
    For Each dicKey In resultDict.Keys
    
        Set tempDict = resultDict(dicKey)
        
        ws.Cells(row, 1).Value = tempDict("lcNo")
        If tempDict("lcDt") <> "" Then
            ws.Cells(row, 2).Value = IIf(IsDate(tempDict("lcDt")), CDate(tempDict("lcDt")), tempDict("lcDt"))
        End If
        If tempDict("expiryDt") <> "" Then
            ws.Cells(row, 3).Value = IIf(IsDate(tempDict("expiryDt")), CDate(tempDict("expiryDt")), tempDict("expiryDt"))
        End If
        ws.Cells(row, 4).Value = tempDict("beneficiary")
        ws.Cells(row, 5).Value = tempDict("amount")
        If tempDict("shipmentDt") <> "" Then
            ws.Cells(row, 6).Value = IIf(IsDate(tempDict("shipmentDt")), CDate(tempDict("shipmentDt")), tempDict("shipmentDt"))
        End If
        ws.Cells(row, 7).Value = tempDict("pi")

            ' just for testing
        ws.Cells(row, 8).value = tempDict("pdfProperties")("totalPageCount")
        ws.Cells(row, 9).value = tempDict("pdfProperties")("textPagesCount")
        ws.Cells(row, 10).value = tempDict("pdfProperties")("textPagesList")
        ws.Cells(row, 11).value = tempDict("pdfProperties")("blankPagesCount")
        ws.Cells(row, 12).value = tempDict("pdfProperties")("blankPagesList")

        row = row + 1
    
    Next dicKey
    
End Sub
