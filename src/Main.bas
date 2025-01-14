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
    
    row = row + 1
    
    For Each dicKey In resultDict.Keys
    
        Set tempDict = resultDict(dicKey)
        
        ws.Cells(row, 1).Value = tempDict("lcNo")
        ws.Cells(row, 2).Value = CDate(tempDict("lcDt"))
        ws.Cells(row, 3).Value = CDate(tempDict("expiryDt"))
        ws.Cells(row, 4).Value = tempDict("beneficiary")
        ws.Cells(row, 5).Value = tempDict("amount")
        ws.Cells(row, 6).Value = CDate(tempDict("shipmentDt"))
        ws.Cells(row, 7).Value = tempDict("pi")
        
        row = row + 1
    
    Next dicKey
    
End Sub
