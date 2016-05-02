Sub Main()
    ' Step 1:
    '  import FirstCountFile and InventoryOnHand
    If Import = False Then
        Exit Sub
    End If

    ' Step 2:
    SanitizeAll
    
    ' Step 3:
    ' Compute variance
    BuildVarianceReport
End Sub

'Import selected files to activeworkbook
Function Import() As Boolean
    Dim ret As Boolean
    
    ret = ImportSheetFromFile(1, "FirstCountShop", "Select Shop file")

    If ret = False Then
        'import of first count shop failed
        Import = False
        Exit Function
    End If
    
    ret = ImportSheetFromFile(1, "InventoryOnHand", "Select InventoryOnHand")

    If ret = False Then
        'import of InventoryOnHand failed
        Import = False
        Exit Function
    End If

    Import = True
End Function



Function ImportSheetFromFile(ImportSheetId As Integer, Name As String, caption As String) As Boolean
    Dim wb, activeWorkbook As Workbook
    Dim filter As String
    Dim selectedFilename As Variant
    
    Set activeWorkbook = Application.activeWorkbook
    filter = "Excel files (*.xlsx),*.xlsx"
    selectedFilename = Application.GetOpenFilename(filter, , caption)
    
    If selectedFilename = False Then
        'No file selected
        ImportSheetFromFile = False
        Exit Function
    End If
    
    'open workbook of selected file
    Set wb = Workbooks.Open(selectedFilename)
    wb.Sheets(ImportSheetId).Move After:=activeWorkbook.Sheets(activeWorkbook.Sheets.count)
    
    ActiveSheet.Name = Name
    ImportSheetFromFile = True
End Function


Sub SanitizeAll()
    Sheets("Sheet1").Name = "VarianceReport"
    SanitizeInventoryInHand
    SanitizeFirstCountShop
End Sub


' delete useless lines 1 to 6
' Expecting the file to start with following lines :
'
' ==========================================================
' Kit and Ace
' Kit and Ace Holdings Inc. (CAD) : Kit and Ace Operating US
' Inventory Qty On Hand with AVG Cost by Location
' As of April 27, 2016
'
' Options: Show Zeros
' ==========================================================
Sub SanitizeInventoryInHand()
    Sheets("InventoryOnHand").Range("1:6").Delete
End Sub


'Swap cells content if content is not numeric for a given column
Sub SanitizeFirstCountShop()
    Dim i, lastLine As Long
    Dim upc, desc As Long
    Dim firstCountShopWS As Worksheet
    
    Set firstCountShopWS = Sheets("FirstCountShop")
    
    desc = 3
    upc = 2
    
    lastLine = NbOfLinesInCol(desc)
    
    For i = 1 To lastLine
        If Not IsNumeric(firstCountShopWS.Cells(i, upc)) Then
            ' Swap cells content
            Dim temp
            temp = firstCountShopWS.Cells(i, desc).Value
            firstCountShopWS.Cells(i, desc) = firstCountShopWS.Cells(i, upc)
            firstCountShopWS.Cells(i, upc) = temp
        End If
    Next
End Sub


'Retrieve the letter associated to a column number
'
'col : the colunm number e.g 2
'return : the letter associated to the column number e.g "B"
Function ColLetter(col As Long) As String
    Dim vArr
    vArr = Split(Cells(1, col).Address(True, False), "$")
    ColLetter = vArr(0)
End Function


'Retrieve last line number in a given column number
'
'col : the colunm number e.g 2
'return : last "non empty" line number
Function NbOfLinesInCol(col As Long) As Long
    Dim descColLetter As String
    Dim lastLine As Long
    
    descColLetter = ColLetter(col)
    lastLine = Range(descColLetter & Rows.count).End(xlUp).Row
    NbOfLinesInCol = lastLine
End Function


'import data from InventoryOnHand to VarianceReport
'
'UPC code
'display name
'price
'value
'qty
Sub BuildVarianceReport()
    Dim codeColID, descColID, priceColID, valueColID, qtyColID As String
    Dim inventoryWS, vReportWS As Worksheet
   
    codeColID = "B"
    desColID = "C"
    priceColID = "F"
    valueColID = "G"
    qtyColID = "H"
    
    Set inventoryWS = Sheets("InventoryOnHand")
    Set vReportWS = Sheets("VarianceReport")
    'Copy selected columns to VarianceReport Sheet
    inventoryWS.Columns(codeColID).Copy Destination:=vReportWS.Columns("A")
    inventoryWS.Columns(desColID).Copy Destination:=vReportWS.Columns("B")
    inventoryWS.Columns(priceColID).Copy Destination:=vReportWS.Columns("C")
    inventoryWS.Columns(valueColID).Copy Destination:=vReportWS.Columns("D")
    inventoryWS.Columns(qtyColID).Copy Destination:=vReportWS.Columns("E")
End Sub
