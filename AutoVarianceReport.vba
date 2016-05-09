Sub Main()


    If SheetExists("VarianceReport") Then
        If MsgBox("Delete current report and create a new one?", vbYesNo, "Confirm") = vbYes Then
            Reset
        Else 'user selected No, do nothing
            Exit Sub
        End If
    End If

    ' Step 1:
    '  import FirstCountFile and InventoryOnHand
    If Import = False Then
        Exit Sub
    End If
    
    Debug.Print "Import OK"

    ' Step 2:
    SanitizeAll
    
    Debug.Print "Sanitize OK"
    
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



Function ImportSheetFromFile(ImportSheetId As Integer, name As String, caption As String) As Boolean
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
    
    ActiveSheet.name = name
    ImportSheetFromFile = True
End Function


Sub SanitizeAll()
    Sheets("Sheet1").name = "VarianceReport"
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



' Todo : Add a general descriptop,
'
'
'
Sub SanitizeFirstCountShop()
    Dim firstShopWS As Worksheet
    Dim codeColID, qtyColId As Integer
    Dim items As Collection
    Dim code As Long
    Dim codeKey As String

    codeColID = 2
    qtyColId = 5

    Set firstShopWS = Sheets("FirstCountShop")
    Set items = GetItemsCollection

    For r = firstShopWS.UsedRange.Rows.count To 2 Step -1
        code = firstShopWS.Cells(r, codeColID)
        codeKey = CStr(code)
        If Contains(items, codeKey) Then
            firstShopWS.Cells(r, qtyColId).Value = items(codeKey)
            items.Remove codeKey
        Else
            ' We already processed this guy, removing this line
            firstShopWS.Rows(r).EntireRow.Delete
        End If
    Next
End Sub


' Todo : Add a general description
'
'
'
Function GetItemsCollection() As Collection
    Dim firstShopWS As Worksheet
    Dim items As Collection
    Dim codeColID, descColID, qtyColId As Integer
    Dim code, qty As Long
    Dim codeKey As String
    
    codeColID = 2
    descColID = 3
    qtyColId = 5

    Set items = New Collection
    Set firstShopWS = Sheets("FirstCountShop")
    For r = 2 To firstShopWS.UsedRange.Rows.count
        If Not IsNumeric(firstShopWS.Cells(r, codeColID)) Then
            ' Code is at desc position!
            code = firstShopWS.Cells(r, descColID).Value
            ' Swapping
            firstShopWS.Cells(r, descColID).Value = firstShopWS.Cells(r, codeColID).Value
            firstShopWS.Cells(r, codeColID).Value = code
        Else
            code = firstShopWS.Cells(r, codeColID).Value
        End If
        codeKey = CStr(code)
        qty = firstShopWS.Cells(r, qtyColId).Value

        If Contains(items, codeKey) Then
            ' Update qty for this code
            Dim tmp As Long
            tmp = items(codeKey)
            items.Remove (codeKey)
            items.Add qty + tmp, codeKey
        Else
            ' new item adding it to the collection
            items.Add qty, codeKey
        End If
    Next
    Set GetItemsCollection = items
End Function


 
Public Function Contains(col As Collection, key As Variant) As Boolean
    On Error Resume Next
    col (key) ' Just try it. If it fails, Err.Number will be nonzero.
    Contains = (Err.Number = 0)
    Err.Clear
End Function

'import data from InventoryOnHand to VarianceReport
'
'UPC code
'display name
'price
'value
'qty
Sub BuildVarianceReport()
    Dim codeColID, internalColId, descColID, priceColID, valueColID, qtyColId As String
    Dim inventoryWS, vReportWS As Worksheet
   
    codeColID = "A"
    internalColId = "B"
    desColID = "C"
    priceColID = "F"
    valueColID = "G"
    qtyColId = "H"
    
    Set inventoryWS = Sheets("InventoryOnHand")
    Set vReportWS = Sheets("VarianceReport")
    'Copy selected columns to VarianceReport Sheet
    inventoryWS.Columns(codeColID).Copy Destination:=vReportWS.Columns("A")
    inventoryWS.Columns(internalColId).Copy Destination:=vReportWS.Columns("B")
    inventoryWS.Columns(desColID).Copy Destination:=vReportWS.Columns("C")
    inventoryWS.Columns(priceColID).Copy Destination:=vReportWS.Columns("D")
    inventoryWS.Columns(valueColID).Copy Destination:=vReportWS.Columns("E")
    inventoryWS.Columns(qtyColId).Copy Destination:=vReportWS.Columns("F")
    vReportWS.Range("G1").Value = "Inv On Shop"
    vReportWS.Range("H1").Value = "Variance"
    
    CreateVLookup
    ComputeVarianceValue
    
    
    ' Apply filter
    vReportWS.AutoFilterMode = False
    vReportWS.Range("A:H").AutoFilter Field:=8, Criteria1:="<>0", VisibleDropDown:=True
    
    ' Now, resize columns to AutoFit size
    For c = 1 To vReportWS.UsedRange.Columns.count
        vReportWS.Columns(c).AutoFit
    Next
    
    ' Apply Conditional Color
    With vReportWS.Range(vReportWS.Range("H2"), vReportWS.Range("H2").End(xlDown))
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlNotEqual, Formula1:="=0"
        .FormatConditions(1).Interior.ColorIndex = 6
    End With
    
    vReportWS.Activate
End Sub


Sub CreateVLookup()
    Dim firstShopWS, vReportWS As Worksheet
    
    Set firstShopWS = Sheets("FirstCountShop")
    Set vReportWS = Sheets("VarianceReport")
    
    lookupFormula = "VLOOKUP(RC[-6], 'FirstCountShop'!R2C2:R500C5, 4, False)"
    errorMsg = "Not Found"
    For r = 2 To vReportWS.UsedRange.Rows.count
        form = "=IFERROR(" & lookupFormula & "," & Chr(34) & errorMsg & Chr(34) & ")"
        vReportWS.Cells(r, 7).FormulaR1C1 = form
    Next
End Sub


Sub ComputeVarianceValue()
    Dim vReportWS As Worksheet
    
    Set vReportWS = Sheets("VarianceReport")
    
    For r = 2 To vReportWS.UsedRange.Rows.count
        vReportWS.Cells(r, 8).FormulaR1C1 = "=(RC[-1] - RC[-2])"
    Next
End Sub



Sub Reset()
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("FirstCountShop").Delete
    ThisWorkbook.Sheets("InventoryOnHand").Delete
    ThisWorkbook.Sheets("VarianceReport").AutoFilterMode = False
    ThisWorkbook.Sheets("VarianceReport").UsedRange.ClearContents
    ThisWorkbook.Sheets("VarianceReport").name = "Sheet1"
    Application.DisplayAlerts = True
    On Error GoTo 0
End Sub


Function SheetExists(name As String) As Boolean
  SheetExists = False
  For Each WS In Worksheets
    If name = WS.name Then
      SheetExists = True
      Exit Function
    End If
  Next WS
End Function
