''
' Help to compute the variance from first count shop and Inventory
' On Hand data
'
' The main function here is Main (the entry point)
'



' Information related to InventoryOnHand file layout
'
Type InventoryOnHandInfo
    name As String
    codeColID As Integer
    internalColId As Integer
    desColID As Integer
    priceColID As Integer
    valueColID As Integer
    qtyColId As Integer
End Type

' Information related to first cunt shop file layout
'
Type firstCountShopInfo
    name As String
    codeColID As Integer
    descColID As Integer
    qtyColId As Integer
End Type

' Information related to first count shop file layout
'
Type VarianceReportInfo
    name As String
    
    codeColID As Integer
    'internalColId As Integer
    desColID As Integer
    itemColId As Integer
    'priceColID As Integer
    'valueColID As Integer
    qtyColId As Integer
    invOnShopColId As Integer
    varianceColId As Integer
    reCountColId As Integer
End Type

' Information related to master file layout
'
Type MasterFileInfo
    name As String
    codeColID As Integer
    desColID As Integer
    priceColID As Integer
    qtyColId As Integer
    locationColId As Integer
End Type


Public iohInfo As InventoryOnHandInfo
Public fcsInfo As firstCountShopInfo
Public vrInfo As VarianceReportInfo
Public mfInfo As MasterFileInfo


' Intialize global variable info
'
' This Mapping helps to refactor easly in case file format changes.
Sub InitizaliteInfos()
    iohInfo.name = "InventoryOnHand"
    iohInfo.codeColID = 1
    iohInfo.internalColId = 2
    iohInfo.desColID = 3
    iohInfo.priceColID = 6
    iohInfo.valueColID = 7
    iohInfo.qtyColId = 8
    
    fcsInfo.name = "FirstCountShop"
    fcsInfo.codeColID = 2
    fcsInfo.descColID = 3
    fcsInfo.qtyColId = 5
    
    vrInfo.name = "VarianceReport"
    vrInfo.codeColID = 1
    vrInfo.desColID = 2
    vrInfo.itemColId = 3
    vrInfo.qtyColId = 4
    vrInfo.invOnShopColId = 5
    vrInfo.varianceColId = 6
    vrInfo.reCountColId = 7

    
    mfInfo.name = "MasterFile"
    mfInfo.codeColID = 1
    mfInfo.desColID = 2
    mfInfo.priceColID = 3
    mfInfo.qtyColId = 4
    mfInfo.locationColId = 5
End Sub

' Entry point
' Select this function as the Macro entry point
'
Sub Main()
    InitizaliteInfos
    Application.ScreenUpdating = False
    
    If SheetExists(vrInfo.name) Then
        If MsgBox("Delete current report and create a new one?", vbYesNo, "Confirm") = vbYes Then
            Reset
        Else ' user selected No, do nothing
            Exit Sub
        End If
    End If

    ' Step 1:
    '  import FirstCountFile and InventoryOnHand
    If Import = False Then
        Debug.Print "Import Failed"
        Exit Sub
    End If
    
    Debug.Print "Import OK"

    ' Step 2:
    SanitizeAll
    
    Debug.Print "Sanitize OK"
    
    ' Step 3:
    ' Compute variance
    BuildVarianceReport
    Application.ScreenUpdating = True
End Sub


' Import selected files to activeworkbook
'
Function Import() As Boolean
    Dim ret As Boolean
    
    ret = ImportSheetFromTextFile(fcsInfo.name, "Select Shop file")

    If ret = False Then
        ' import of first count shop failed
        Import = False
        Exit Function
    End If
    
    ret = ImportSheetFromFile(1, iohInfo.name, "Select InventoryOnHand")

    If ret = False Then
        ' import of InventoryOnHand failed
        Import = False
        Exit Function
    End If
    
    
    ret = ImportSheetFromFile(1, mfInfo.name, "Select Master file")

    If ret = False Then
        ' import of master file failed
        Import = False
        Exit Function
    End If
    
    Import = True
End Function


' Todo : Add a description
'
Function ImportSheetFromFile(ImportSheetId As Integer, name As String, caption As String) As Boolean
    Dim wb, activeWorkbook As Workbook
    Dim filter As String
    Dim selectedFilename As Variant
    
    Set activeWorkbook = Application.activeWorkbook
    filter = "Excel files (*.xlsx),*.xlsx"
    selectedFilename = Application.GetOpenFilename(filter, , caption)
    
    If selectedFilename = False Then
        ' No file selected
        ImportSheetFromFile = False
        Exit Function
    End If
    
    ' open workbook of selected file
    Set wb = Workbooks.Open(selectedFilename)
    wb.Sheets(ImportSheetId).Move After:=activeWorkbook.Sheets(activeWorkbook.Sheets.count)
    
    ActiveSheet.name = name
    ImportSheetFromFile = True
End Function


' Todo : Add a description
'
Sub SanitizeAll()
    Sheets("Sheet1").name = vrInfo.name
    SanitizeInventoryInHand
    SanitizeFirstCountShop
    SanitizeMasterFile
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
    Sheets(iohInfo.name).Range("1:6").Delete
End Sub


' Remove the swapped column
'
Sub SanitizeMasterFile()
    Dim mfWS As Worksheet
    
    Set mfWS = Sheets(mfInfo.name)
    
    For r = mfWS.UsedRange.Rows.count To 2 Step -1
        If Not IsNumeric(mfWS.Cells(r, mfInfo.codeColID)) Then
            mfWS.Rows(r).EntireRow.Delete
        End If
    Next
End Sub


' Todo : Add a description
'
Sub SanitizeFirstCountShop()
    Dim firstShopWS As Worksheet
    Dim codeColID, qtyColId As Integer
    Dim items As Collection
    Dim code As Long
    Dim codeKey As String

    codeColID = fcsInfo.codeColID
    qtyColId = fcsInfo.qtyColId

    Set firstShopWS = Sheets(fcsInfo.name)
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


' Creates a collection that contains every information required to
' sanitize the original file
'
Function GetItemsCollection() As Collection
    Dim firstShopWS As Worksheet
    Dim items As Collection
    Dim codeColID, descColID, qtyColId As Integer
    Dim code, qty As Long
    Dim codeKey As String
    
    codeColID = fcsInfo.codeColID
    descColID = fcsInfo.descColID
    qtyColId = fcsInfo.qtyColId

    Set items = New Collection
    Set firstShopWS = Sheets(fcsInfo.name)
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


' Todo : Add a description
'
Public Function Contains(col As Collection, key As Variant) As Boolean
    On Error Resume Next
    col (key) ' Just try it. If it fails, Err.Number will be nonzero.
    Contains = (Err.Number = 0)
    Err.Clear
End Function


' import data from InventoryOnHand to VarianceReport
'
Sub BuildVarianceReport()
    Dim masterWS, vReportWS As Worksheet
    
    Set masterWS = Sheets(mfInfo.name)
    Set vReportWS = Sheets(vrInfo.name)
    ' Copy selected columns to VarianceReport Sheet
    
        
    masterWS.Columns(mfInfo.codeColID).Copy Destination:=vReportWS.Columns(vrInfo.codeColID)
    masterWS.Columns(mfInfo.desColID).Copy Destination:=vReportWS.Columns(vrInfo.desColID)

    
    vReportWS.Cells(1, vrInfo.itemColId) = "Item"
    vReportWS.Cells(1, vrInfo.qtyColId) = "NetSuite"
    vReportWS.Cells(1, vrInfo.invOnShopColId) = "Count"
    vReportWS.Cells(1, vrInfo.varianceColId) = "Variance"
    vReportWS.Cells(1, vrInfo.reCountColId) = "ReCount"
    
    ImportItemsFromInventoryOnHand
    importNetSuiteFromInventoryOnHand
    ImportCountFromFirstCountShop
    ComputeVarianceValue


    ' Apply filter
    vReportWS.AutoFilterMode = False
    vReportWS.Range("A:G").AutoFilter Field:=6, Criteria1:="<>0", VisibleDropDown:=True

    ' Now, resize columns to AutoFit size
    For C = 1 To vReportWS.UsedRange.Columns.count
        vReportWS.Columns(C).AutoFit
    Next

    With vReportWS.Range(vReportWS.Cells(2, vrInfo.varianceColId), vReportWS.Cells(2, vrInfo.varianceColId).End(xlDown))
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlNotEqual, Formula1:="=0"
        .FormatConditions(1).Interior.ColorIndex = 6
    End With

    vReportWS.Activate
End Sub


Sub ImportItemsFromInventoryOnHand()
    Dim iohWS, vReportWS As Worksheet
    Dim nbRows As Integer
    
    Set iohWS = Sheets(iohInfo.name)
    Set vReportWS = Sheets(vrInfo.name)
    
    nbRows = iohWS.UsedRange.Rows.count + 50  'add margin
    lookupFormula = "VLOOKUP(RC[-2]," & iohWS.name & "!R2C1:R" & nbRows & "C5, 5, False)"
    errorMsg = "Not Found"
    For r = 2 To vReportWS.UsedRange.Rows.count
        form = "=IFERROR(" & lookupFormula & "," & Chr(34) & errorMsg & Chr(34) & ")"
        vReportWS.Cells(r, vrInfo.itemColId).FormulaR1C1 = form
    Next
End Sub


Sub importNetSuiteFromInventoryOnHand()
    Dim iohWS, vReportWS As Worksheet
    Dim nbRows As Integer
    
    Set iohWS = Sheets(iohInfo.name)
    Set vReportWS = Sheets(vrInfo.name)
    
    nbRows = iohWS.UsedRange.Rows.count + 50  'add margin
    lookupFormula = "VLOOKUP(RC[-3]," & iohWS.name & "!R2C1:R" & nbRows & "C8, 8, False)"
    errorMsg = "Not Found"
    For r = 2 To vReportWS.UsedRange.Rows.count
        form = "=IFERROR(" & lookupFormula & "," & Chr(34) & errorMsg & Chr(34) & ")"
        vReportWS.Cells(r, vrInfo.qtyColId).FormulaR1C1 = form
    Next
End Sub



' Create the VLookup to retrieve information from the first count shop
'
Sub ImportCountFromFirstCountShop()
    Dim firstShopWS, vReportWS As Worksheet
    Dim nbRowsInFsc As Integer
    
    Set firstShopWS = Sheets(fcsInfo.name)
    Set vReportWS = Sheets(vrInfo.name)
    
    nbRowsInFsc = firstShopWS.UsedRange.Rows.count + 50  'add margin
    lookupFormula = "VLOOKUP(RC[-4]," & fcsInfo.name & "!R2C2:R" & nbRowsInFsc & "C5, 4, False)"
    errorMsg = "0"
    For r = 2 To vReportWS.UsedRange.Rows.count
        form = "=IFERROR(" & lookupFormula & "," & errorMsg & ")"
        vReportWS.Cells(r, vrInfo.invOnShopColId).FormulaR1C1 = form
    Next
End Sub


' Todo : Add a description
'
Sub ComputeVarianceValue()
    Dim vReportWS As Worksheet
    
    Set vReportWS = Sheets(vrInfo.name)
    
    For r = 2 To vReportWS.UsedRange.Rows.count
        vReportWS.Cells(r, vrInfo.varianceColId).FormulaR1C1 = "=(RC[-1] - RC[-2])"
    Next
End Sub


' Reset the file to its original point
' Clear everything, leave a "Sheet1"
'
Sub Reset()
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets(fcsInfo.name).Delete
    ThisWorkbook.Sheets(iohInfo.name).Delete
    ThisWorkbook.Sheets(mfInfo.name).Delete
    ThisWorkbook.Sheets(vrInfo.name).AutoFilterMode = False
    ThisWorkbook.Sheets(vrInfo.name).UsedRange.ClearContents
    ThisWorkbook.Sheets(vrInfo.name).name = "Sheet1"
    Application.DisplayAlerts = True
    On Error GoTo 0
End Sub


' Check if a sheet exists according to its name
'
Function SheetExists(name As String) As Boolean
  SheetExists = False
  For Each WS In Worksheets
    If name = WS.name Then
      SheetExists = True
      Exit Function
    End If
  Next WS
End Function



Sub test()
    InitizaliteInfos
    ImportFirstCountTextFile
End Sub


Function ImportSheetFromTextFile(name As String, caption As String) As Boolean
    Dim i, j As Integer
    Dim fd As Integer: fd = FreeFile
    Dim filter As String
    Dim selectedFilename As Variant
    Dim lines() As String
    
    filter = "Text Files (*.txt),*.txt"
    selectedFilename = Application.GetOpenFilename(filter, , caption)
    
    If selectedFilename = False Then
        ' No file selected
        ImportFirstCountTextFile = False
        Exit Function
    End If
    
    Dim newWS As Worksheet
    Set newWS = Sheets.Add
    newWS.name = name
    
    
    Open selectedFilename For Input As #fd
        lines = Split(Input$(LOF(fd), #fd), vbNewLine)
    Close #fd
    
    For i = 0 To UBound(lines)
        Dim arr
        arr = Split(lines(i), vbTab)
        For j = 0 To UBound(arr)
            newWS.Cells(i + 1, j + 1).Value = arr(j)
        Next
    Next
    
    ImportSheetFromTextFile = True
End Function
