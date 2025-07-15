'--------------------------------------------------------------
' Module: ProductionPlanAutomation
' Description: Automates end-to-end production plan generation
'              using Power Query for data transformations and a VBA
'              macro to assemble the final workbook.
'--------------------------------------------------------------
Option Explicit

Sub BuildProductionPlan()
    ' 1. Refresh all Power Query connections
    Dim conn As WorkbookConnection
    For Each conn In ThisWorkbook.Connections
        If conn.Type = xlConnectionTypeWORKSHEET Then
            ' Skip worksheet connections
        Else
            conn.Refresh
        End If
    Next conn
    
    ' 2. Load transformed tables into staging sheets
    Dim wsOrders As Worksheet, wsMaster As Worksheet, wsStocks As Worksheet
    On Error Resume Next
    Set wsOrders = ThisWorkbook.Worksheets("PQ_Orders")
    Set wsMaster = ThisWorkbook.Worksheets("PQ_MasterData")
    Set wsStocks = ThisWorkbook.Worksheets("PQ_Stocks")
    On Error GoTo 0
    
    ' 3. Compute net production requirements (handled by Power Query into "PQ_NetReq")
    
    ' 4. Distribute quantities across lines
    Call DistributeAcrossLines
    
    ' 5. Generate per-line worksheets
    Call GenerateLineSheets
    
    ' 6. Update raw material daily requirement sheet
    Call UpdateRawMaterialSheet
    
    ' 7. (Optional) Update Storage (Equaliser) sheet if needed
    '    Uncomment and implement UpdateStorageSheet for Storage department
    'Call UpdateStorageSheet
    
    MsgBox "Production plan built successfully!", vbInformation
End Sub

Private Sub DistributeAcrossLines()
    ' Example stub: customize allocation logic as needed
    Dim wsNet As Worksheet, wsAlloc As Worksheet
    Set wsNet = ThisWorkbook.Worksheets("PQ_NetReq")
    On Error Resume Next
    Set wsAlloc = ThisWorkbook.Worksheets("Allocations")
    If wsAlloc Is Nothing Then
        Set wsAlloc = ThisWorkbook.Worksheets.Add(After:=wsNet)
        wsAlloc.Name = "Allocations"
    End If
    On Error GoTo 0
    
    wsAlloc.Cells.Clear
    wsNet.Rows(1).Copy Destination:=wsAlloc.Rows(1)
    wsNet.Range("A2").CurrentRegion.Resize(, wsNet.Cells(1, wsNet.Columns.Count).End(xlToLeft).Column).Copy _
        Destination:=wsAlloc.Rows(2)
End Sub

Private Sub GenerateLineSheets()
    ' Loop through unique Line names in Allocations and export to separate sheets
    Dim wsAlloc As Worksheet: Set wsAlloc = ThisWorkbook.Worksheets("Allocations")
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim lastRow As Long, i As Long, lineName As String
    lastRow = wsAlloc.Cells(wsAlloc.Rows.Count, "A").End(xlUp).Row
    For i = 2 To lastRow
        lineName = wsAlloc.Cells(i, 2).Value
        If Not dict.Exists(lineName) Then dict.Add lineName, True
    Next i
    Dim key As Variant, wsLine As Worksheet
    For Each key In dict.Keys
        On Error Resume Next
        Set wsLine = ThisWorkbook.Worksheets(key)
        If wsLine Is Nothing Then
            Set wsLine = ThisWorkbook.Worksheets.Add(After:=Sheets(Sheets.Count))
            wsLine.Name = key
        Else
            wsLine.Cells.Clear
        End If
        On Error GoTo 0
        wsAlloc.Rows(1).Copy Destination:=wsLine.Rows(1)
        wsAlloc.Range("A1").CurrentRegion.AutoFilter Field:=2, Criteria1:=key
        wsAlloc.Range("A1").CurrentRegion.Offset(1, 0).SpecialCells(xlCellTypeVisible).Copy Destination:=wsLine.Rows(2)
        wsLine.Cells.AutoFit
    Next key
    wsAlloc.AutoFilterMode = False
End Sub

Private Sub UpdateRawMaterialSheet()
    ' Copy daily raw material requirements to "Raw Material Daily Requirement" sheet
    Dim wsPQRM As Worksheet, wsRM As Worksheet
    Set wsPQRM = ThisWorkbook.Worksheets("PQ_RawMaterials")
    Set wsRM = ThisWorkbook.Worksheets("Raw Material Daily Requirement")
    wsRM.Cells.Clear
    wsPQRM.UsedRange.Copy Destination:=wsRM.Range("A1")
    wsRM.Cells.AutoFit
End Sub

'Optional: Storage department (Equaliser) update routine
'Private Sub UpdateStorageSheet()
'    Dim wsPQStore As Worksheet, wsStore As Worksheet
'    Set wsPQStore = ThisWorkbook.Worksheets("PQ_Storage")
'    Set wsStore = ThisWorkbook.Worksheets("Equaliser")
'    wsStore.Cells.Clear
'    wsPQStore.UsedRange.Copy Destination:=wsStore.Range("A1")
'    wsStore.Cells.AutoFit
'End Sub
