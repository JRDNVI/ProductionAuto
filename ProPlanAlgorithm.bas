Attribute VB_Name = "ProductionPlanAlgorithm"
Option Explicit

'--------------------------------------------------------------
' Module: ProductionPlanAlgorithm
' Description: Main VBA module implementing the production plan
'              algorithm driven by parameters, Power Query staging, and
'              allocation logic. Export this as a .bas file.
'--------------------------------------------------------------

'========================================
' Constants for table names
'========================================
Private Const PARAM_TABLE As String = "tblParameters"
Private Const ORDERS_SHEET As String = "PQ_Orders"
Private Const MASTER_SHEET As String = "PQ_MasterData"
Private Const STOCKS_SHEET As String = "PQ_Stocks"
Private Const NETREQ_SHEET As String = "PQ_NetReq"
Private Const RAWMAT_SHEET As String = "PQ_RawMaterials"
Private Const STORAGE_SHEET As String = "PQ_Storage"
Private Const ALLOC_SHEET As String = "Allocations"
Private Const RAWMAT_OUT_SHEET As String = "Raw Material Daily Requirement"
Private Const STORAGE_OUT_SHEET As String = "Equaliser"

'========================================
' Entry point: refresh PQ, read params, run subroutines
'========================================
Public Sub BuildProductionPlan()
    Dim params As Dictionary
    Set params = ReadParameters()
    
    RefreshAllQueries
    
    Call DistributeAcrossLines(params)
    Call GenerateLineSheets
    Call UpdateRawMaterialSheet
    Call UpdateStorageSheet
    
    MsgBox "Production plan built successfully!", vbInformation
End Sub

'========================================
' Refresh all Power Query connections
'========================================
Private Sub RefreshAllQueries()
    Dim conn As WorkbookConnection
    For Each conn In ThisWorkbook.Connections
        On Error Resume Next
        conn.Refresh
        On Error GoTo 0
    Next conn
End Sub

'========================================
' Read thresholds and settings from Parameters table
'========================================
Private Function ReadParameters() As Dictionary
    Dim dict As Dictionary: Set dict = New Dictionary
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Parameters")
    Dim tbl As ListObject
    Set tbl = ws.ListObjects(PARAM_TABLE)
    Dim i As Long, key As String, val As Variant
    With tbl.DataBodyRange
        For i = 1 To .Rows.Count
            key = .Cells(i, tbl.ListColumns("Parameter").Index).Value
            val = .Cells(i, tbl.ListColumns("Value").Index).Value
            dict(key) = val
        Next i
    End With
    Set ReadParameters = dict
End Function

'========================================
' Distribute net requirements across lines
'========================================
Private Sub DistributeAcrossLines(params As Dictionary)
    Dim wsNet As Worksheet, wsAlloc As Worksheet
    Set wsNet = Sheets(NETREQ_SHEET)
    On Error Resume Next
    Set wsAlloc = Sheets(ALLOC_SHEET)
    If wsAlloc Is Nothing Then
        Set wsAlloc = Sheets.Add(After:=wsNet)
        wsAlloc.Name = ALLOC_SHEET
    Else
        wsAlloc.Cells.Clear
    End If
    On Error GoTo 0
    
    ' Copy headers
    wsNet.Rows(1).Copy wsAlloc.Rows(1)
    
    ' Example: simple copy. Insert capacity logic below.
    wsNet.Range("A2").CurrentRegion.Copy wsAlloc.Range("A2")
    
    ' TODO: apply params("MaxLineCapacityPerShift"), params("MinProductionLot"), etc.
End Sub

'========================================
' Generate per-line worksheets from Allocations
'========================================
Private Sub GenerateLineSheets()
    Dim wsAlloc As Worksheet: Set wsAlloc = Sheets(ALLOC_SHEET)
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim lastRow As Long, i As Long, lineName As String
    lastRow = wsAlloc.Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow
        lineName = wsAlloc.Cells(i, 2).Value
        If Not dict.Exists(lineName) Then dict(lineName) = True
    Next i
    Dim key As Variant, wsLine As Worksheet
    For Each key In dict.Keys
        On Error Resume Next
        Set wsLine = Sheets(key)
        If wsLine Is Nothing Then
            Set wsLine = Sheets.Add(After:=Sheets(Sheets.Count))
            wsLine.Name = key
        Else
            wsLine.Cells.Clear
        End If
        On Error GoTo 0
        wsAlloc.Rows(1).Copy wsLine.Rows(1)
        wsAlloc.Range("A1").CurrentRegion.AutoFilter Field:=2, Criteria1:=key
        wsAlloc.Range("A1").CurrentRegion.Offset(1).SpecialCells(xlCellTypeVisible).Copy wsLine.Rows(2)
        wsLine.Cells.AutoFit
    Next key
    wsAlloc.AutoFilterMode = False
End Sub

'========================================
' Update Raw Material Daily Requirement
'========================================
Private Sub UpdateRawMaterialSheet()
    Dim wsPQ As Worksheet: Set wsPQ = Sheets(RAWMAT_SHEET)
    Dim wsOut As Worksheet: Set wsOut = Sheets(RAWMAT_OUT_SHEET)
    wsOut.Cells.Clear
    wsPQ.UsedRange.Copy wsOut.Range("A1")
    wsOut.Cells.AutoFit
End Sub

'========================================
' Update Storage (Equaliser) sheet
'========================================
Private Sub UpdateStorageSheet()
    Dim wsPQ As Worksheet, wsOut As Worksheet
    On Error Resume Next
    Set wsPQ = Sheets(STORAGE_SHEET)
    Set wsOut = Sheets(STORAGE_OUT_SHEET)
    If wsPQ Is Nothing Or wsOut Is Nothing Then Exit Sub
    wsOut.Cells.Clear
    wsPQ.UsedRange.Copy wsOut.Range("A1")
    wsOut.Cells.AutoFit
    On Error GoTo 0
End Sub
