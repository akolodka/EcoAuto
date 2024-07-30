VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelWorksheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Excel.Worksheet")
'@PredeclaredId
Option Explicit

Implements IExcelWorksheet
Implements IClonableExcelWorksheet
Implements IEditableExcelWorksheet
Implements ICellFindableExcelWorksheet
Implements IComparisonDataExcelWorksheet
Implements IChartableExcelWorksheet

Private Type TExcelWorksheet
    Worksheet As Excel.Worksheet
End Type

Private this As TExcelWorksheet

Public Function Create( _
       ByVal Worksheet As Excel.Worksheet) As IExcelWorksheet
    
    With New ExcelWorksheet
    
        Set .Worksheet = Worksheet
        Set Create = .Self
    
    End With
    
End Function

Public Property Get Worksheet() As Excel.Worksheet
    Set Worksheet = this.Worksheet
End Property
Public Property Set Worksheet(ByVal RHS As Excel.Worksheet)
    Set this.Worksheet = RHS
End Property

Public Property Get Self() As IExcelWorksheet
    Set Self = Me
End Property

Public Function GetChart() As IExcelChart
    
    If (Worksheet.ChartObjects.Count = vbEmpty) Then _
        Exit Function
    
    Dim ChartObject As Excel.ChartObject
    Set ChartObject = Worksheet.ChartObjects(1)
    
    Dim Result As IExcelChart
    Set Result = ExcelChart.Create(ChartObject)
    
    Set GetChart = Result

End Function

Public Function GetCell( _
       ByVal Row As Long, _
       ByVal Column As Long) As IExcelRange
    
    If (VarType(Worksheet.Cells(Row, Column)) = vbError) Then _
        Exit Function
    
    Dim Cell As Excel.Range
    Set Cell = Worksheet.Cells(Row, Column)
    
    Dim Result As IExcelRange
    Set Result = ExcelRange.Create(Cell)
    
    Set GetCell = Result

End Function
Public Sub CopyBefore(ByVal Sheet As ExcelWorksheet)
    Worksheet.Copy Sheet.Worksheet
End Sub
Public Sub CloneRowDown( _
       ByVal RowIndex As Long)
    
    Worksheet.Rows(RowIndex).Copy
    Worksheet.Rows(RowIndex).Insert Shift:=xlDown
    
End Sub
Public Sub DeleteRow( _
       ByVal RowIndex As Long)
       
    Worksheet.Rows(RowIndex).Delete
    
End Sub

Public Property Get Name() As String
    Name = Worksheet.Name
End Property

Public Property Let Name(ByVal RHS As String)
    Worksheet.Name = RHS
End Property
Public Function GetUsedRange() As IExcelRange
    
    Dim WorkRange As Excel.Range
    Set WorkRange = Worksheet.Range(Worksheet.UsedRange.Address)

    Dim Result As IExcelRange
    Set Result = ExcelRange.Create(WorkRange)
    
    Set GetUsedRange = Result
    
End Function
'---------------------------------------------
Private Function IChartableExcelWorksheet_GetChart() As IExcelChart
    Set IChartableExcelWorksheet_GetChart = GetChart()
End Function
'---------------------------------------------
Private Sub IClonableExcelWorksheet_CopyBefore(ByVal Sheet As IExcelWorksheet)
    CopyBefore Sheet
End Sub
'---------------------------------------------
Private Sub IEditableExcelWorksheet_CloneRowDown(ByVal RowIndex As Long)
    CloneRowDown RowIndex
End Sub
Private Sub IEditableExcelWorksheet_DeleteRow(ByVal RowIndex As Long)
    DeleteRow RowIndex
End Sub
Private Property Get IEditableExcelWorksheet_Cells(ByVal Row As Long, ByVal Column As Long) As IExcelRange
    Set IEditableExcelWorksheet_Cells = GetCell(Row, Column)
End Property

Private Property Get IEditableExcelWorksheet_Name() As String
    IEditableExcelWorksheet_Name = Name
End Property
Private Property Let IEditableExcelWorksheet_Name(ByVal RHS As String)
    Name = RHS
End Property
'---------------------------------------------
Private Function ICellFindableExcelWorksheet_GetCell(ByVal Row As Long, ByVal Column As Long) As IExcelRange
    Set ICellFindableExcelWorksheet_GetCell = GetCell(Row, Column)
End Function
'---------------------------------------------
Private Function IComparisonDataExcelWorksheet_GetUsedRange() As IExcelRange
    Set IComparisonDataExcelWorksheet_GetUsedRange = GetUsedRange
End Function
