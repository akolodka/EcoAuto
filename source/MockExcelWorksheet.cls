VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MockExcelWorksheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Tests.Mocks.Excel.Worksheet")
'@PredeclaredId
Option Explicit

Implements IExcelWorksheet
Implements IClonableExcelWorksheet
Implements IEditableExcelWorksheet
Implements ICellFindableExcelWorksheet
Implements IComparisonDataExcelWorksheet
Implements IChartableExcelWorksheet

Public Function Create() As IExcelWorksheet
    
    With New MockExcelWorksheet
        Set Create = .Self
    End With
    
End Function

Public Property Get Self() As IExcelWorksheet
    Set Self = Me
End Property

Public Function GetChart() As IExcelChart
    Set GetChart = MockExcelChart.Create()
End Function

Public Function GetCell( _
       ByVal Row As Long, _
       ByVal Column As Long) As IExcelRange

Debug.Print
Debug.Print "Mock MockExcelWorksheet.GetCell" & _
            "   Row: "; Row; _
            "   Column: "; Column

    Set GetCell = MockExcelRange.Create()

End Function
Public Sub CopyBefore(ByVal Sheet As IExcelWorksheet)

Debug.Print
Debug.Print "Mock MockExcelWorksheet.CopyBefore"

End Sub
Public Sub CloneRowDown( _
       ByVal RowIndex As Long)
    
Debug.Print
Debug.Print "Mock MockExcelWorksheet.CloneRowDown"
    
End Sub
Public Sub DeleteRow( _
       ByVal RowIndex As Long)
       
Debug.Print
Debug.Print "Mock MockExcelWorksheet.DeleteRow"
    
End Sub
Public Property Get Name() As String
    Name = "Mock MockExcelWorksheet.Name"
End Property

Public Property Let Name(ByVal RHS As String)

Debug.Print
Debug.Print "Mock Set MockExcelWorksheet.Name"

End Property

Public Function GetUsedRange() As IExcelRange
    
    Dim Result As IExcelRange
    Set Result = MockExcelRange.Create()
    
    Set GetUsedRange = Result
    
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
Private Function IComparisonDataExcelWorksheet_GetUsedRange() As IExcelRange
    Set IComparisonDataExcelWorksheet_GetUsedRange = GetUsedRange
End Function
'---------------------------------------------
Private Function ICellFindableExcelWorksheet_GetCell(ByVal Row As Long, ByVal Column As Long) As IExcelRange
    Set ICellFindableExcelWorksheet_GetCell = GetCell(Row, Column)
End Function
'---------------------------------------------
Private Function IChartableExcelWorksheet_GetChart() As IExcelChart
    Set IChartableExcelWorksheet_GetChart = GetChart
End Function

