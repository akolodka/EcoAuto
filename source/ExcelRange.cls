VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Excel.Range")
'@PredeclaredId
Option Explicit

Implements IExcelRange
Implements IMarkableExcelRange
Implements IComparisonDataExcelRange
Implements IComparisonDataRow

Private Type TExcelRange
    Range As Excel.Range
End Type

Private this As TExcelRange

Public Function Create( _
       ByVal Range As Excel.Range) As IExcelRange
       
    With New ExcelRange
    
        Set .Range = Range
        Set Create = .Self
        
    End With

End Function

Public Property Get Range() As Excel.Range
    Set Range = this.Range
End Property
Public Property Set Range(ByVal RHS As Excel.Range)
    Set this.Range = RHS
End Property
Public Sub Calculate()
    Range.Calculate
End Sub

Public Property Get Self() As IExcelRange
    Set Self = Me
End Property

Public Sub Mark()
    Range.Interior.Color = vbYellow
End Sub

Public Function GetRow( _
       ByVal Index As Long) As IExcelRange
    
    If (Index > RowsCount) Then _
        Exit Function
    
    Dim Result As IExcelRange
    Set Result = ExcelRange.Create(Range.Rows(Index))
    
    Set GetRow = Result
    
End Function

Public Property Get RowsCount() As Long
    RowsCount = Range.Rows.Count
End Property

Public Property Get ColumnsCount() As Long
    ColumnsCount = Range.Columns.Count
End Property


Public Function GetValue( _
       ByVal Column As Long) As String
    
    If (Range Is Nothing = True) Then _
       Exit Function

    Dim Result As String
    Result = "—"
    
    If (IsError(Range.Value2(1, Column)) = False) Then _
        Result = Range.Value2(1, Column)

    GetValue = Result
    
End Function
'---------------------------------------------
Private Property Get IExcelRange_Value() As String
    IExcelRange_Value = Range.Value2
End Property
Private Property Let IExcelRange_Value(ByVal RHS As String)
    
    If IsNumeric(RHS) Then
        Range.Value2 = CDbl(RHS)
    Else
        Range.Value2 = RHS
    End If
    
End Property

Private Sub IExcelRange_Calculate()
    Calculate
End Sub
'---------------------------------------------
Private Sub IMarkableExcelRange_Mark()
    Mark
End Sub
'---------------------------------------------
Private Function IComparisonDataExcelRange_GetRow(ByVal Index As Long) As IExcelRange
    Set IComparisonDataExcelRange_GetRow = GetRow(Index)
End Function
Private Property Get IComparisonDataExcelRange_RowsCount() As Long
    IComparisonDataExcelRange_RowsCount = RowsCount
End Property
'---------------------------------------------
Private Function IComparisonDataRow_GetValue(ByVal Column As Long) As String
    IComparisonDataRow_GetValue = GetValue(Column)
End Function
Private Property Get IComparisonDataRow_ColumnsCount() As Long
    IComparisonDataRow_ColumnsCount = ColumnsCount
End Property

