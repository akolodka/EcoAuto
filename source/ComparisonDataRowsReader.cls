VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComparisonDataRowsReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Initialization.Comparison Mapping.Data Rows Reader"
'@PredeclaredId
Option Explicit
Implements IComparisonDataRowsReader

Private Type TComparisonDataRowsMapper
    UsedRange As IComparisonDataExcelRange
End Type

Private this As TComparisonDataRowsMapper

Public Function Create( _
       ByVal SourceSheet As IComparisonDataExcelWorksheet) As IComparisonDataRowsReader

    With New ComparisonDataRowsReader
            
        If (SourceSheet Is Nothing = False) Then _
            Set .UsedRange = SourceSheet.GetUsedRange()
            
        Set Create = .Self
    
    End With
    
End Function

Public Property Get UsedRange() As IComparisonDataExcelRange
    Set UsedRange = this.UsedRange
End Property
Public Property Set UsedRange(ByVal RHS As IComparisonDataExcelRange)
    Set this.UsedRange = RHS
End Property

Public Property Get Self() As IComparisonDataRowsReader
    Set Self = Me
End Property

Public Function GetKeysRow() As IExcelRange
    
    Const KeysRowIndex As Long = 1
    
    Set GetKeysRow = GetDataRow(KeysRowIndex)
    
End Function
    Public Function GetDataRow( _
           ByVal Index As Long) As IExcelRange
        
        If (UsedRange Is Nothing = True) Then _
           Exit Function
        
        Dim Result As IExcelRange
        Set Result = UsedRange.GetRow(Index)
        
        Set GetDataRow = Result
        
    End Function

Public Function GetComparisonDataRows() As List_IExcelRange_
    
    Dim Result As List_IExcelRange_
    Set Result = New List_IExcelRange_
    
    Const InitialDataRowIndex As Long = 4
    
    Dim i As Long
    For i = InitialDataRowIndex To UsedRange.RowsCount
        
        Dim Model As IExcelRange
        Set Model = UsedRange.GetRow(i)
        
        Result.Add Model
        
    Next
    
    Set GetComparisonDataRows = Result
    
End Function
'---------------------------------------------
Private Function IComparisonDataRowsReader_GetKeysRow() As IExcelRange
    Set IComparisonDataRowsReader_GetKeysRow = GetKeysRow
End Function
Private Function IComparisonDataRowsReader_GetComparisonDataRows() As List_IExcelRange_
    Set IComparisonDataRowsReader_GetComparisonDataRows = GetComparisonDataRows
End Function


