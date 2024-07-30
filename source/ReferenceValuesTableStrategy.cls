VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ReferenceValuesTableStrategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Comparison Results Decoration.Tables Strategies"
Option Explicit

Implements IComparisonDecorationStrategy
Implements IComparisonTemplateFileStrategy
Implements IComparisonProgressNotifyStrate
Implements ITableRowAppendStrategy
Implements ITableDecorationStrategy

Const FileMask As String = "referenceValuesTable"
Const ProgressHeaderText As String = "Таблица приписанных значений"

Private Sub AppendRow( _
        ByVal Table As IAppendableWordTable, _
        ByVal Model As IReferenceTableRowModel)
    
    If (Table Is Nothing = True) Then _
       Exit Sub
    
    If (Model Is Nothing = True) Then _
       Exit Sub
    
    If (Table.IsLastRowFilled = True) Then _
       Table.AddRow
    
    Table.AppendValue 1, Model.Name
    Table.AppendValue 2, Model.Subname
    Table.AppendValue 3, Model.Value
    Table.AppendValue 4, Model.Uncertainty
    
End Sub

Private Sub ApplyFinalDecoration( _
        ByVal Table As IFormattableWordTable)
    
    If (Table Is Nothing = True) Then _
       Exit Sub
    
    Table.UnderlineHeader
    Table.MergeEqualCellsByRows 1
    
End Sub
'---------------------------------------------
Private Property Get IComparisonTemplateFileStrategy_TemplateFileMask() As String
    IComparisonTemplateFileStrategy_TemplateFileMask = FileMask
End Property
Private Property Get IComparisonProgressNotifyStrate_ProgressFooter() As String
    IComparisonProgressNotifyStrate_ProgressFooter = ProgressHeaderText
End Property
Private Sub ITableRowAppendStrategy_AppendRow(ByVal Table As IAppendableWordTable, ByVal Model As ITableRowModel, Optional ByVal isBold As Boolean = False)
    AppendRow Table, Model
End Sub
Private Sub ITableDecorationStrategy_ApplyTableDecoration(ByVal Table As IAppendableWordTable)
    ApplyFinalDecoration Table
End Sub


