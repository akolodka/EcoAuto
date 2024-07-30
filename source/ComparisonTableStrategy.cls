VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComparisonTableStrategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Comparison Results Decoration.Tables Strategies"
Option Explicit

Implements ITableRowAppendStrategy
Implements ITableDecorationStrategy

Const DefaultName As String = "(скрыто)"

Private Sub AppendRow( _
        ByVal Table As IAppendableWordTable, _
        ByVal Model As IComparisonTableRowModel, _
        ByVal isBold As Boolean)
    
    If (Table Is Nothing = True) Then _
       Exit Sub
       
    If (Model Is Nothing = True) Then _
       Exit Sub
    
    If (Table.IsLastRowFilled = True) Then _
       Table.AddRow
    
    Table.AppendValue 1, Model.Number, isBold
    
    Dim Name As String
    Name = IIf(isBold = True, Model.Name, DefaultName)
    
    Table.AppendValue 2, Name, isBold
    Table.AppendValue 3, Model.Value, isBold
    Table.AppendValue 4, Model.Uncertainty, isBold
    Table.AppendValue 5, Model.EnCriteria, isBold

End Sub
Private Sub ApplyFinalDecoration( _
        ByVal Table As IFormattableWordTable)
    
    If (Table Is Nothing = True) Then _
       Exit Sub
    
    Table.UnderlineHeader
    
End Sub
'---------------------------------------------
Private Sub ITableRowAppendStrategy_AppendRow(ByVal Table As IAppendableWordTable, ByVal Model As ITableRowModel, Optional ByVal isBold As Boolean = False)
    AppendRow Table, Model, isBold
End Sub
Private Sub ITableDecorationStrategy_ApplyTableDecoration(ByVal Table As IAppendableWordTable)
    ApplyFinalDecoration Table
End Sub


