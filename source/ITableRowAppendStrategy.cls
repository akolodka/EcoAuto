VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITableRowAppendStrategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Comparison Results Decoration.Tables Strategies.Contracts"
'@Interface
Option Explicit

Public Sub AppendRow( _
       ByVal Table As IAppendableWordTable, _
       ByVal Model As ITableRowModel, _
       Optional ByVal isBold As Boolean = False)
End Sub


