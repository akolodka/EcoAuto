VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IFormattableWordTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Word.Table.Contracts")
'@Interface
Option Explicit

Public Sub ApplyFont(ByVal FontName As String)
End Sub

Public Sub UnderlineHeader()
End Sub

Public Sub MergeEqualCellsByRows(ByVal TargetColumn As Long)
End Sub

Public Sub SetRowsHeight(ByVal HeightInCentimeters As Double)
End Sub
