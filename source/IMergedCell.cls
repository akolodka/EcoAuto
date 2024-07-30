VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IMergedCell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Word.Table.Merged Cell")
'@Interface
Option Explicit

Public Property Get StartRow() As Long
End Property

Public Property Get EndRow() As Long
End Property

Public Property Let EndRow(ByVal RHS As Long)
End Property

Public Property Get Value() As String
End Property

