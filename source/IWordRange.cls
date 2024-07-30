VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IWordRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Word.Range")
'@Interface
Option Explicit

Public Property Get Tables() As List_IWordTable_
End Property

Public Property Get LastTable() As IFormattableWordTable
End Property




