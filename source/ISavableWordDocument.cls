VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISavableWordDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Word.Document")
'@Interface
Option Explicit

Public Sub Save()
End Sub

Public Sub CloseWithoutSave()
End Sub

Public Sub SaveAs(ByVal FilePath As String)
End Sub
