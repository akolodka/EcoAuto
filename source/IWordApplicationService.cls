VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IWordApplicationService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Word.Application.Contracts")
'@Interface
Option Explicit

Public Function DocumentsAdd(Optional ByVal Visible As Boolean = False) As IWordDocument
End Function

Public Function DocumentsOpen(ByVal FilePath As String, Optional ByVal Visible As Boolean = False) As IWordDocument
End Function

Public Sub Quit()
End Sub

Public Property Get IsWordApplicationExists() As Boolean
End Property

Public Function GetApplicationDocuments() As List_IWordDocument_
End Function

