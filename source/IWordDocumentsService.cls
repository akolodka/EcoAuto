VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IWordDocumentsService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Word"
'@Interface
Option Explicit

Public Function GetTemporaryDocument() As IWordDocument
End Function

Public Function GetFrom(ByVal Path As String) As IWordDocument
End Function

Public Sub Dispose()
End Sub

