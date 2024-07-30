VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICacheContentEditService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Cache.Content Edit"
'@Interface
Option Explicit

Public Function GetFromContent(ByVal Key As String) As String
End Function

Public Function InsertIntoContent(ByVal Key As String, ByVal Value As String) As IFileContent
End Function

Public Property Get Content() As IFileContent
End Property

