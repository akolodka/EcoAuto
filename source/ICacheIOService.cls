VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICacheIOService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Cache.IO")
'@Interface
Option Explicit

Public Function ReadContent() As IFileContent
End Function

Public Sub WriteContent(ByVal Content As IFileContent)
End Sub

