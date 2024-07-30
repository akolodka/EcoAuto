VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IResultPresentationConfiguratio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Configuration.Contracts")
'@Interface
Option Explicit

Public Property Get TemporaryFolderPath() As String
End Property

Public Property Get ResultFolderPath() As String
End Property
