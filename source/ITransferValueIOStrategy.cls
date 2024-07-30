VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITransferValueIOStrategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Document IO.IO Strategies.Contracts"
'@Interface
Option Explicit

Public Function IsMatchesForRead(ByVal FilePath As String) As Boolean
End Function

Public Function IsMatchesForWrite(ByVal Value As ITransferValue) As Boolean
End Function

