VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IKeyValuePair"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Document IO.KeyValue Pair"
'@Interface
Option Explicit

Public Property Get Key() As ITransferKey
End Property

Public Property Get Value() As ITransferValue
Attribute Value.VB_UserMemId = 0
End Property
