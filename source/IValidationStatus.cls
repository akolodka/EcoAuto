VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IValidationStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Initialization.Validator.Statuses")
'@Interface
Option Explicit

Public Property Get Message() As String
End Property

Public Property Get Color() As Long
End Property

Public Property Get IsTransferAvailable() As Boolean
End Property

