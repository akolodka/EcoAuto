VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PartialReadyStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Initialization.Validator.Statuses")
Option Explicit
Implements IValidationStatus

Private Property Get IValidationStatus_Message() As String
    IValidationStatus_Message = "Возможен частичный перенос."
End Property

Private Property Get IValidationStatus_Color() As Long
    IValidationStatus_Color = &HC0C0&
End Property

Private Property Get IValidationStatus_IsTransferAvailable() As Boolean
    IValidationStatus_IsTransferAvailable = True
End Property

