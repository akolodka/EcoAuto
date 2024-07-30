VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITransferDialogInitialization"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Initialization.Contracts")
'@Interface
Option Explicit

Public Property Get SubcontractFiles() As List_IFile_
End Property

Public Property Get Participants() As List_IParticipant_
End Property

Public Property Get SubcontractHandler() As ISubcontractHandler
End Property

Public Property Get ValidationStatus() As IValidationStatus
End Property

Public Property Get TransferService() As ITransferService
End Property

Public Property Get FileSystemObject() As IFileSystemObject
End Property

Public Property Get ControlObjectHandler() As IControlObjectHandler
End Property

