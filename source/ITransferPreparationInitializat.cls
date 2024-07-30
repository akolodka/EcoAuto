VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITransferPreparationInitializat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Initialization.Contracts")
'@Interface
Option Explicit

Public Property Get Config() As ITransferPreparationConfig
End Property

Public Property Get Fso() As IFileSystemObject
End Property

Public Property Get Closer() As IDocumentCloser
End Property

Public Property Get Cleaner() As IFolderCleaner
End Property

Public Property Get Mover() As ITemplatesMover
End Property

