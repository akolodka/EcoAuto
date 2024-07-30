VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TransferResultsPresentator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Results Presentator")
'@PredeclaredId
Option Explicit
Implements ITransferResultsPresentator

Private Type TTransferResultsPresentator

    Opener As IExplorerFolderOpener
    Config As IResultPresentationConfiguratio
    
End Type

Private this As TTransferResultsPresentator

Public Function Create( _
       ByVal Opener As IExplorerFolderOpener, _
       ByVal Config As IResultPresentationConfiguratio) As ITransferResultsPresentator

    With New TransferResultsPresentator
        
        Set .Opener = Opener
        Set .Config = Config
        
        Set Create = .Self
        
    End With

End Function

Public Property Get Opener() As IExplorerFolderOpener
    Set Opener = this.Opener
End Property
Public Property Set Opener(ByVal RHS As IExplorerFolderOpener)
    Set this.Opener = RHS
End Property

Public Property Get Config() As IResultPresentationConfiguratio
    Set Config = this.Config
End Property
Public Property Set Config(ByVal RHS As IResultPresentationConfiguratio)
    Set this.Config = RHS
End Property

Public Property Get Self() As ITransferResultsPresentator
    Set Self = Me
End Property

Public Sub OpenResultsFolder()
    Opener.ShowInExplorer Config.ResultFolderPath
End Sub
Public Sub OpenTemporaryFolder()
    Opener.ShowInExplorer Config.TemporaryFolderPath
End Sub
'---------------------------------------------
Private Sub ITransferResultsPresentator_OpenResultsFolder()
   OpenResultsFolder
End Sub
Private Sub ITransferResultsPresentator_OpenTemporaryFolder()
    OpenTemporaryFolder
End Sub

