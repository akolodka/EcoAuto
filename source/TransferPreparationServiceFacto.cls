VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TransferPreparationServiceFacto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Preparation Document Templates"
'@PredeclaredId
Option Explicit
Implements ITransferPreparationServiceFact

Private Type TTransferPreparationFactory

    Closer As IDocumentCloser
    Cleaner As IFolderCleaner
    Mover As ITemplatesMover
    
    Fso As IFileSystemObject
    Config As ITransferPreparationConfig
    
End Type

Private this As TTransferPreparationFactory

Public Function Create( _
       ByVal Initial As ITransferPreparationInitializat) As ITransferPreparationServiceFact

    With New TransferPreparationServiceFacto
        
        Set .Closer = Initial.Closer
        Set .Cleaner = Initial.Cleaner
        
        Set .Mover = Initial.Mover

        Set .Fso = Initial.Fso
        Set .Config = Initial.Config
        
        Set Create = .Self
        
    End With
    
End Function

Public Property Get Closer() As IDocumentCloser
    Set Closer = this.Closer
End Property
Public Property Set Closer(ByVal RHS As IDocumentCloser)
    Set this.Closer = RHS
End Property

Public Property Get Cleaner() As IFolderCleaner
    Set Cleaner = this.Cleaner
End Property
Public Property Set Cleaner(ByVal RHS As IFolderCleaner)
    Set this.Cleaner = RHS
End Property

Public Property Get Mover() As ITemplatesMover
    Set Mover = this.Mover
End Property
Public Property Set Mover(ByVal RHS As ITemplatesMover)
    Set this.Mover = RHS
End Property

Public Property Get Fso() As IFileSystemObject
    Set Fso = this.Fso
End Property
Public Property Set Fso(ByVal RHS As IFileSystemObject)
    Set this.Fso = RHS
End Property

Public Property Get Config() As ITransferPreparationConfig
    Set Config = this.Config
End Property
Public Property Set Config(ByVal RHS As ITransferPreparationConfig)
    Set this.Config = RHS
End Property

Public Property Get Self() As ITransferPreparationServiceFact
    Set Self = Me
End Property
'---------------------------------------------
Private Function ITransferPreparationServiceFact_Create( _
        ByVal Model As IExcelTransferModel) As ITransferPreparationService
    
    With New TransferPreparationService
    
        Set .Participants = Model.Participants

        Set .Closer = Closer
        Set .Cleaner = Cleaner
        Set .Mover = Mover
        
        Set .Fso = Fso
        Set .Config = Config
        
        Set ITransferPreparationServiceFact_Create = .Self
        
    End With
    
End Function


