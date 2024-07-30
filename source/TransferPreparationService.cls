VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TransferPreparationService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Preparation Document Templates"
Option Explicit
Implements ITransferPreparationService

Private Type TTransferPreparationService

    Closer As IDocumentCloser
    Cleaner As IFolderCleaner
    
    Participants As List_IParticipant_
    Mover As ITemplatesMover
    
    Fso As IFileSystemObject
    Config As ITransferPreparationConfig

End Type

Private this As TTransferPreparationService

Public Property Get Self() As ITransferPreparationService
    Set Self = Me
End Property

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

Public Property Get Participants() As List_IParticipant_
    Set Participants = this.Participants
End Property
Public Property Set Participants(ByVal RHS As List_IParticipant_)
    Set this.Participants = RHS
End Property

Private Sub Execute()

    Closer.CloseOpenedWordDocuments Config.ResultFolderPath
    Closer.CloseOpenedWordDocuments Config.TemporaryFolderPath
    
    Cleaner.CleanFolder Config.ResultFolderPath
    Cleaner.CleanFolder Config.TemporaryFolderPath
    
    Mover.CopyTemplatesFolder Config.DocumentTemplatesFolderPath, Config.TemporaryFolderPath
    
    MoveComparisonChartsTemplate
    MoveAuxillaryTemplates

End Sub
    Private Sub MoveComparisonChartsTemplate()
        
        Dim WorkbookName As String
        WorkbookName = Fso.GetBaseName(Config.ComparisonChartsTemplateWorkbookPath) & "." & Fso.GetExtensionName(Config.ComparisonChartsTemplateWorkbookPath)
        
        Dim DestinationPath As String
        DestinationPath = Fso.BuildPath(Config.TemporaryFolderPath, WorkbookName)
        
        Fso.CopyFile Config.ComparisonChartsTemplateWorkbookPath, DestinationPath
        
    End Sub
    Private Sub MoveAuxillaryTemplates()
                
        Dim i As Long
        For i = 1 To Participants.Count
                
            Dim ParticipantNumber As String
            ParticipantNumber = GetNumber(Participants(i))
                
            Dim DestinationPath As String
            DestinationPath = Fso.BuildPath(Config.TemporaryFolderPath, ParticipantNumber)
                    
            Mover.CopyTemplatesFolder _
                Config.AuxillaryTemplatesFolderPath, DestinationPath
                    
        Next
                
    End Sub
        Private Function GetNumber( _
                ByVal Model As IComparisonDataParticipant) As String
                
            GetNumber = Model.Number.Value
                    
        End Function

Private Sub Dispose()
    Fso.DeletePath Config.TemporaryFolderPath
End Sub
'---------------------------------------------
Private Sub ITransferPreparationService_PrepareTemplates()
    Execute
End Sub
Private Sub ITransferPreparationService_Dispose()
    Dispose
End Sub


