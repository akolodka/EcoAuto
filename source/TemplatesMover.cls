VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TemplatesMover"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Preparation Document Templates.Templates Mover"
'@PredeclaredId
Option Explicit
Implements ITemplatesMover

Private Type TTemplatesMover

    Fso As IFileSystemObject
    Progress As IProgressService
    
End Type

Private this As TTemplatesMover

Public Function Create( _
       ByVal Fso As IFileSystemObject, _
       ByVal Progress As IProgressService) As ITemplatesMover
    
    With New TemplatesMover
        
        Set .Fso = Fso
        Set .Progress = Progress
        
        Set Create = .Self
    
    End With

End Function

Public Property Get Fso() As IFileSystemObject
    Set Fso = this.Fso
End Property
Public Property Set Fso(ByVal RHS As IFileSystemObject)
    Set this.Fso = RHS
End Property

Public Property Get Progress() As IProgressService
    Set Progress = this.Progress
End Property
Public Property Set Progress(ByVal RHS As IProgressService)
    Set this.Progress = RHS
End Property

Public Property Get Self() As ITemplatesMover
    Set Self = Me
End Property

Private Sub MoveTemplates( _
        ByVal SourceFolderPath As String, _
        ByVal DestinationFolderPath As String)
    
    Progress.Show _
        "Перенос шаблонов"
    
    If (Fso.FolderExists(SourceFolderPath) = False) Then _
       Exit Sub
    
    Fso.CopyFolder SourceFolderPath, DestinationFolderPath
  
End Sub

'---------------------------------------------
Private Sub ITemplatesMover_CopyTemplatesFolder( _
        ByVal SourceFolderPath As String, _
        ByVal DestiantionFolderPath As String)
    
    MoveTemplates SourceFolderPath, DestiantionFolderPath

End Sub


