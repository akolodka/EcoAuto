VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FolderCleaner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Preparation Document Templates.Folder Cleaner"
'@PredeclaredId
Option Explicit
Implements IFolderCleaner

Private Type TFolderCleaner

    Fso As IFileSystemObject
    Progress As IProgressService
    
End Type

Private this As TFolderCleaner

Public Function Create( _
       ByVal Fso As IFileSystemObject, _
       ByVal Progress As IProgressService) As IFolderCleaner
    
    With New FolderCleaner
        
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

Public Property Get Self() As IFolderCleaner
    Set Self = Me
End Property

Private Sub CleanFolder( _
        ByVal folderPath As String)
    
    Progress.Show _
        "Очистка каталога", _
        "Каталог: " & Fso.GetBaseName(folderPath)
    
    If (folderPath = vbNullString) Then _
       Exit Sub
        
    If (folderPath = Application.PathSeparator) Then _
       Exit Sub
    
    If (Fso.FolderExists(folderPath) = False) Then
    
        Fso.CreateFolder folderPath
        Exit Sub
    
    End If
    
    Fso.DeletePath folderPath
    
    Fso.CreateFolder folderPath
    
End Sub

'---------------------------------------------
Private Sub IFolderCleaner_CleanFolder( _
        ByVal folderPath As String)
    
    CleanFolder folderPath
    
End Sub


