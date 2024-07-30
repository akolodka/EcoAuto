VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileSystemObjectService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Initialization.File System Object")
Option Explicit
Implements IFileSystemObject

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)
Const SleepInterval As Long = 750

Private Type TFileSystemObjectCustom
    Fso As FileSystemObject
End Type

Private this As TFileSystemObjectCustom
Private Sub Class_Initialize()
    Set Fso = New FileSystemObject
End Sub

Public Property Get Fso() As FileSystemObject
    Set Fso = this.Fso
End Property
Public Property Set Fso(ByVal RHS As FileSystemObject)
    Set this.Fso = RHS
End Property

'---------------------------------------------
Private Function IFileSystemObject_FileExists( _
        ByVal Path As String) As Boolean
    
    IFileSystemObject_FileExists = Fso.FileExists(Path)
    
End Function

Private Function IFileSystemObject_GetBaseName( _
        ByVal Path As String) As String
    
    IFileSystemObject_GetBaseName = Fso.GetBaseName(Path)
    
End Function

Private Function IFileSystemObject_FolderExists( _
        ByVal Path As String) As Boolean
    
    IFileSystemObject_FolderExists = Fso.FolderExists(Path)
    
End Function

Private Function IFileSystemObject_CreateFolder( _
        ByVal Path As String) As Folder
    
    Set IFileSystemObject_CreateFolder = Fso.CreateFolder(Path)
    
End Function

Private Function IFileSystemObject_BuildPath( _
        ByVal Path As String, _
        ByVal Name As String) As String
    
    IFileSystemObject_BuildPath = Fso.BuildPath(Path, Name)
    
End Function

Private Sub IFileSystemObject_CopyFile( _
        ByVal Source As String, _
        ByVal Destination As String, _
        Optional ByVal OverWriteFiles As Boolean = True)
    
    Fso.CopyFile Source, Destination, OverWriteFiles
    
End Sub

Private Sub IFileSystemObject_DeletePath( _
        ByVal Path As String, _
        Optional ByVal Force As Boolean = False)
    
    Shell "cmd /c rd /S/Q """ & Path & """"
    Sleep SleepInterval
    
End Sub

Private Function IFileSystemObject_GetFolderFiles( _
        ByVal Path As String) As List_IFile_
    
    Set IFileSystemObject_GetFolderFiles = New List_IFile_
    
    If (Fso.FolderExists(Path) = False) Then _
       Exit Function
    
    Dim folderFiles As Files
    Set folderFiles = Fso.GetFolder(Path).Files
    
    If (folderFiles.Count = vbEmpty) Then _
       Exit Function
        
    Dim Result As List_IFile_
    Set Result = New List_IFile_
    
    Dim File As File
    For Each File In folderFiles
    
        If (IsWordTemporatyFile(File.Name) = False) Then
                
            Result.Add FileDTO.Create( _
                       File.Path, Fso.GetBaseName(File.Name))
                    
        End If
        
    Next
    
    Set IFileSystemObject_GetFolderFiles = Result
    
End Function

Private Function IsWordTemporatyFile( _
        ByVal FileName As String) As Boolean

    IsWordTemporatyFile = (FileName Like "~$*")
        
End Function

Private Function IFileSystemObject_GetExtensionName( _
        ByVal Path As String) As String
    
    IFileSystemObject_GetExtensionName = Fso.GetExtensionName(Path)
    
End Function

Private Sub IFileSystemObject_OpenByShell( _
        ByVal Path As String, _
        Optional ByVal WindowStyle As VbAppWinStyle = 1&)
    
    Shell "explorer.exe " & Path, WindowStyle

End Sub

Private Function IFileSystemObject_ReadFile( _
        ByVal Path As String) As IFileContent
    
    If (Fso.FileExists(Path) = False) Then _
    
        Set IFileSystemObject_ReadFile = FileContent.Create(vbNullString)
        Exit Function
        
    End If
    
    Dim Content As String
    Content = Fso.OpenTextFile(Path) _
              .ReadAll
     
    Dim Result As IFileContent
    Set Result = FileContent.Create(Content)
    
    Set IFileSystemObject_ReadFile = Result
    
End Function

Private Sub IFileSystemObject_CopyFolder( _
        ByVal Source As String, _
        ByVal Destination As String)
    
    Fso.CopyFolder Source, Destination

End Sub

Private Sub IFileSystemObject_WriteFile( _
        ByVal Path As String, _
        ByVal Content As IFileContent)
    
    Dim createIfNotExists As Boolean
    createIfNotExists = True
    
    Fso.OpenTextFile( _
        Path, ForWriting, createIfNotExists) _
        .Write _
        Content.Text

End Sub
