VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MockFileSystemObjectService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "Tests.Mocks.Initialization.File System Object"
'@PredeclaredId
Option Explicit
Implements IFileSystemObject

Public Function Create() As IFileSystemObject
    
    With New MockFileSystemObjectService
        Set Create = .Self
    End With
    
End Function

Public Property Get Self() As IFileSystemObject
    Set Self = Me
End Property

'---------------------------------------------
Private Function IFileSystemObject_FileExists( _
        ByVal Path As String) As Boolean
    
    Dim Result As Boolean
    Result = True
    
    If InStr(Path, Application.PathSeparator) = vbEmpty Then _
        Result = False
    
Debug.Print
Debug.Print "Mock fso.FileExists" & vbNewLine & _
            "   path: "; Path & vbNewLine & _
            "   return: "; Result
    
    IFileSystemObject_FileExists = Result
    
End Function

Private Function IFileSystemObject_GetBaseName( _
        ByVal Path As String) As String
    
    Dim Fso As New FileSystemObject
    
    Dim Result As String
    Result = Fso.GetBaseName(Path)
    
Debug.Print
Debug.Print "Mock fso.GetBaseName" & vbNewLine & _
            "   path: "; Path & vbNewLine & _
            "   return: "; Result
    
    IFileSystemObject_GetBaseName = Result
    
End Function

Private Function IFileSystemObject_FolderExists( _
        ByVal Path As String) As Boolean
    
    Static Result As Boolean
    Result = Not Result
    
Debug.Print
Debug.Print "Mock fso.FolderExists" & vbNewLine & _
            "   Path: "; Path & vbNewLine & _
            "   return: "; Result

    IFileSystemObject_FolderExists = Result
    
End Function

Private Function IFileSystemObject_CreateFolder( _
        ByVal Path As String) As Folder
    
Debug.Print
Debug.Print "Mock fso.CreateFolder"
Debug.Print "    Path: "; Path
    
End Function

Private Function IFileSystemObject_BuildPath( _
        ByVal Path As String, _
        ByVal Name As String) As String
    
    Dim Result As String
    Result = Path & Application.PathSeparator & Name
    
Debug.Print
Debug.Print "Mock fso.BuildPath" & vbNewLine & _
            "   return: "; Result

    IFileSystemObject_BuildPath = Result
    
End Function

Private Sub IFileSystemObject_CopyFile( _
        ByVal Source As String, _
        ByVal Destination As String, _
        Optional ByVal OverWriteFiles As Boolean = True)
    
Debug.Print
Debug.Print "Mock fso.CopyFile" & vbNewLine & _
            "   Source: "; Source; "" & vbNewLine & _
            "   Destination: "; Destination


End Sub

Private Sub IFileSystemObject_DeletePath( _
        ByVal FileSpec As String, _
        Optional ByVal Force As Boolean = False)
    
Debug.Print
Debug.Print "Mock fso.DeletePath" & vbNewLine & _
            "   path: "; FileSpec

End Sub

Private Function IFileSystemObject_GetFolderFiles( _
        ByVal Path As String) As List_IFile_
    
    Select Case True
        
        Case InStr(LCase(Path), "static") > vbEmpty
            Set IFileSystemObject_GetFolderFiles = MockFileList.CreateStaticFiles()
        
        Case InStr(LCase(Path), "subcontract") > vbEmpty
            Set IFileSystemObject_GetFolderFiles = MockFileList.CreateSubcontractFiles()
        
        Case InStr(LCase(Path), "auxillary") > vbEmpty
            Set IFileSystemObject_GetFolderFiles = MockFileList.CreateAuxillaryComparisonFiles
        
        Case InStr(LCase(Path), "templates") > vbEmpty
            Set IFileSystemObject_GetFolderFiles = MockFileList.CreateDocumentTeplateFiles
        
        Case InStr(LCase(Path), "temporary") > vbEmpty
            Set IFileSystemObject_GetFolderFiles = MockFileList.CreateAuxillaryComparisonFiles
        
        Case Else
            Set IFileSystemObject_GetFolderFiles = GetDefaultMockFileList
            
    End Select
    
End Function
    Private Function GetDefaultMockFileList()
        
        Dim Result As List_IFile_
        Set Result = New List_IFile_
        
        Result.Add FileDTO.Create(" ...\mockWordRangeFile.docx", "mockWordRangeFile")
        Result.Add FileDTO.Create(" ...\mockPlainTextFile.txt", "mockPlainTextFile")
        Result.Add FileDTO.Create(" ...\_default.txt", "_default")
        
        Set GetDefaultMockFileList = Result
    
    End Function

Private Function IFileSystemObject_GetExtensionName( _
        ByVal Path As String) As String
    
    Dim Fso As New FileSystemObject
    IFileSystemObject_GetExtensionName = Fso.GetExtensionName(Path)
    
End Function

Private Sub IFileSystemObject_OpenByShell( _
        ByVal Path As String, _
        Optional ByVal WindowStyle As VbAppWinStyle = 1&)
    
Debug.Print
Debug.Print "Mock fso.OpenByShell" & vbNewLine & _
            "    Path: "; Path; ""

End Sub

Private Function IFileSystemObject_ReadFile( _
        ByVal Path As String) As IFileContent
    
    Dim Content As String
    Content = "Mock TextFile Content"
    
    Dim Result As IFileContent
    Set Result = FileContent.Create(Content)
    
Debug.Print
Debug.Print "Mock fso.ReadFile" & vbNewLine & _
            "    Path: "; Path; "" & vbNewLine & _
            "    return: "; Content
    
    Set IFileSystemObject_ReadFile = Result
    
End Function

Private Sub IFileSystemObject_CopyFolder( _
        ByVal Source As String, _
        ByVal Destination As String)
    
Debug.Print
Debug.Print "Mock fso.CopyFolder" & vbNewLine & _
            "    source: "; Source; "" & vbNewLine & _
            "    destination: "; Destination
    
End Sub

Private Sub IFileSystemObject_WriteFile( _
        ByVal Path As String, _
        ByVal Content As IFileContent)
    
Debug.Print
Debug.Print "Mock fso.WriteFile" & vbNewLine & _
            "    Path: "; Path; "" & vbNewLine & _
            "    Writed Content: "; Content.Text

End Sub


