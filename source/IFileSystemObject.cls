VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IFileSystemObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Initialization.File System Object")
'@Interface
Option Explicit

Public Function FileExists(ByVal Path As String) As Boolean
End Function

Public Function GetBaseName(ByVal Path As String) As String
End Function

Public Function FolderExists(ByVal Path As String) As Boolean
End Function

Public Function CreateFolder(ByVal Path As String) As Folder
End Function

Public Function BuildPath(ByVal Path As String, ByVal Name As String) As String
End Function

Public Sub CopyFile(ByVal Source As String, ByVal Destination As String, Optional ByVal OverWriteFiles As Boolean = True)
End Sub

Public Sub DeletePath(ByVal Path As String, Optional ByVal Force As Boolean = False)
End Sub

Public Function GetFolderFiles(ByVal Path As String) As List_IFile_
End Function

Public Function GetExtensionName(ByVal Path As String) As String
End Function

Public Function ReadFile(ByVal Path As String) As IFileContent
End Function

Public Sub WriteFile(ByVal Path As String, ByVal Content As IFileContent)
End Sub

Public Sub CopyFolder(ByVal Source As String, ByVal Destination As String)
End Sub

Public Sub OpenByShell(ByVal Path As String, Optional ByVal WindowStyle As VbAppWinStyle = vbNormalFocus)
End Sub

