VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CacheIOService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Cache.IO")
'@PredeclaredId
Option Explicit
Implements ICacheIOService

Const CacheFileName As String = "cache.txt"

Private Type TCacheIOService

    Fso As IFileSystemObject
    Path As String
    
End Type

Private this As TCacheIOService

Public Function Create( _
       ByVal FileSystemObject As IFileSystemObject) As ICacheIOService
    
    With New CacheIOService
    
        Set .Fso = FileSystemObject
        Set Create = .Self
        
    End With
    
End Function

Public Property Get Fso() As IFileSystemObject
    Set Fso = this.Fso
End Property
Public Property Set Fso(ByVal RHS As IFileSystemObject)
    Set this.Fso = RHS
End Property

Public Property Get Path() As String
    Path = this.Path
End Property
Public Property Let Path(ByVal RHS As String)
    this.Path = RHS
End Property

Public Property Get Self() As ICacheIOService
    
    Path = Fso.BuildPath(ThisWorkbook.Path, CacheFileName)
    Set Self = Me
    
End Property

'---------------------------------------------
Private Function ICacheIOService_ReadContent() As IFileContent
    
    If (Fso Is Nothing = True) Then _
       Exit Function

    Dim readResult As IFileContent
    Set readResult = Fso.ReadFile(Path)
    
    Set ICacheIOService_ReadContent = readResult
    
End Function

Private Sub ICacheIOService_WriteContent( _
        ByVal Content As IFileContent)
    
    Fso.WriteFile Path, Content
    
End Sub


