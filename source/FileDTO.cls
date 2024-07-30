VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileDTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Initialization.File System Object.File")
'@PredeclaredId
Option Explicit

Implements IFile

Private Type TFileObject

    Path As String
    BaseName As String
    
End Type

Private this As TFileObject

Public Function Create( _
       ByVal Path As String, _
       ByVal BaseName As String) As IFile
    
    With New FileDTO
    
        .Path = Path
        .BaseName = BaseName
        
        Set Create = .Self
        
    End With
    
End Function

Public Property Get Path() As String
    Path = this.Path
End Property
Public Property Let Path(ByVal RHS As String)
    this.Path = RHS
End Property

Public Property Get BaseName() As String
    BaseName = this.BaseName
End Property
Public Property Let BaseName(ByVal RHS As String)
    this.BaseName = RHS
End Property

Public Property Get Self() As IFile
    Set Self = Me
End Property

'---------------------------------------------
Private Property Get IFile_Path() As String
    IFile_Path = Path
End Property

Private Property Get IFile_NameWithoutExtension() As String
    IFile_NameWithoutExtension = BaseName
End Property


