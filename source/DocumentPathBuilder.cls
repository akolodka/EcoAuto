VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DocumentPathBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Word.Path Builder"
'@PredeclaredId
Option Explicit

Implements IDocumentPathBuilderFactory
Implements IDocumentPathBuilder

Const SuffixDelimiter As String = " -- "

Private Type TDocumentPathBuilder

    Fso As IFileSystemObject
    Progress As IProgressService
    
    DocumentName As String
    ResultFolderPath As String
    
    SuffixText As String
    
End Type

Private this As TDocumentPathBuilder

Public Function Create( _
       ByVal FileSystemObject As IFileSystemObject, _
       ByVal Progress As IProgressService, _
       ByVal ResultFolderPath As String) As IDocumentPathBuilderFactory
    
    With New DocumentPathBuilder
        
        Set .Fso = FileSystemObject
        Set .Progress = Progress
        
        .ResultFolderPath = ResultFolderPath
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

Public Property Get ResultFolderPath() As String
    ResultFolderPath = this.ResultFolderPath
End Property
Public Property Let ResultFolderPath(ByVal RHS As String)
    this.ResultFolderPath = RHS
End Property

Public Property Get Self() As IDocumentPathBuilderFactory
    Set Self = Me
End Property

'---------------------------------------------
Private Function IDocumentPathBuilderFactory_Create( _
        ByVal Name As String, _
        ByVal Suffix As String) As IDocumentPathBuilder
    
    DocumentName = Name
    SuffixText = Suffix
    
    Set IDocumentPathBuilderFactory_Create = Self

End Function

Public Property Get DocumentName() As String
    DocumentName = this.DocumentName
End Property
Public Property Let DocumentName(ByVal RHS As String)
    this.DocumentName = RHS
End Property

Public Property Get SuffixText() As String
    SuffixText = this.SuffixText
End Property
Public Property Let SuffixText(ByVal RHS As String)
    this.SuffixText = RHS
End Property

Public Function GetSaveName() As String

    Progress.Show "Сохранение результатов участника"
    
    Dim fileNewName As String
    fileNewName = GetFileNewSafeName()
    
    Dim isFileNameFree As Boolean
    isFileNameFree = (Fso.FileExists(fileNewName) = False)
    
    Dim Index As Long
    Do While (isFileNameFree = False)
        
        Index = Index + 1
        fileNewName = GetFileNewSafeName(Index)
        
        isFileNameFree = (Fso.FileExists(fileNewName) = False)
        
    Loop
    
    Dim Result As String
    Result = Fso.BuildPath(ResultFolderPath, fileNewName)
    
    GetSaveName = Result
    
End Function

Private Function GetFileNewSafeName( _
        Optional ByVal uniqueIndex As Long) As String
        
    Dim nameBase As String
    nameBase = Fso.GetBaseName(DocumentName)
        
    Dim constNamePart As String
    constNamePart = nameBase & SuffixDelimiter & SuffixText
        
    Dim fileExt As String
    fileExt = Fso.GetExtensionName(DocumentName)
        
    Dim unsafeResult As String
    unsafeResult = IIf(uniqueIndex > vbEmpty, _
                       constNamePart & " (" & uniqueIndex & ")." & fileExt, _
                       constNamePart & "." & fileExt)
        
    Dim safeResult As String
    safeResult = GetSafeFileName(unsafeResult)
        
    GetFileNewSafeName = safeResult
        
End Function

Private Function GetSafeFileName( _
        ByVal unsafeName As String) As String
            
    Dim forbiddenCharacters() As String
    forbiddenCharacters = Split( _
                          Expression:="\ : ? > "" / * < |", _
                          Delimiter:=" ")
            
    Dim Result As String
    Result = unsafeName
            
    Const SAFE_CHARACTER As String = "_"
            
    Dim i As Long
    For i = LBound(forbiddenCharacters) To UBound(forbiddenCharacters)
            
        Result = Replace(Result, forbiddenCharacters(i), SAFE_CHARACTER)
                
    Next i
            
    GetSafeFileName = Result
            
End Function

'---------------------------------------------
Private Function IDocumentPathBuilder_GetSaveName() As String
    IDocumentPathBuilder_GetSaveName = GetSaveName()
End Function


