VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DocumentCloser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Preparation Document Templates.Document Closer"
'@PredeclaredId
Option Explicit
Implements IDocumentCloser

Private Type TDocumentCloser

    WordApp As IWordApplicationService
    Fso As IFileSystemObject
    Progress As IProgressService
    
End Type

Private this As TDocumentCloser

Public Function Create( _
       ByVal WordApp As IWordApplicationService, _
       ByVal Fso As IFileSystemObject, _
       ByVal Progress As IProgressService) As IDocumentCloser
    
    With New DocumentCloser
        
        Set .WordApp = WordApp
        Set .Fso = Fso
        Set .Progress = Progress
        
        Set Create = .Self
    
    End With

End Function

Public Property Get WordApp() As IWordApplicationService
    Set WordApp = this.WordApp
End Property
Public Property Set WordApp(ByVal RHS As IWordApplicationService)
    Set this.WordApp = RHS
End Property

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

Public Property Get Self() As IDocumentCloser
    Set Self = Me
End Property

Private Sub CloseOpenedWordDocuments( _
        ByVal folderPath As String)
    
    Progress.Show _
        "Выгрузка открытых документов", _
        "Каталог: " & Fso.GetBaseName(folderPath)
    
    If (WordApp Is Nothing = True) Then _
       Exit Sub

    Dim isWordOpen As Boolean
    isWordOpen = WordApp.IsWordApplicationExists
    
    If (isWordOpen = False) Then _
       Exit Sub
    
    Dim OpenedDocuments As List_IWordDocument_
    Set OpenedDocuments = WordApp.GetApplicationDocuments()
    
    Dim i As Long
    For i = 1 To OpenedDocuments.Count
        
        Dim doc As IWordDocument
        Set doc = OpenedDocuments(i)
        
        If (IsDocumentOpen(doc.Path, folderPath) = True) Then _
           CloseDocument doc
        
    Next

End Sub
    Private Sub CloseDocument( _
        ByVal Document As ISavableWordDocument)
        
        Document.CloseWithoutSave

    End Sub

    Private Function IsDocumentOpen( _
            ByVal DocPath As String, _
            ByVal folderPath As String) As Boolean
            
        IsDocumentOpen = (InStr(DocPath, folderPath) > vbEmpty)
        
    End Function

'---------------------------------------------
Private Sub IDocumentCloser_CloseOpenedWordDocuments( _
        ByVal folderPath As String)
    
    CloseOpenedWordDocuments folderPath
    
End Sub


