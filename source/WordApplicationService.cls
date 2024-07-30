VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WordApplicationService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Word.Application")
Option Explicit

Implements IWordApplicationService
Implements IWordApplicationQuitService

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)
Const SleepInterval As Long = 500

Private Type TWordApplicationService
    
    Progress As IProgressService
    Application As Word.Application
    
    IsCreatedNewWord As Boolean
    UntouchedDocuments As List_IWordDocument_
    
End Type

Private this As TWordApplicationService

Public Property Get Self() As IWordApplicationService
    Set Self = Me
End Property

Public Property Get Progress() As IProgressService
    Set Progress = this.Progress
End Property
Public Property Set Progress(ByVal RHS As IProgressService)
    Set this.Progress = RHS
End Property

Public Property Get Application() As Word.Application
    Set Application = this.Application
End Property
Public Property Set Application(ByVal RHS As Word.Application)
    Set this.Application = RHS
End Property

Public Function DocumentsAdd( _
       Optional ByVal Visible As Boolean = False) As IWordDocument
    
    If (IsWordInitialized = False) Then _
       InitializeWordApplication
    
    Dim Document As Word.Document
    Set Document = Application.Documents.Add(Visible:=Visible)
    
    Set DocumentsAdd = WordDocument.Create(Document)

End Function

Private Function IsWordInitialized() As Boolean
    IsWordInitialized = (Application Is Nothing = False)
End Function

Private Sub InitializeWordApplication()
    On Error GoTo CleanFail
        
    Progress.Show "Подключение к Word"
        
    If (Application Is Nothing = False) Then _
       Exit Sub
        
    Set Application = GetObject(, "Word.Application")
    Set UntouchedDocuments = GetApplicationDocuments()
        
    Exit Sub
        
CleanFail:

    this.IsCreatedNewWord = True
        
    Set Application = New Word.Application
    Set UntouchedDocuments = New List_IWordDocument_
        
End Sub

Public Property Get UntouchedDocuments() As List_IWordDocument_
    Set UntouchedDocuments = this.UntouchedDocuments
End Property
Public Property Set UntouchedDocuments(ByVal RHS As List_IWordDocument_)
    Set this.UntouchedDocuments = RHS
End Property

Public Property Get IsCreatedNewWord() As Boolean
    IsCreatedNewWord = this.IsCreatedNewWord
End Property

Public Function DocumentsOpen( _
       ByVal FilePath As String, _
       Optional ByVal Visible As Boolean = False) As IWordDocument
    
    If (IsWordInitialized = False) Then _
       InitializeWordApplication
    
    Dim Document As Word.Document
    Set Document = Application.Documents.Open( _
                   FileName:=FilePath, Visible:=Visible)
    
    Set DocumentsOpen = WordDocument.Create(Document)
    
End Function

Public Sub Quit()
    
    If Application Is Nothing Then _
       Exit Sub
       
    Progress.Show "Очистка памяти"

    If this.IsCreatedNewWord = True Then
    
        this.Application.Quit SaveChanges:=False
        Sleep SleepInterval
        
        Exit Sub
        
    End If

    CloseProgramRaisedDocuments
    
End Sub

Public Sub EmergencyQuit()
    
    If Application Is Nothing Then _
       Exit Sub
       
    this.Application.Quit SaveChanges:=False
    
End Sub

Private Sub CloseProgramRaisedDocuments()
    
    Dim EndDocuments As List_IWordDocument_
    Set EndDocuments = GetApplicationDocuments()

    Dim i As Long
    For i = 1 To EndDocuments.Count
     
        If (IsDocumentProgramRaised(EndDocuments(i)) = True) Then _
           CloseDocument EndDocuments(i)
            
    Next
        
    Set EndDocuments = GetApplicationDocuments()
        
End Sub
    Private Function IsDocumentProgramRaised( _
            ByVal Document As IWordDocument) As Boolean
                
        Dim i As Long
        For i = 1 To UntouchedDocuments.Count
                
            If Document.Name = UntouchedDocuments(i).Name Then _
               Exit Function
                    
        Next
                
        IsDocumentProgramRaised = True
    
    End Function
    Private Sub CloseDocument( _
        ByVal Document As ISavableWordDocument)
    
        Document.CloseWithoutSave
        
    End Sub
Public Function IsWordExists() As Boolean
    On Error GoTo CleanFail
    
    Dim wdApp As Word.Application
    Set wdApp = GetObject(, "Word.Application")
    
    IsWordExists = True

CleanFail:
    'Здесь ничего не должно быть

End Function

Public Function GetApplicationDocuments() As List_IWordDocument_
    
    If (IsWordInitialized = False) Then _
       InitializeWordApplication
    
    Dim Result As List_IWordDocument_
    Set Result = New List_IWordDocument_
        
    'Обязательно без ForEach, иначе возникают дубли в добавляемых документах
    Dim i As Long
    For i = 1 To Application.Documents.Count
        
        Dim doc As Word.Document
        Set doc = Application.Documents(i)
        
        Result.Add WordDocument.Create(doc)

    Next

    Set GetApplicationDocuments = Result
    
End Function
'---------------------------------------------
Private Function IWordApplicationService_DocumentsAdd( _
        Optional ByVal Visible As Boolean = False) As IWordDocument
        
    Set IWordApplicationService_DocumentsAdd = DocumentsAdd(Visible)

End Function
Private Function IWordApplicationService_DocumentsOpen( _
        ByVal FilePath As String, _
        Optional ByVal Visible As Boolean = False) As IWordDocument
    
    Set IWordApplicationService_DocumentsOpen = DocumentsOpen(FilePath, Visible)
    
End Function
Private Property Get IWordApplicationService_IsCreatedNewWord() As Boolean
    IWordApplicationService_IsCreatedNewWord = IsCreatedNewWord
End Property
Private Sub IWordApplicationService_Quit()
    Quit
End Sub
Private Property Get IWordApplicationService_IsWordApplicationExists() As Boolean
    IWordApplicationService_IsWordApplicationExists = IsWordExists
End Property
Private Function IWordApplicationService_GetApplicationDocuments() As List_IWordDocument_
    Set IWordApplicationService_GetApplicationDocuments = GetApplicationDocuments()
End Function
Private Sub IWordApplicationQuitService_EmergencyQuit()
    EmergencyQuit
End Sub
