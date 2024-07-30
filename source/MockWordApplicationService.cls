VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MockWordApplicationService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Tests.Mocks.Word.Application"
Option Explicit
Implements IWordApplicationService
Implements IWordApplicationQuitService

Private Type TMockWordApplicationService

    Progress As IProgressService
    IsWordInitialized As Boolean
    
End Type

Private this As TMockWordApplicationService

Public Property Get Progress() As IProgressService
    Set Progress = this.Progress
End Property

Public Property Set Progress(ByVal RHS As IProgressService)
    Set this.Progress = RHS
End Property

Public Property Get IsWordInitialized() As Boolean
    IsWordInitialized = this.IsWordInitialized
End Property

Public Property Let IsWordInitialized(ByVal RHS As Boolean)
    this.IsWordInitialized = RHS
End Property

Public Property Get Self() As IWordApplicationService
    Set Self = Me
End Property

Private Sub IWordApplicationQuitService_EmergencyQuit()
   
Debug.Print
Debug.Print "+ + + + + + + + + + + + + + + + " & vbNewLine & _
            "   Mock WordApp.EmergencyQuit" & vbNewLine & _
            "+ + + + + + + + + + + + + + + + " & vbNewLine
 
End Sub

'---------------------------------------------
Private Function IWordApplicationService_DocumentsAdd( _
        Optional ByVal Visible As Boolean = False) As IWordDocument
    
    If (IsWordInitialized = False) Then _
       InitializeWordApplication
    
Debug.Print
Debug.Print "Mock WordApp.DocumentsAdd"
    
    Set IWordApplicationService_DocumentsAdd = MockWordDocument.Create()

End Function

Private Sub InitializeWordApplication()
        
    Progress.Show "Mock Подключение к Word"
    IsWordInitialized = True
        
End Sub

Private Function IWordApplicationService_DocumentsOpen( _
        ByVal FilePath As String, _
        Optional ByVal Visible As Boolean = False) As IWordDocument
    
    If (IsWordInitialized = False) Then _
       InitializeWordApplication
    
Debug.Print
Debug.Print "Mock WordApp.DocumentsOpen" & vbNewLine & _
            "    FilePath: "; FilePath; "" & vbNewLine & _
            "    Visible: "; Visible
    
    Set IWordApplicationService_DocumentsOpen = MockWordDocument.Create()
    
End Function

Private Function IWordApplicationService_GetApplicationDocuments() As List_IWordDocument_
    
    Dim Result As List_IWordDocument_
    Set Result = New List_IWordDocument_
    
    Result.Add MockWordDocument.Create
    Result.Add MockWordDocument.Create
    
    Set IWordApplicationService_GetApplicationDocuments = Result
    
End Function

Private Property Get IWordApplicationService_IsCreatedNewWord() As Boolean
    
    Dim Result As Boolean
    
Debug.Print
Debug.Print "Mock IsCreatedNewWord" & vbNewLine & _
            "    Return: "; Result
                
    IWordApplicationService_IsCreatedNewWord = Result
    
End Property

Private Property Get IWordApplicationService_IsWordApplicationExists() As Boolean
    
    Dim Result As Boolean
    
Debug.Print
Debug.Print "Mock IsWordApplicationExists" & vbNewLine & _
            "    Return: "; Result
                
    IWordApplicationService_IsWordApplicationExists = Result
    
End Property

Private Sub IWordApplicationService_Quit()
    Progress.Show "Mock WordApp.Quit"
End Sub
