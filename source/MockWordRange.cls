VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MockWordRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "Tests.Mocks.Word.Range"
'@PredeclaredId
Option Explicit

Implements IWordRange
Implements IFindableWordRange
Implements IEditableWordRange
Implements IUpdatableWordRange

Private Type TMockWordRange
    
    FindResultCount As Long
    FindedText As String

    DocumentRange As String
    
End Type

Private this As TMockWordRange

Public Function Create() As IWordRange
    
    With New MockWordRange
     
        .DocumentRange = "MockWordRange.Value"
        Set Create = .Self
        
    End With
    
End Function

Public Property Get Self() As IWordRange
    Set Self = Me
End Property

Public Property Get DocumentRange() As String
    DocumentRange = this.DocumentRange
End Property
Public Property Let DocumentRange(ByVal RHS As String)
    this.DocumentRange = RHS
End Property

Public Property Get FindResultCount() As Long
    FindResultCount = this.FindResultCount
End Property
Public Property Let FindResultCount(ByVal RHS As Long)
    this.FindResultCount = RHS
End Property

Public Property Get FindedText() As String
    FindedText = this.FindedText
End Property
Public Property Let FindedText(ByVal RHS As String)
    this.FindedText = RHS
End Property

Public Function FindExecute( _
        Optional ByVal FindText As String) As Boolean
    
    If FindedText <> FindText Then
        
        FindResultCount = vbEmpty
        FindedText = FindText
        
    End If
    
    FindResultCount = FindResultCount + 1
    FindExecute = FindResultCount < 3
    
Debug.Print
Debug.Print "Mock WordRange.FindExecute"
Debug.Print "    FindText: "; FindText
Debug.Print "    Result: "; FindExecute
Debug.Print "    Count: "; FindResultCount
    
End Function

Public Sub ClearContent()

Debug.Print
Debug.Print "Mock WordRange.ClearContent"

End Sub

Public Property Get LastTable() As IFormattableWordTable
    Set LastTable = MockWordTable.Create()
End Property

Private Sub PasteAndFormat()
    
Debug.Print
Debug.Print "Mock WordRange.PasteAndFormat"
    
End Sub

Public Sub PasteAsText( _
        ByVal Text As String)
    
Debug.Print
Debug.Print "Mock WordRange.PasteAsText"
Debug.Print "    text: "; Text

End Sub

Public Sub Copy()
    
Debug.Print
Debug.Print "Mock WordRange.Copy"
    
End Sub

Public Sub UpdateFields()
    
Debug.Print
Debug.Print "Mock WordRange.Fields.Update"
    
End Sub

Public Property Get Tables() As List_IWordTable_
    
    Dim Result As List_IWordTable_
    Set Result = New List_IWordTable_
    
    Result.Add MockWordTable.Create()
    
    Set Tables = Result
    
End Property
'---------------------------------------------
Private Property Get IWordRange_LastTable() As IFormattableWordTable
    Set IWordRange_LastTable = LastTable
End Property
Private Property Get IWordRange_Tables() As List_IWordTable_
    Set IWordRange_Tables = Tables
End Property
'---------------------------------------------
Private Function IFindableWordRange_FindExecute(Optional ByVal FindText As String) As Boolean
    IFindableWordRange_FindExecute = FindExecute(FindText)
End Function
'---------------------------------------------
Private Sub IEditableWordRange_Copy()
    Copy
End Sub
Private Sub IEditableWordRange_PasteAndFormat()
    PasteAndFormat
End Sub
Private Sub IEditableWordRange_PasteAsText(ByVal Text As String)
    PasteAsText Text
End Sub
Private Sub IEditableWordRange_ClearContent()
    ClearContent
End Sub
'---------------------------------------------
Private Sub IUpdatableWordRange_UpdateFields()
    UpdateFields
End Sub
