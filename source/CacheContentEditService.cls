VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CacheContentEditService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Cache.Content Edit"
'@PredeclaredId
Option Explicit
Implements ICacheContentEditService

Const KeyValueDelimiter As String = "="
Const ItemDelimiter As String = vbNewLine

Private Type TCacheContentEditService
    Content As IFileContent
End Type

Private this As TCacheContentEditService

Public Function Create( _
       ByVal Content As IFileContent) As ICacheContentEditService
    
    With New CacheContentEditService
        
        Set .Content = Content
        Set Create = .Self
        
    End With
    
End Function

Public Property Get Content() As IFileContent
    Set Content = this.Content
End Property
Public Property Set Content(ByVal RHS As IFileContent)
    Set this.Content = RHS
End Property

Public Property Get Self() As ICacheContentEditService
    Set Self = Me
End Property

Private Function GetFromContent( _
        ByVal Key As String) As String
    
    If (Content.Text = vbNullString) Then _
       Exit Function
    
    Dim Start As Long
    Start = GetStartPosition(Key)
    
    If Start = vbEmpty Then _
       Exit Function

    Dim length As Long
    length = GetValueLength(Start)

    Dim Result As String
    Result = Mid$(Content.Text, Start, length)
    
    GetFromContent = Result
    
End Function

Private Function GetStartPosition( _
        ByVal Key As String) As Long
        
    Dim findKey As String
    findKey = GetFindKey(Key)

    Dim Start As Long
    Start = InStr(Content.Text, findKey)

    If Start = vbEmpty Then _
       Exit Function

    Dim Result As Long
    Result = Start + Len(findKey)

    GetStartPosition = Result

End Function

Private Function GetFindKey( _
        ByVal Key As String) As String
            
    GetFindKey = Key & KeyValueDelimiter
            
End Function

Private Function GetValueLength( _
        ByVal Start As Long) As Long

    Dim endPosition As Long
    endPosition = InStr(Start, Content.Text, ItemDelimiter)
        
    'Обязательно + 1, иначе последний символ будет обрезан
    If endPosition = vbEmpty Then _
       endPosition = Len(Content.Text) + 1
        
    Dim Result As Long
    Result = endPosition - Start

    GetValueLength = Result
        
End Function

Private Function InsertIntoContent( _
        ByVal Key As String, _
        ByVal Value As String) As IFileContent
    
    Dim Start As Long
    Start = GetStartPosition(Key)
    
    If Start = vbEmpty Then
        InsertNewValue Key, Value
    Else
        ReplaceStoredValue Key, Value
    End If
    
    Set InsertIntoContent = Content

End Function

Private Sub InsertNewValue( _
        ByVal Key As String, _
        ByVal Value As String)
        
    Dim findKey As String
    findKey = GetFindKey(Key)
        
    Dim newValueString As String
    newValueString = findKey & Value
        
    Dim newContentText As String
        
    If Content.Text = vbNullString Then
        newContentText = newValueString
    Else
        newContentText = Content.Text & ItemDelimiter & newValueString
    End If
        
    Set Content = FileContent.Create(newContentText)

End Sub

Private Sub ReplaceStoredValue( _
        ByVal Key As String, _
        ByVal Value As String)
        
    Dim findKey As String
    findKey = GetFindKey(Key)
        
    Dim storedValue As String
    storedValue = GetFromContent(Key)
        
    Dim oldValue As String
    oldValue = findKey & storedValue
        
    Dim NewValue As String
    NewValue = findKey & Value
        
    Set Content = FileContent.Create( _
                  Replace(Content.Text, oldValue, NewValue))
        
End Sub

'---------------------------------------------
Private Function ICacheContentEditService_GetFromContent( _
        ByVal Key As String) As String

    ICacheContentEditService_GetFromContent = GetFromContent(Key)
    
End Function

Private Function ICacheContentEditService_InsertIntoContent( _
        ByVal Key As String, _
        ByVal Value As String) As IFileContent

    Set ICacheContentEditService_InsertIntoContent = InsertIntoContent(Key, Value)
    
End Function

Private Property Get ICacheContentEditService_Content() As IFileContent
    Set ICacheContentEditService_Content = Content
End Property


