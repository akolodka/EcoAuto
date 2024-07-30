VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SubcontractValuesHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Initialization.Subcontract")
'@PredeclaredId
Option Explicit
Implements ISubcontractHandler

Const DefaultValueFileName As String = "_default"

Private Type TSubcontractValuesHandler

    SubcontractFiles As List_IFile_
    ReaderFactory As IDocumentReaderFactory
    
End Type

Private this As TSubcontractValuesHandler

Public Function Create( _
       ByVal SubcontractFiles As List_IFile_, _
       ByVal ReaderFactory As IDocumentReaderFactory) As ISubcontractHandler
    
    With New SubcontractValuesHandler
        
        Set .SubcontractFiles = SubcontractFiles
        Set .ReaderFactory = ReaderFactory

        Set Create = .Self
        
    End With

End Function

Public Property Get SubcontractFiles() As List_IFile_
    Set SubcontractFiles = this.SubcontractFiles
End Property
Public Property Set SubcontractFiles(ByVal RHS As List_IFile_)
    Set this.SubcontractFiles = RHS
End Property

Public Property Get ReaderFactory() As IDocumentReaderFactory
    Set ReaderFactory = this.ReaderFactory
End Property
Public Property Set ReaderFactory(ByVal RHS As IDocumentReaderFactory)
    Set this.ReaderFactory = RHS
End Property

Public Property Get Self() As ISubcontractHandler
    Set Self = Me
End Property

Public Function GetDefaultValue() As String
    GetDefaultValue = GetSubcontractValueFromName(DefaultValueFileName)
End Function

Private Function GetSubcontractValueFromName( _
        ByVal Name As String) As String
        
    Dim File As IFile
    Set File = GetSubcontractFileFromName(Name)
        
    If (File Is Nothing = True) Then _
       Exit Function
        
    Dim Reader As IDocumentReader
    Set Reader = ReaderFactory.Create(File.Path)
        
    Dim Result As ITransferValue
    Set Result = Reader.ReadContent
        
    'На случай, если в static files почему-то окажется *docx
    If IsObject(Result.Value) = True Then _
       Exit Function
                
    GetSubcontractValueFromName = Result.Value
    
End Function

Private Function GetSubcontractFileFromName( _
        ByVal Name As String) As IFile
            
    Dim i As Long
    For i = 1 To SubcontractFiles.Count
                
        If SubcontractFiles(i).NameWithoutExtension = Name Then
                
            Set GetSubcontractFileFromName = SubcontractFiles(i)
            Exit Function
                    
        End If
                
    Next
    
End Function

Public Function GetValueFromSeveralNames( _
       ByVal Names As List_String_) As String
    
    Dim Files As List_IFile_
    Set Files = GetApprovedSubcontractFileList(Names)
    
    If Files.Count = vbEmpty Then
    
        GetValueFromSeveralNames = GetDefaultValue
        Exit Function
        
    End If
    
    Dim Result As String
    
    Dim i As Long
    For i = 1 To Files.Count

        Result = GetResult(Result, Files(i))
        
    Next
    
    GetValueFromSeveralNames = Result

End Function

Private Function GetApprovedSubcontractFileList( _
        ByVal Names As List_String_) As List_IFile_
        
    Dim Result As List_IFile_
    Set Result = New List_IFile_
    
    Dim i As Long
    For i = 1 To Names.Count
        
        Dim File As IFile
        Set File = GetSubcontractFileFromName(Names(i))
            
        If (File Is Nothing = False) Then _
           Result.Add File
            
    Next
        
    Set GetApprovedSubcontractFileList = Result
        
End Function

Private Function GetResult( _
        ByVal Source As String, _
        ByVal File As IFile) As String
        
    Dim Result As String
        
    If Source <> vbNullString Then _
       Result = Source & vbNewLine
            
    Result = Result & GetSubcontractValueFromName( _
             File.NameWithoutExtension)
            
    GetResult = Result
        
End Function

'---------------------------------------------
Private Function ISubcontractHandler_GetValue( _
        ByVal Selected As List_String_) As String
    
    If SubcontractFiles.Count = vbEmpty Then _
       Exit Function
    
    If Selected.Count = vbEmpty Then
        ISubcontractHandler_GetValue = GetDefaultValue
    Else
        ISubcontractHandler_GetValue = GetValueFromSeveralNames(Selected)
    End If
    
End Function


